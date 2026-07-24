//! Comments support (python-docx 1.2 comments API).

use crate::docmodel::{with_core, PyDocument};
use crate::docmodel_add::mutate_document;
use crate::pyclasses::PyDocxTemplate;
use crate::template::{TplCore, DOCUMENT_PART};
use crate::xmldom::{Document, Element, Node};
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

const COMMENTS_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

/// ISO 8601 UTC timestamp (unix -> civil date).
pub fn now_iso8601() -> String {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    unix_to_iso(secs)
}

fn unix_to_iso(secs: i64) -> String {
    let days = secs.div_euclid(86400);
    let tod = secs.rem_euclid(86400);
    let (h, m, s) = (tod / 3600, (tod % 3600) / 60, tod % 60);
    // civil from days (Howard Hinnant's algorithm)
    let z = days + 719468;
    let era = z.div_euclid(146097);
    let doe = z.rem_euclid(146097);
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let mo = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if mo <= 2 { y + 1 } else { y };
    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        y, mo, d, h, m, s
    )
}

/// Ensure the comments part exists; return max comment id currently used.
fn ensure_comments_part(core: &mut TplCore) -> Result<i64, String> {
    core.init_docx(false)?;
    {
        let pkg = core.package.as_mut().ok_or("package not loaded")?;
        if !pkg.contains("word/comments.xml") {
            let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:comments>";
            pkg.set("word/comments.xml", xml.as_bytes().to_vec());
            pkg.ensure_content_type_override("word/comments.xml", COMMENTS_CT);
            if pkg.rels(DOCUMENT_PART).by_type(crate::package::rel_type::COMMENTS).next().is_none() {
                pkg.add_rel(DOCUMENT_PART, crate::package::rel_type::COMMENTS, "comments.xml", false);
            }
        }
    }
    let pkg = core.package.as_ref().ok_or("package not loaded")?;
    let xml = pkg.get_string("word/comments.xml").unwrap_or_default();
    let mut max_id: i64 = -1;
    let re = fancy_regex::Regex::new(r#"<w:comment [^>]*w:id="(\d+)""#).unwrap();
    for cap in re.captures_iter(&xml).flatten() {
        if let Ok(n) = cap[1].parse::<i64>() {
            max_id = max_id.max(n);
        }
    }
    Ok(max_id)
}

/// Append a comment entry to the comments part, returns its id.
fn append_comment(
    core: &mut TplCore,
    text: &str,
    author: &str,
    initials: &str,
) -> Result<i64, String> {
    let id = ensure_comments_part(core)? + 1;
    let date = now_iso8601();
    let mut comment = String::from("<w:comment");
    comment.push_str(&format!(
        " w:id=\"{}\" w:author=\"{}\" w:initials=\"{}\" w:date=\"{}\">",
        id,
        crate::package::escape_xml_attr(author),
        crate::package::escape_xml_attr(initials),
        date
    ));
    for para in text.split('\n') {
        comment.push_str(&format!(
            "<w:p><w:pPr><w:pStyle w:val=\"CommentText\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
            crate::richtext::html_escape(para)
        ));
    }
    comment.push_str("</w:comment>");

    let pkg = core.package.as_mut().ok_or("package not loaded")?;
    let mut xml = pkg.get_string("word/comments.xml").unwrap_or_default();
    if let Some(pos) = xml.rfind("</w:comments>") {
        xml.insert_str(pos, &comment);
    } else {
        return Err("invalid comments part".into());
    }
    pkg.set("word/comments.xml", xml.into_bytes());
    Ok(id)
}

/// Document.add_comment: anchor a comment to the given run(s).
pub fn doc_add_comment(
    doc: &PyDocument,
    py: Python<'_>,
    runs: &Bound<'_, PyAny>,
    text: &str,
    author: &str,
    initials: &str,
) -> PyResult<PyComment> {
    // normalize runs into (first, last) (para, index) pairs
    let mut run_refs: Vec<(usize, usize)> = Vec::new();
    if let Ok(r) = runs.extract::<pyo3::PyRef<'_, crate::docmodel::PyRun>>() {
        run_refs.push((r.para, r.index));
    } else if let Ok(seq) = runs.try_iter() {
        for item in seq {
            let item = item?;
            let r = item
                .extract::<pyo3::PyRef<'_, crate::docmodel::PyRun>>()
                .map_err(|_| PyValueError::new_err("runs must be Run objects"))?;
            run_refs.push((r.para, r.index));
        }
    }
    if run_refs.is_empty() {
        return Err(PyValueError::new_err("runs must be a Run or a non-empty sequence of Runs"));
    }
    let first = run_refs[0];
    let last = run_refs[run_refs.len() - 1];

    let comment_id = with_core(&doc.tpl, py, |core| {
        append_comment(core, text, author, initials)
    })
    .map_err(py_err)?;

    with_core(&doc.tpl, py, |core| {
        mutate_document(core, |body| {
            use crate::docmodel_add::nth_direct;
            let id_str = comment_id.to_string();
            // commentRangeStart before the first run
            if let Some(p) = nth_direct(body, "w:p", first.0) {
                let pos = p
                    .children
                    .iter()
                    .position(|c| matches!(c, Node::Elem(e) if e.name == "w:r"))
                    .map(|i| {
                        // position of the first.1-th run
                        p.children
                            .iter()
                            .enumerate()
                            .filter(|(_, c)| matches!(c, Node::Elem(e) if e.name == "w:r"))
                            .nth(first.1)
                            .map(|(i, _)| i)
                            .unwrap_or(i)
                    })
                    .unwrap_or(0);
                let mut start = Element::new("w:commentRangeStart");
                start.set_attr("w:id", &id_str);
                p.children.insert(pos.min(p.children.len()), Node::Elem(start));
            }
            // commentRangeEnd + reference run after the last run
            if let Some(p) = nth_direct(body, "w:p", last.0) {
                let pos = p
                    .children
                    .iter()
                    .enumerate()
                    .filter(|(_, c)| matches!(c, Node::Elem(e) if e.name == "w:r"))
                    .nth(last.1)
                    .map(|(i, _)| i + 1)
                    .unwrap_or(p.children.len());
                let mut end = Element::new("w:commentRangeEnd");
                end.set_attr("w:id", &id_str);
                let mut rpr = Element::new("w:rPr");
                let mut rs = Element::new("w:rStyle");
                rs.set_attr("w:val", "CommentReference");
                rpr.children.push(Node::Elem(rs));
                let mut refr = Element::new("w:commentReference");
                refr.set_attr("w:id", &id_str);
                let mut run = Element::new("w:r");
                run.children.push(Node::Elem(rpr));
                run.children.push(Node::Elem(refr));
                let at = (pos + 1).min(p.children.len());
                p.children.insert(pos.min(p.children.len()), Node::Elem(end));
                p.children.insert(at, Node::Elem(run));
            }
        })
    })
    .map_err(py_err)?;

    Ok(PyComment {
        tpl: doc.tpl.clone_ref(py),
        comment_id,
    })
}

/// A single comment.
#[pyclass(name = "Comment", unsendable)]
pub struct PyComment {
    pub tpl: Py<PyDocxTemplate>,
    pub comment_id: i64,
}

impl PyComment {
    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            core.init_docx(false).ok()?;
            let xml = core.package.as_ref()?.get_string("word/comments.xml")?;
            let dom = Document::parse(&xml).ok()?;
            let mut comments: Vec<&Element> = Vec::new();
            dom.root.iter_descendants("w:comment", &mut comments);
            comments
                .into_iter()
                .find(|c| {
                    c.get_attr("w:id")
                        .and_then(|v| v.parse::<i64>().ok())
                        == Some(self.comment_id)
                })
                .map(|c| f(c))
        })
    }
}

#[pymethods]
impl PyComment {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        self.read(py, |c| crate::docmodel::element_text(c))
            .unwrap_or_default()
    }
    #[getter]
    fn author(&self, py: Python<'_>) -> String {
        self.read(py, |c| c.get_attr("w:author").unwrap_or("").to_string())
            .unwrap_or_default()
    }
    #[getter]
    fn initials(&self, py: Python<'_>) -> String {
        self.read(py, |c| c.get_attr("w:initials").unwrap_or("").to_string())
            .unwrap_or_default()
    }
    #[getter]
    fn timestamp(&self, py: Python<'_>) -> String {
        self.read(py, |c| c.get_attr("w:date").unwrap_or("").to_string())
            .unwrap_or_default()
    }
    #[getter]
    fn comment_id(&self) -> i64 {
        self.comment_id
    }
}

/// The comments collection.
#[pyclass(name = "Comments", unsendable)]
pub struct PyComments {
    pub tpl: Py<PyDocxTemplate>,
}

#[pymethods]
impl PyComments {
    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let comments = self.comment_list(py);
        let list = pyo3::types::PyList::new(py, comments)?;
        Ok(list.call_method0("__iter__")?.unbind())
    }

    fn comment_list(&self, py: Python<'_>) -> Vec<PyComment> {
        with_core(&self.tpl, py, |core| {
            core.init_docx(false).ok();
            let xml = core
                .package
                .as_ref()
                .and_then(|p| p.get_string("word/comments.xml"));
            xml.and_then(|x| Document::parse(&x).ok())
                .map(|dom| {
                    let mut out = Vec::new();
                    let mut comments: Vec<&Element> = Vec::new();
                    dom.root.iter_descendants("w:comment", &mut comments);
                    for c in comments {
                        if let Some(id) = c.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()) {
                            out.push(PyComment {
                                tpl: self.tpl.clone_ref(py),
                                comment_id: id,
                            });
                        }
                    }
                    out
                })
                .unwrap_or_default()
        })
    }

    fn __len__(&self, py: Python<'_>) -> usize {
        self.comment_list(py).len()
    }

    /// Add an (unanchored) comment (python-docx comments.add_comment).
    #[pyo3(signature = (text="", author="", initials=""))]
    fn add_comment(&self, py: Python<'_>, text: &str, author: &str, initials: &str) -> PyResult<PyComment> {
        let id = with_core(&self.tpl, py, |core| append_comment(core, text, author, initials))
            .map_err(py_err)?;
        Ok(PyComment {
            tpl: self.tpl.clone_ref(py),
            comment_id: id,
        })
    }
}
