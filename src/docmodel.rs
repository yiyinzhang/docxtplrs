//! A python-docx-inspired facade over the current document:
//! live, writable proxies (Paragraph/Run/Table/Cell), sections, styles,
//! inline shapes and core properties.

use crate::docmodel_add::{mutate_document, nth_direct};
use crate::pyclasses::PyDocxTemplate;
use crate::template::{TplCore, DOCUMENT_PART};
use crate::xmldom::{Document, Element, Node};
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

// ---------------- parsing helpers ----------------

pub(crate) fn element_text(el: &Element) -> String {
    let mut s = String::new();
    collect_wt(el, &mut s);
    s
}

fn collect_wt(el: &Element, out: &mut String) {
    for c in &el.children {
        match c {
            Node::Elem(e) => {
                if e.name == "w:t" {
                    e.push_text_content(out);
                } else if e.name == "w:tab" {
                    out.push('\t');
                } else if e.name == "w:br" {
                    out.push('\n');
                } else {
                    collect_wt(e, out);
                }
            }
            Node::Text(_) => {}
        }
    }
}

pub(crate) fn read_body<R>(core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
    let dom = core.document_dom().ok()?;
    let body = dom.root.find("w:body")?;
    Some(f(body))
}

pub(crate) fn with_core<R>(tpl: &Py<PyDocxTemplate>, py: Python<'_>, f: impl FnOnce(&mut TplCore) -> R) -> R {
    f(&mut tpl.bind(py).borrow().core.borrow_mut())
}

fn count_in_body(core: &mut TplCore, name: &str) -> usize {
    read_body(core, |b| {
        b.children
            .iter()
            .filter(|c| matches!(c, Node::Elem(e) if e.name == name))
            .count()
    })
    .unwrap_or(0)
}

// ---------------- proxies ----------------

/// A paragraph in the document (live proxy).
#[pyclass(name = "Paragraph", unsendable)]
pub struct PyParagraph {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

impl PyParagraph {
    pub(crate) fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            let gen = core.doc_gen;
            let mut cur = core.para_cursor;
            mutate_document(core, |body| {
                if let Some(p) = nth_cursor_mut(body, "w:p", self.index, &mut cur, gen) {
                    result = Some(f(p));
                }
            })
            .map_err(py_err)?;
            core.para_cursor = cur;
            result.ok_or_else(|| PyValueError::new_err("paragraph not found"))
        })
    }

    pub(crate) fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            let gen = core.doc_gen;
            let mut cur = core.para_cursor;
            let r = read_body(core, |body| {
                nth_cursor_ref(body, "w:p", self.index, &mut cur, gen).map(|p| f(p))
            })
            .flatten();
            core.para_cursor = cur;
            r
        })
    }
}

#[pymethods]
impl PyParagraph {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        self.read(py, |p| element_text(p)).unwrap_or_default()
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |p| {
            // python-docx: clear content, add a single run
            p.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:r" || e.name == "w:hyperlink"));
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v));
            r.children.push(Node::Elem(t));
            p.children.push(Node::Elem(r));
        })
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| {
            p.find("w:pPr")
                .and_then(|ppr| ppr.find("w:pStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let sid = with_core(&self.tpl, py, |core| {
            crate::subdocbuilder::resolve_style_id(core, &v)
        });
        self.edit(py, |p| {
            let has_ppr = p.children.iter().any(|c| matches!(c, Node::Elem(e) if e.name == "w:pPr"));
            if !has_ppr {
                p.children.insert(0, Node::Elem(Element::new("w:pPr")));
            }
            let ppr = p.find_mut("w:pPr").unwrap();
            if let Some(ps) = ppr.find_mut("w:pStyle") {
                ps.set_attr("w:val", &sid);
            } else {
                let mut ps = Element::new("w:pStyle");
                ps.set_attr("w:val", &sid);
                ppr.children.insert(0, Node::Elem(ps));
            }
        })
    }


    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    #[getter]
    fn runs(&self, py: Python<'_>) -> Vec<PyRun> {
        let n = self
            .read(py, |p| {
                p.children
                    .iter()
                    .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:r"))
                    .count()
            })
            .unwrap_or(0);
        (0..n)
            .map(|i| PyRun {
                tpl: self.tpl.clone_ref(py),
                para: self.index,
                index: i,
            })
            .collect()
    }

    #[pyo3(signature = (text=""))]
    fn add_run(&self, py: Python<'_>, text: &str) -> PyResult<PyRun> {
        let n = self.edit(py, |p| {
            let n = p
                .children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:r"))
                .count();
            let mut r = Element::new("w:r");
            if !text.is_empty() {
                let mut t = Element::new("w:t");
                t.set_attr("xml:space", "preserve");
                t.children.push(Node::Text(text.to_string()));
                r.children.push(Node::Elem(t));
            }
            p.children.push(Node::Elem(r));
            n
        })?;
        Ok(PyRun {
            tpl: self.tpl.clone_ref(py),
            para: self.index,
            index: n,
        })
    }

    /// Paragraph formatting (python-docx paragraph_format).
    #[getter]
    fn paragraph_format(&self, py: Python<'_>) -> crate::docmodel_fmt::PyParagraphFormat {
        crate::docmodel_fmt::PyParagraphFormat {
            tpl: self.tpl.clone_ref(py),
            target: crate::docmodel_fmt::PfTarget::Para { index: self.index },
        }
    }

    /// Alignment shortcut (WD_ALIGN_PARAGRAPH int; xml name also accepted).
    #[getter]
    fn alignment(&self, py: Python<'_>) -> Option<i64> {
        self.paragraph_format(py).alignment(py)
    }
    #[setter]
    fn set_alignment(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        self.paragraph_format(py).set_alignment(py, v)
    }

    /// Remove all content, keeping the paragraph properties (python-docx clear).
    fn clear(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |p| {
            p.children
                .retain(|c| matches!(c, Node::Elem(e) if e.name == "w:pPr"));
        })
    }

    /// Field codes in this paragraph (w:fldSimple and complex fields).
    #[getter]
    fn fields(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyField> {
        let n = self
            .read(py, |p| crate::docmodel_fmt::field_spans(p).len())
            .unwrap_or(0);
        (0..n)
            .map(|i| crate::docmodel_fmt::PyField {
                tpl: self.tpl.clone_ref(py),
                para: self.index,
                index: i,
            })
            .collect()
    }

    /// Append a complex field (begin/instrText/separate/cached/end) to this
    /// paragraph. `instr` e.g. "PAGE" or 'TOC \\o "1-3"'; `cached` is the
    /// placeholder result shown until the field is updated (set
    /// `settings.update_fields_on_open = True` to refresh on open).
    #[pyo3(signature = (instr, cached=""))]
    fn add_field(&self, py: Python<'_>, instr: &str, cached: &str) -> PyResult<crate::docmodel_fmt::PyField> {
        let instr = instr.trim().to_string();
        let cached = cached.to_string();
        let index = self.edit(py, |p| {
            let index = crate::docmodel_fmt::field_spans(p).len();
            let mk = |child: Element| {
                let mut r = Element::new("w:r");
                r.children.push(Node::Elem(child));
                Node::Elem(r)
            };
            let mut begin = Element::new("w:fldChar");
            begin.set_attr("w:fldCharType", "begin");
            p.children.push(mk(begin));
            let mut it = Element::new("w:instrText");
            it.set_attr("xml:space", "preserve");
            it.children.push(Node::Text(format!(" {} ", instr)));
            p.children.push(mk(it));
            let mut sep = Element::new("w:fldChar");
            sep.set_attr("w:fldCharType", "separate");
            p.children.push(mk(sep));
            if !cached.is_empty() {
                let mut t = Element::new("w:t");
                t.set_attr("xml:space", "preserve");
                t.children.push(Node::Text(cached));
                p.children.push(mk(t));
            }
            let mut end = Element::new("w:fldChar");
            end.set_attr("w:fldCharType", "end");
            p.children.push(mk(end));
            index
        })?;
        Ok(crate::docmodel_fmt::PyField {
            tpl: self.tpl.clone_ref(py),
            para: self.index,
            index,
        })
    }

    /// Hyperlinks in this paragraph (read-only proxies).
    #[getter]
    fn hyperlinks(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyHyperlink> {
        let n = self
            .read(py, |p| {
                p.children
                    .iter()
                    .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:hyperlink"))
                    .count()
            })
            .unwrap_or(0);
        (0..n)
            .map(|i| crate::docmodel_fmt::PyHyperlink {
                tpl: self.tpl.clone_ref(py),
                para: self.index,
                index: i,
            })
            .collect()
    }

    /// True when a rendered page break occurs in this paragraph.
    #[getter]
    fn contains_page_break(&self, py: Python<'_>) -> bool {
        self.read(py, |p| {
            let mut out = Vec::new();
            p.iter_descendants("w:lastRenderedPageBreak", &mut out);
            !out.is_empty()
        })
        .unwrap_or(false)
    }

    /// Rendered page breaks (w:lastRenderedPageBreak markers written by Word
    /// at save time) in this paragraph.
    #[getter]
    fn rendered_page_breaks(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyRenderedPageBreak> {
        let n = self
            .read(py, |p| {
                let mut out = Vec::new();
                p.iter_descendants("w:lastRenderedPageBreak", &mut out);
                out.len()
            })
            .unwrap_or(0);
        (0..n)
            .map(|_| crate::docmodel_fmt::PyRenderedPageBreak {})
            .collect()
    }

    /// Runs and hyperlinks of this paragraph in document order.
    fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let kinds: Vec<bool> = self
            .read(py, |p| {
                p.children
                    .iter()
                    .filter_map(|c| match c {
                        Node::Elem(e) if e.name == "w:r" => Some(true),
                        Node::Elem(e) if e.name == "w:hyperlink" => Some(false),
                        _ => None,
                    })
                    .collect()
            })
            .unwrap_or_default();
        let mut ri = 0usize;
        let mut hi = 0usize;
        let mut out = Vec::new();
        for is_run in kinds {
            if is_run {
                if let Ok(v) = Py::new(
                    py,
                    PyRun {
                        tpl: self.tpl.clone_ref(py),
                        para: self.index,
                        index: ri,
                    },
                ) {
                    out.push(v.into_any());
                }
                ri += 1;
            } else {
                if let Ok(v) = Py::new(
                    py,
                    crate::docmodel_fmt::PyHyperlink {
                        tpl: self.tpl.clone_ref(py),
                        para: self.index,
                        index: hi,
                    },
                ) {
                    out.push(v.into_any());
                }
                hi += 1;
            }
        }
        out
    }

    /// Insert a new paragraph before this one (python-docx
    /// insert_paragraph_before); returns the new paragraph.
    #[pyo3(signature = (text=None, style=None))]
    fn insert_paragraph_before(
        &self,
        py: Python<'_>,
        text: Option<&str>,
        style: Option<&str>,
    ) -> PyResult<PyParagraph> {
        let sid = match style {
            Some(s) => Some(with_core(&self.tpl, py, |core| {
                crate::subdocbuilder::resolve_style_id(core, s)
            })),
            None => None,
        };
        with_core(&self.tpl, py, |core| {
            let mut found = false;
            mutate_document(core, |body| {
                let mut seen = 0usize;
                let mut pos = None;
                for (i, c) in body.children.iter().enumerate() {
                    if matches!(c, Node::Elem(e) if e.name == "w:p") {
                        if seen == self.index {
                            pos = Some(i);
                            break;
                        }
                        seen += 1;
                    }
                }
                if let Some(i) = pos {
                    let mut p = Element::new("w:p");
                    if let Some(sid) = &sid {
                        let mut ppr = Element::new("w:pPr");
                        let mut ps = Element::new("w:pStyle");
                        ps.set_attr("w:val", sid);
                        ppr.children.push(Node::Elem(ps));
                        p.children.push(Node::Elem(ppr));
                    }
                    if let Some(t) = text {
                        if !t.is_empty() {
                            let mut r = Element::new("w:r");
                            let mut wt = Element::new("w:t");
                            wt.set_attr("xml:space", "preserve");
                            wt.children.push(Node::Text(t.to_string()));
                            r.children.push(Node::Elem(wt));
                            p.children.push(Node::Elem(r));
                        }
                    }
                    body.children.insert(i, Node::Elem(p));
                    found = true;
                }
            })
            .map_err(py_err)?;
            if !found {
                return Err(PyValueError::new_err("paragraph not found"));
            }
            Ok(PyParagraph {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
            })
        })
    }
}

/// A run inside a paragraph (live proxy).
#[pyclass(name = "Run", unsendable)]
pub struct PyRun {
    pub tpl: Py<PyDocxTemplate>,
    pub para: usize,
    pub index: usize,
}

impl PyRun {
    pub(crate) fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            let gen = core.doc_gen;
            let mut cur = core.para_cursor;
            mutate_document(core, |body| {
                if let Some(p) = nth_cursor_mut(body, "w:p", self.para, &mut cur, gen) {
                    if let Some(r) = nth_direct(p, "w:r", self.index) {
                        result = Some(f(r));
                    }
                }
            })
            .map_err(py_err)?;
            core.para_cursor = cur;
            result.ok_or_else(|| PyValueError::new_err("run not found"))
        })
    }

    pub(crate) fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            let gen = core.doc_gen;
            let mut cur = core.para_cursor;
            let r = read_body(core, |body| {
                nth_cursor_ref(body, "w:p", self.para, &mut cur, gen)
                    .and_then(|p| nth_direct_ref(p, "w:r", self.index))
                    .map(|r| f(r))
            })
            .flatten();
            core.para_cursor = cur;
            r
        })
    }
}

pub(crate) fn nth_direct_ref<'a>(el: &'a Element, name: &str, n: usize) -> Option<&'a Element> {
    el.children
        .iter()
        .filter_map(|c| match c {
            Node::Elem(e) if e.name == name => Some(e),
            _ => None,
        })
        .nth(n)
}

/// nth child named `name`, resuming from a validated sequential-access
/// cursor (doc_gen, index, child_pos) when possible: sequential proxy
/// iteration becomes O(1) amortized per access instead of rescanning from
/// the head (O(n^2) for a full document walk). The cursor is validated
/// against the document's mutation generation, so any DOM change safely
/// falls back to a full scan.
pub(crate) fn nth_cursor_ref<'a>(
    el: &'a Element,
    name: &str,
    n: usize,
    cur: &mut (u64, usize, usize),
    gen: u64,
) -> Option<&'a Element> {
    let (cgen, ci, cp) = *cur;
    let (mut idx, mut i) = if cgen == gen
        && ci <= n
        && cp < el.children.len()
        && matches!(&el.children[cp], Node::Elem(e) if e.name == name)
    {
        (ci, cp)
    } else {
        (0, 0)
    };
    while i < el.children.len() {
        if let Node::Elem(e) = &el.children[i] {
            if e.name == name {
                if idx == n {
                    *cur = (gen, n, i);
                    return Some(e);
                }
                idx += 1;
            }
        }
        i += 1;
    }
    None
}

/// nth_cursor_ref for mutable access
pub(crate) fn nth_cursor_mut<'a>(
    el: &'a mut Element,
    name: &str,
    n: usize,
    cur: &mut (u64, usize, usize),
    gen: u64,
) -> Option<&'a mut Element> {
    let (cgen, ci, cp) = *cur;
    let (mut idx, mut i) = if cgen == gen
        && ci <= n
        && cp < el.children.len()
        && matches!(&el.children[cp], Node::Elem(e) if e.name == name)
    {
        (ci, cp)
    } else {
        (0, 0)
    };
    let mut found = None;
    while i < el.children.len() {
        if let Node::Elem(e) = &el.children[i] {
            if e.name == name {
                if idx == n {
                    found = Some(i);
                    break;
                }
                idx += 1;
            }
        }
        i += 1;
    }
    let pos = found?;
    *cur = (gen, n, pos);
    match &mut el.children[pos] {
        Node::Elem(e) => Some(e),
        _ => None,
    }
}

pub(crate) fn ensure_rpr(run: &mut Element) -> &mut Element {
    let has = run
        .children
        .iter()
        .any(|c| matches!(c, Node::Elem(e) if e.name == "w:rPr"));
    if !has {
        run.children.insert(0, Node::Elem(Element::new("w:rPr")));
    }
    run.find_mut("w:rPr").unwrap()
}

pub(crate) fn set_flag(r: &mut Element, tag: &str, on: bool) {
    let rpr = ensure_rpr(r);
    let exists = rpr.children.iter().any(|c| matches!(c, Node::Elem(e) if e.name == tag));
    if on && !exists {
        rpr.children.push(Node::Elem(Element::new(tag)));
    } else if !on && exists {
        rpr.children
            .retain(|c| !matches!(c, Node::Elem(e) if e.name == tag));
    }
}

pub(crate) fn set_val_tag(r: &mut Element, tag: &str, attr: &str, val: &str) {
    let rpr = ensure_rpr(r);
    if let Some(el) = rpr.find_mut(tag) {
        el.set_attr(attr, val);
    } else {
        let mut el = Element::new(tag);
        el.set_attr(attr, val);
        rpr.children.push(Node::Elem(el));
    }
}

pub(crate) fn read_flag(el: &Element, tag: &str) -> Option<bool> {
    let rpr = el.find("w:rPr")?;
    rpr.find(tag).map(|e| {
        !matches!(
            e.get_attr("w:val"),
            Some("0") | Some("false") | Some("off") | Some("none")
        )
    })
}

#[pymethods]
impl PyRun {

    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        self.read(py, |r| element_text(r)).unwrap_or_default()
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| {
            r.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:t" || e.name == "w:tab" || e.name == "w:br" || e.name == "w:cr"));
            run_append_text(r, &v);
        })
    }

    #[getter]
    fn bold(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |r| read_flag(r, "w:b")).flatten()
    }
    #[setter]
    fn set_bold(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| set_flag(r, "w:b", v))
    }

    #[getter]
    fn italic(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |r| read_flag(r, "w:i")).flatten()
    }
    #[setter]
    fn set_italic(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| set_flag(r, "w:i", v))
    }

    #[getter]
    fn strike(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |r| read_flag(r, "w:strike")).flatten()
    }
    #[setter]
    fn set_strike(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| set_flag(r, "w:strike", v))
    }

    #[getter]
    fn underline(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:u"))
                .and_then(|u| u.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_underline(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        if let Ok(s) = v.extract::<String>() {
            self.edit(py, |r| set_val_tag(r, "w:u", "w:val", &s))
        } else if v.is_truthy()? {
            self.edit(py, |r| set_val_tag(r, "w:u", "w:val", "single"))
        } else {
            self.edit(py, |r| set_flag(r, "w:u", false))
        }
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:rStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let sid = with_core(&self.tpl, py, |core| {
            crate::subdocbuilder::resolve_style_id(core, &v)
        });
        self.edit(py, |r| set_val_tag(r, "w:rStyle", "w:val", &sid))
    }

    #[getter]
    fn font_name(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:rFonts"))
                .and_then(|e| e.get_attr("w:ascii").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_font(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| {
            let rpr = ensure_rpr(r);
            if let Some(el) = rpr.find_mut("w:rFonts") {
                el.set_attr("w:ascii", &v);
                el.set_attr("w:hAnsi", &v);
                el.set_attr("w:cs", &v);
            } else {
                let mut el = Element::new("w:rFonts");
                el.set_attr("w:ascii", &v);
                el.set_attr("w:hAnsi", &v);
                el.set_attr("w:cs", &v);
                rpr.children.push(Node::Elem(el));
            }
        })
    }

    #[getter]
    fn size(&self, py: Python<'_>) -> Option<u32> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:sz"))
                .and_then(|e| e.get_attr("w:val").and_then(|v| v.parse().ok()))
        })
        .flatten()
    }
    #[setter]
    fn set_size(&self, py: Python<'_>, v: u32) -> PyResult<()> {
        let s = v.to_string();
        self.edit(py, |r| {
            set_val_tag(r, "w:sz", "w:val", &s);
            set_val_tag(r, "w:szCs", "w:val", &s);
        })
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:color"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_color(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let c = v.strip_prefix('#').unwrap_or(&v).to_string();
        self.edit(py, |r| set_val_tag(r, "w:color", "w:val", &c))
    }

    #[getter]
    fn highlight(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:shd"))
                .and_then(|e| e.get_attr("w:fill").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_highlight(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let c = v.strip_prefix('#').unwrap_or(&v).to_string();
        self.edit(py, |r| set_val_tag(r, "w:shd", "w:fill", &c))
    }

    #[getter]
    fn subscript(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:vertAlign"))
                .map(|e| e.get_attr("w:val") == Some("subscript"))
        })
        .flatten()
    }
    #[setter]
    fn set_subscript(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        if v {
            self.edit(py, |r| set_val_tag(r, "w:vertAlign", "w:val", "subscript"))
        } else {
            self.edit(py, |r| set_flag(r, "w:vertAlign", false))
        }
    }

    #[getter]
    fn superscript(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:vertAlign"))
                .map(|e| e.get_attr("w:val") == Some("superscript"))
        })
        .flatten()
    }
    #[setter]
    fn set_superscript(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        if v {
            self.edit(py, |r| set_val_tag(r, "w:vertAlign", "w:val", "superscript"))
        } else {
            self.edit(py, |r| set_flag(r, "w:vertAlign", false))
        }
    }

    /// Full python-docx Font facade (tri-state booleans, all 29 properties).
    #[getter]
    fn font(&self, py: Python<'_>) -> crate::docmodel_fmt::PyFont {
        crate::docmodel_fmt::PyFont {
            tpl: self.tpl.clone_ref(py),
            target: crate::docmodel_fmt::FontTarget::Run {
                para: self.para,
                index: self.index,
            },
        }
    }

    /// Add a break; break_type is a WD_BREAK int (6=line, 7=page, 8=column,
    /// 9/10/11=textWrapping clear left/right/all).
    #[pyo3(signature = (break_type=6))]
    fn add_break(&self, py: Python<'_>, break_type: i64) -> PyResult<()> {
        self.edit(py, |r| {
            let mut br = Element::new("w:br");
            match break_type {
                7 => br.set_attr("w:type", "page"),
                8 => br.set_attr("w:type", "column"),
                9 => {
                    br.set_attr("w:type", "textWrapping");
                    br.set_attr("w:clear", "left");
                }
                10 => {
                    br.set_attr("w:type", "textWrapping");
                    br.set_attr("w:clear", "right");
                }
                11 => {
                    br.set_attr("w:type", "textWrapping");
                    br.set_attr("w:clear", "all");
                }
                _ => {}
            }
            r.children.push(Node::Elem(br));
        })
    }

    /// Append a tab character (w:tab).
    fn add_tab(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |r| r.children.push(Node::Elem(Element::new("w:tab"))))
    }

    /// Append text (w:t), preserving leading/trailing whitespace.
    fn add_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        let text = text.to_string();
        self.edit(py, |r| run_append_text(r, &text))
    }

    /// Remove all content, keeping run properties (python-docx clear).
    fn clear(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |r| {
            r.children
                .retain(|c| matches!(c, Node::Elem(e) if e.name == "w:rPr"));
        })
    }

    /// True when a rendered page break (w:lastRenderedPageBreak) occurs in
    /// this run (hard breaks are not counted, python-docx semantics).
    #[getter]
    fn contains_page_break(&self, py: Python<'_>) -> bool {
        self.read(py, |r| {
            let mut out = Vec::new();
            r.iter_descendants("w:lastRenderedPageBreak", &mut out);
            !out.is_empty()
        })
        .unwrap_or(false)
    }

    /// Content items of this run in order (python-docx iter_inner_content):
    /// contiguous text-ish ranges as strings, drawings as live XmlElement
    /// proxies, rendered page breaks as RenderedPageBreak markers.
    fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        // element-index path of this run: [body, para, run]
        let path_and_items: Option<(Vec<usize>, Vec<(bool, usize, String)>)> = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                // element-child index of the nth w:p
                let mut seen = 0usize;
                let mut p_pos = None;
                for (i, c) in body.children.iter().enumerate() {
                    if matches!(c, Node::Elem(e) if e.name == "w:p") {
                        if seen == self.para {
                            p_pos = Some(i);
                            break;
                        }
                        seen += 1;
                    }
                }
                let p_pos = p_pos?;
                let p = match &body.children[p_pos] {
                    Node::Elem(e) => e,
                    _ => return None,
                };
                let mut seen = 0usize;
                let mut r_pos = None;
                for (i, c) in p.children.iter().enumerate() {
                    if matches!(c, Node::Elem(e) if e.name == "w:r") {
                        if seen == self.index {
                            r_pos = Some(i);
                            break;
                        }
                        seen += 1;
                    }
                }
                let r_pos = r_pos?;
                let r = match &p.children[r_pos] {
                    Node::Elem(e) => e,
                    _ => return None,
                };
                let mut items: Vec<(bool, usize, String)> = Vec::new();
                let mut cur = String::new();
                for (i, c) in r.children.iter().enumerate() {
                    if let Node::Elem(e) = c {
                        match e.name.as_str() {
                            "w:t" => cur.push_str(&e.text_content()),
                            "w:tab" => cur.push('\t'),
                            "w:br" | "w:cr" => cur.push('\n'),
                            "w:noBreakHyphen" => cur.push('\u{2011}'),
                            "w:drawing" => {
                                if !cur.is_empty() {
                                    items.push((false, 0, std::mem::take(&mut cur)));
                                }
                                items.push((true, i, String::new())); // drawing
                            }
                            "w:lastRenderedPageBreak" => {
                                if !cur.is_empty() {
                                    items.push((false, 0, std::mem::take(&mut cur)));
                                }
                                items.push((true, i, "pb".to_string())); // marker
                            }
                            _ => {}
                        }
                    }
                }
                if !cur.is_empty() {
                    items.push((false, 0, cur));
                }
                Some((vec![p_pos, r_pos], items))
            })
            .flatten()
        });
        let Some((tail, items)) = path_and_items else {
            return Vec::new();
        };
        // element index of w:body within w:document (almost always 0)
        let body_idx = with_core(&self.tpl, py, |core| {
            core.document_dom().ok().and_then(|dom| {
                dom.root
                    .children
                    .iter()
                    .filter_map(|c| match c {
                        Node::Elem(e) => Some(e),
                        _ => None,
                    })
                    .position(|e| e.name == "w:body")
            })
        })
        .unwrap_or(0);
        let mut base = vec![body_idx];
        base.extend(tail);
        let mut out = Vec::new();
        for (is_elem, pos, text) in items {
            if !is_elem {
                out.push(pyo3::types::PyString::new(py, &text).into_any().unbind());
            } else if text == "pb" {
                if let Ok(v) = Py::new(py, crate::docmodel_fmt::PyRenderedPageBreak {}) {
                    out.push(v.into_any());
                }
            } else {
                let mut path = base.clone();
                path.push(pos);
                if let Ok(v) = Py::new(
                    py,
                    crate::pyxml::PyXmlElement {
                        tpl: self.tpl.clone_ref(py),
                        part: DOCUMENT_PART.to_string(),
                        path,
                    },
                ) {
                    out.push(v.into_any());
                }
            }
        }
        out
    }

    /// Append a picture to this run (python-docx run.add_picture).
    #[pyo3(signature = (image_descriptor, width=None, height=None))]
    fn add_picture(
        &self,
        py: Python<'_>,
        image_descriptor: &Bound<'_, PyAny>,
        width: Option<i64>,
        height: Option<i64>,
    ) -> PyResult<()> {
        let (blob, filename) = crate::docmodel_add::read_image_source(image_descriptor)?;
        let drawing = with_core(&self.tpl, py, |core| -> Result<String, String> {
            core.init_docx(false)?;
            crate::inline_image::drawing_xml(
                core,
                DOCUMENT_PART,
                &blob,
                filename.as_deref(),
                width,
                height,
                None,
                None,
                None,
            )
        })
        .map_err(py_err)?;
        self.edit(py, |r| {
            if let Ok(frag) = crate::subdoc::parse_body_fragment(&drawing) {
                r.children.extend(frag.root.children);
            }
        })
    }

    /// Mark the range from this run to `last_run` as belonging to the
    /// comment `comment_id` (python-docx run.mark_comment_range).
    fn mark_comment_range(&self, py: Python<'_>, last_run: Bound<'_, PyRun>, comment_id: i64) -> PyResult<()> {
        let (lpara, lindex) = (last_run.borrow().para, last_run.borrow().index);
        let id = comment_id.to_string();
        with_core(&self.tpl, py, |core| {
            mutate_document(core, |body| {
                // end marker first so positions for the start marker are stable
                for (para_idx, run_idx, is_start) in [
                    (lpara, lindex, false),
                    (self.para, self.index, true),
                ] {
                    let Some(p) = nth_direct(body, "w:p", para_idx) else {
                        continue;
                    };
                    // child position of the nth w:r
                    let mut seen = 0usize;
                    let mut pos = None;
                    for (i, c) in p.children.iter().enumerate() {
                        if matches!(c, Node::Elem(e) if e.name == "w:r") {
                            if seen == run_idx {
                                pos = Some(i);
                                break;
                            }
                            seen += 1;
                        }
                    }
                    let Some(i) = pos else { continue };
                    if is_start {
                        let mut cs = Element::new("w:commentRangeStart");
                        cs.set_attr("w:id", &id);
                        p.children.insert(i, Node::Elem(cs));
                    } else {
                        let mut ce = Element::new("w:commentRangeEnd");
                        ce.set_attr("w:id", &id);
                        let mut rr = Element::new("w:r");
                        let mut cr = Element::new("w:commentReference");
                        cr.set_attr("w:id", &id);
                        rr.children.push(Node::Elem(cr));
                        p.children.insert(i + 1, Node::Elem(rr));
                        p.children.insert(i + 1, Node::Elem(ce));
                    }
                }
            })
            .map_err(py_err)
        })
    }
}

/// Append text to a run, expanding \t -> w:tab, \n -> w:br, \r -> w:cr
/// (python-docx Run.text semantics).
pub(crate) fn run_append_text(r: &mut Element, text: &str) {
    let mut buf = String::new();
    let flush = |r: &mut Element, buf: &mut String| {
        if buf.is_empty() {
            return;
        }
        let mut t = Element::new("w:t");
        t.set_attr("xml:space", "preserve");
        t.children.push(Node::Text(std::mem::take(buf)));
        r.children.push(Node::Elem(t));
    };
    for ch in text.chars() {
        match ch {
            '\t' => {
                flush(r, &mut buf);
                r.children.push(Node::Elem(Element::new("w:tab")));
            }
            '\n' => {
                flush(r, &mut buf);
                r.children.push(Node::Elem(Element::new("w:br")));
            }
            '\r' => {
                flush(r, &mut buf);
                r.children.push(Node::Elem(Element::new("w:cr")));
            }
            _ => buf.push(ch),
        }
    }
    flush(r, &mut buf);
}

/// A table in the document (live proxy).
#[pyclass(name = "Table", unsendable)]
pub struct PyTable {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

impl PyTable {
    pub(crate) fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            let gen = core.doc_gen;
            let mut cur = core.tbl_cursor;
            let r = read_body(core, |body| {
                nth_cursor_ref(body, "w:tbl", self.index, &mut cur, gen).map(|t| f(t))
            })
            .flatten();
            core.tbl_cursor = cur;
            r
        })
    }

    pub(crate) fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            let gen = core.doc_gen;
            let mut cur = core.tbl_cursor;
            mutate_document(core, |body| {
                if let Some(t) = nth_cursor_mut(body, "w:tbl", self.index, &mut cur, gen) {
                    result = Some(f(t));
                }
            })
            .map_err(py_err)?;
            core.tbl_cursor = cur;
            result.ok_or_else(|| PyValueError::new_err("table not found"))
        })
    }
}

#[pymethods]
impl PyTable {

    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    #[getter]
    fn rows(&self, py: Python<'_>) -> Vec<PyTableRow> {
        let n = self
            .read(py, |t| {
                t.children
                    .iter()
                    .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tr"))
                    .count()
            })
            .unwrap_or(0);
        (0..n)
            .map(|row| PyTableRow {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row,
            })
            .collect()
    }

    fn add_row(&self, py: Python<'_>) -> PyResult<PyTableRow> {
        self.edit(py, |t| {
            let cols = nth_direct_ref(t, "w:tr", 0)
                .map(|r| {
                    r.children
                        .iter()
                        .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                        .count()
                })
                .unwrap_or(1);
            let mut tr = Element::new("w:tr");
            for _ in 0..cols {
                let mut tc = Element::new("w:tc");
                tc.children.push(Node::Elem(Element::new("w:p")));
                tr.children.push(Node::Elem(tc));
            }
            let n = t
                .children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tr"))
                .count();
            t.children.push(Node::Elem(tr));
            PyTableRow {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: n,
            }
        })
    }

    fn cell(&self, py: Python<'_>, i: usize, j: usize) -> PyCell {
        PyCell {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: i,
            col: j,
        }
    }

    /// The table style id (python-docx table.style; accepts a style name or
    /// id on assignment, returns the style id).
    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |t| {
            t.find("w:tblPr")
                .and_then(|p| p.find("w:tblStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let sid = with_core(&self.tpl, py, |core| {
            crate::subdocbuilder::resolve_style_id(core, &v)
        });
        self.edit(py, |t| {
            // w:tblPr must be the first child of w:tbl
            if t.find("w:tblPr").is_none() {
                t.children.insert(0, Node::Elem(Element::new("w:tblPr")));
            }
            let tblpr = t.find_mut("w:tblPr").unwrap();
            // w:tblStyle must be the first child of w:tblPr
            if tblpr.find("w:tblStyle").is_none() {
                tblpr.children.insert(0, Node::Elem(Element::new("w:tblStyle")));
            }
            tblpr.find_mut("w:tblStyle").unwrap().set_attr("w:val", &sid);
        })
    }

    /// Table alignment as a WD_TABLE_ALIGNMENT int (xml name also accepted).
    #[getter]
    fn alignment(&self, py: Python<'_>) -> Option<i64> {
        self.read(py, |t| {
            t.find("w:tblPr")
                .and_then(|p| p.find("w:jc"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
        .map(|s| match s.as_str() {
            "center" => 1,
            "right" => 2,
            _ => 0,
        })
    }
    #[setter]
    fn set_alignment(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let val: Option<&'static str> = if v.is_none() {
            None
        } else if let Ok(i) = v.extract::<i64>() {
            Some(match i {
                1 => "center",
                2 => "right",
                _ => "left",
            })
        } else {
            let s: String = v.extract()?;
            Some(match s.as_str() {
                "center" => "center",
                "right" => "right",
                _ => "left",
            })
        };
        self.edit(py, |t| {
            let tblpr = crate::docmodel_fmt::ensure_tblpr(t);
            match val {
                Some(x) => {
                    let jc = crate::docmodel_fmt::ensure_child(tblpr, "w:jc");
                    jc.set_attr("w:val", x);
                }
                None => tblpr
                    .children
                    .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:jc")),
            }
        })
    }

    /// Autofit (tblLayout type=autofit vs fixed; missing -> True).
    #[getter]
    fn autofit(&self, py: Python<'_>) -> bool {
        self.read(py, |t| {
            t.find("w:tblPr")
                .and_then(|p| p.find("w:tblLayout"))
                .and_then(|e| e.get_attr("w:type"))
                .map(|ty| ty != "fixed")
                .unwrap_or(true)
        })
        .unwrap_or(true)
    }
    #[setter]
    fn set_autofit(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |t| {
            let tblpr = crate::docmodel_fmt::ensure_tblpr(t);
            let l = crate::docmodel_fmt::ensure_child(tblpr, "w:tblLayout");
            l.set_attr("w:type", if v { "autofit" } else { "fixed" });
        })
    }

    /// Append a column of the given width (gridCol + one cell per row).
    fn add_column(&self, py: Python<'_>, width: &Bound<'_, PyAny>) -> PyResult<crate::docmodel_fmt::PyTableColumn> {
        let emu = crate::pyclasses::extract_length_pub(width)?
            .ok_or_else(|| PyValueError::new_err("width is required"))?;
        let twips = (emu / 635).to_string();
        let col = self.edit(py, |t| {
            let col = t
                .find("w:tblGrid")
                .map(|g| {
                    g.children
                        .iter()
                        .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:gridCol"))
                        .count()
                })
                .unwrap_or(0);
            // ensure tblGrid right after tblPr
            if t.find("w:tblGrid").is_none() {
                let pos = if t.find("w:tblPr").is_some() { 1 } else { 0 };
                t.children.insert(pos, Node::Elem(Element::new("w:tblGrid")));
            }
            let grid = t.find_mut("w:tblGrid").unwrap();
            let mut gc = Element::new("w:gridCol");
            gc.set_attr("w:w", &twips);
            grid.children.push(Node::Elem(gc));
            for c in t.children.iter_mut() {
                if let Node::Elem(tr) = c {
                    if tr.name == "w:tr" {
                        let mut tc = Element::new("w:tc");
                        let mut tcpr = Element::new("w:tcPr");
                        let mut tcw = Element::new("w:tcW");
                        tcw.set_attr("w:type", "dxa");
                        tcw.set_attr("w:w", &twips);
                        tcpr.children.push(Node::Elem(tcw));
                        tc.children.push(Node::Elem(tcpr));
                        tc.children.push(Node::Elem(Element::new("w:p")));
                        tr.children.push(Node::Elem(tc));
                    }
                }
            }
            col
        })?;
        Ok(crate::docmodel_fmt::PyTableColumn {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            col,
        })
    }

    #[getter]
    fn columns(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyTableColumn> {
        let n = self
            .read(py, |t| {
                t.find("w:tblGrid")
                    .map(|g| {
                        g.children
                            .iter()
                            .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:gridCol"))
                            .count()
                    })
                    .unwrap_or(0)
            })
            .unwrap_or(0);
        (0..n)
            .map(|col| crate::docmodel_fmt::PyTableColumn {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                col,
            })
            .collect()
    }

    /// Table direction: 0=ltr, 1=rtl (w:bidiVisual).
    #[getter]
    fn table_direction(&self, py: Python<'_>) -> i64 {
        self.read(py, |t| {
            t.find("w:tblPr")
                .map(|p| p.find("w:bidiVisual").is_some() as i64)
        })
        .flatten()
        .unwrap_or(0)
    }
    #[setter]
    fn set_table_direction(&self, py: Python<'_>, v: i64) -> PyResult<()> {
        self.edit(py, |t| {
            let tblpr = crate::docmodel_fmt::ensure_tblpr(t);
            if v == 1 {
                if tblpr.find("w:bidiVisual").is_none() {
                    tblpr.children.push(Node::Elem(Element::new("w:bidiVisual")));
                }
            } else {
                tblpr.children
                    .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:bidiVisual"));
            }
        })
    }

    /// Cells of column `i` (one per row).
    fn column_cells(&self, py: Python<'_>, i: usize) -> Vec<PyCell> {
        let rows = self
            .read(py, |t| {
                t.children
                    .iter()
                    .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tr"))
                    .count()
            })
            .unwrap_or(0);
        (0..rows)
            .map(|row| PyCell {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row,
                col: i,
            })
            .collect()
    }

    /// Cells of row `i`.
    fn row_cells(&self, py: Python<'_>, i: usize) -> Vec<PyCell> {
        let cols = self
            .read(py, |t| {
                nth_direct_ref(t, "w:tr", i)
                    .map(|r| {
                        r.children
                            .iter()
                            .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                            .count()
                    })
                    .unwrap_or(0)
            })
            .unwrap_or(0);
        (0..cols)
            .map(|col| PyCell {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: i,
                col,
            })
            .collect()
    }
}

/// A table row (live proxy).
#[pyclass(name = "TableRow", unsendable)]
pub struct PyTableRow {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
}

impl PyTableRow {
    fn row_grid_cols(&self, py: Python<'_>, tag: &str) -> i64 {
        self.read(py, |r| {
            r.find("w:trPr")
                .and_then(|p| p.find(tag))
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
                .unwrap_or(0)
        })
        .unwrap_or(0)
    }

    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        PyTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
        }
        .read(py, |t| nth_direct_ref(t, "w:tr", self.row).map(|r| f(r)))
        .flatten()
    }

    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        let row = self.row;
        PyTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
        }
        .edit(py, |t| {
            nth_direct(t, "w:tr", row)
                .map(|r| f(r))
                .ok_or_else(|| "row not found".to_string())
        })?
        .map_err(PyValueError::new_err)
    }
}

#[pymethods]
impl PyTableRow {
    /// Row height (w:trPr/w:trHeight w:val).
    #[getter]
    fn height(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |r| {
            r.find("w:trPr")
                .and_then(|p| p.find("w:trHeight"))
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
        })
        .flatten()
        .pipe_map(to_len)
    }
    #[setter]
    fn set_height(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |r| {
            let trpr = crate::docmodel_fmt::ensure_trpr(r);
            let h = crate::docmodel_fmt::ensure_child(trpr, "w:trHeight");
            match twips {
                Some(t) => h.set_attr("w:val", &t.to_string()),
                None => h.attrs.retain(|(k, _)| k != "w:val"),
            }
        })
    }

    /// Row height rule as a WD_ROW_HEIGHT_RULE int (0=auto, 1=atLeast,
    /// 2=exact; xml name also accepted on set).
    #[getter]
    fn height_rule(&self, py: Python<'_>) -> Option<i64> {
        self.read(py, |r| {
            r.find("w:trPr")
                .and_then(|p| p.find("w:trHeight"))
                .and_then(|e| e.get_attr("w:hRule"))
                .map(|s| match s {
                    "atLeast" => 1,
                    "exact" => 2,
                    _ => 0,
                })
        })
        .flatten()
    }
    #[setter]
    fn set_height_rule(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let val: Option<&'static str> = if v.is_none() {
            None
        } else if let Ok(i) = v.extract::<i64>() {
            Some(match i {
                1 => "atLeast",
                2 => "exact",
                _ => "auto",
            })
        } else {
            let s: String = v.extract()?;
            Some(match s.as_str() {
                "atLeast" => "atLeast",
                "exact" => "exact",
                _ => "auto",
            })
        };
        self.edit(py, |r| {
            let trpr = crate::docmodel_fmt::ensure_trpr(r);
            let h = crate::docmodel_fmt::ensure_child(trpr, "w:trHeight");
            match val {
                Some(x) => h.set_attr("w:hRule", x),
                None => h.attrs.retain(|(k, _)| k != "w:hRule"),
            }
        })
    }


    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    /// Grid columns before this row (trPr/gridBefore; default 0).
    #[getter]
    fn grid_cols_before(&self, py: Python<'_>) -> i64 {
        self.row_grid_cols(py, "w:gridBefore")
    }
    /// Grid columns after this row (trPr/gridAfter; default 0).
    #[getter]
    fn grid_cols_after(&self, py: Python<'_>) -> i64 {
        self.row_grid_cols(py, "w:gridAfter")
    }

    #[getter]
    fn cells(&self, py: Python<'_>) -> Vec<PyCell> {
        let n = with_core(&self.tpl, py, |core| {
            let gen = core.doc_gen;
            let mut tcur = core.tbl_cursor;
            let mut rcur = core.row_cursor;
            let r = read_body(core, |body| {
                nth_cursor_ref(body, "w:tbl", self.index, &mut tcur, gen)
                    .and_then(|t| nth_cursor_ref(t, "w:tr", self.row, &mut rcur, gen))
                    .map(|r| {
                        r.children
                            .iter()
                            .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                            .count()
                    })
            })
            .flatten();
            core.tbl_cursor = tcur;
            core.row_cursor = rcur;
            r
        })
        .unwrap_or(0);
        (0..n)
            .map(|col| PyCell {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: self.row,
                col,
            })
            .collect()
    }
}

/// A table cell (live proxy).
#[pyclass(name = "Cell", unsendable)]
pub struct PyCell {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
}

impl PyCell {
    pub(crate) fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            let gen = core.doc_gen;
            let mut tcur = core.tbl_cursor;
            let mut rcur = core.row_cursor;
            mutate_document(core, |body| {
                if let Some(t) = nth_cursor_mut(body, "w:tbl", self.index, &mut tcur, gen) {
                    if let Some(r) = nth_cursor_mut(t, "w:tr", self.row, &mut rcur, gen) {
                        if let Some(c) = nth_direct(r, "w:tc", self.col) {
                            result = Some(f(c));
                        }
                    }
                }
            })
            .map_err(py_err)?;
            core.tbl_cursor = tcur;
            core.row_cursor = rcur;
            result.ok_or_else(|| PyValueError::new_err("cell not found"))
        })
    }
}

#[pymethods]
impl PyCell {

    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| {
            let gen = core.doc_gen;
            let mut tcur = core.tbl_cursor;
            let mut rcur = core.row_cursor;
            let r = read_body(core, |body| {
                nth_cursor_ref(body, "w:tbl", self.index, &mut tcur, gen)
                    .and_then(|t| nth_cursor_ref(t, "w:tr", self.row, &mut rcur, gen))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .map(|c| element_text(c))
            })
            .flatten();
            core.tbl_cursor = tcur;
            core.row_cursor = rcur;
            r
        })
        .unwrap_or_default()
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |c| {
            c.children
                .retain(|ch| !matches!(ch, Node::Elem(e) if e.name == "w:p"));
            let mut p = Element::new("w:p");
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v));
            r.children.push(Node::Elem(t));
            p.children.push(Node::Elem(r));
            c.children.push(Node::Elem(p));
        })
    }

    /// Paragraphs in this cell.
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyCellParagraph> {
        let n = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .map(|c| {
                        c.children
                            .iter()
                            .filter(|ch| matches!(ch, Node::Elem(e) if e.name == "w:p"))
                            .count()
                    })
            })
            .flatten()
        })
        .unwrap_or(0);
        (0..n)
            .map(|para| crate::docmodel_fmt::PyCellParagraph {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: self.row,
                col: self.col,
                para,
            })
            .collect()
    }

    /// Append a paragraph to this cell (python-docx cell.add_paragraph).
    #[pyo3(signature = (text="", style=None))]
    fn add_paragraph(
        &self,
        py: Python<'_>,
        text: &str,
        style: Option<&str>,
    ) -> PyResult<crate::docmodel_fmt::PyCellParagraph> {
        let sid = match style {
            Some(s) => Some(with_core(&self.tpl, py, |core| {
                crate::subdocbuilder::resolve_style_id(core, s)
            })),
            None => None,
        };
        let text = text.to_string();
        let para = self.edit(py, |c| {
            let n = c
                .children
                .iter()
                .filter(|ch| matches!(ch, Node::Elem(e) if e.name == "w:p"))
                .count();
            let mut p = Element::new("w:p");
            if let Some(sid) = &sid {
                let mut ppr = Element::new("w:pPr");
                let mut ps = Element::new("w:pStyle");
                ps.set_attr("w:val", sid);
                ppr.children.push(Node::Elem(ps));
                p.children.push(Node::Elem(ppr));
            }
            if !text.is_empty() {
                let mut r = Element::new("w:r");
                let mut t = Element::new("w:t");
                t.set_attr("xml:space", "preserve");
                t.children.push(Node::Text(text.clone()));
                r.children.push(Node::Elem(t));
                p.children.push(Node::Elem(r));
            }
            c.children.push(Node::Elem(p));
            n
        })?;
        Ok(crate::docmodel_fmt::PyCellParagraph {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
            para,
        })
    }

    /// Vertical alignment as a WD_CELL_VERTICAL_ALIGNMENT int (0=top,
    /// 1=center, 3=bottom, 101=both; xml name also accepted on set).
    #[getter]
    fn vertical_alignment(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .and_then(|c| c.find("w:tcPr"))
                    .and_then(|p| p.find("w:vAlign"))
                    .and_then(|e| e.get_attr("w:val"))
                    .map(|s| match s {
                        "center" => 1,
                        "bottom" => 3,
                        "both" => 101,
                        _ => 0,
                    })
            })
            .flatten()
        })
    }
    #[setter]
    fn set_vertical_alignment(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let val: Option<&'static str> = if v.is_none() {
            None
        } else if let Ok(i) = v.extract::<i64>() {
            Some(match i {
                1 => "center",
                3 => "bottom",
                101 => "both",
                _ => "top",
            })
        } else {
            let s: String = v.extract()?;
            Some(match s.as_str() {
                "center" => "center",
                "bottom" => "bottom",
                "both" => "both",
                _ => "top",
            })
        };
        self.edit(py, |c| {
            let tcpr = tcpr_mut(c);
            match val {
                Some(x) => {
                    let va = crate::docmodel_fmt::ensure_child(tcpr, "w:vAlign");
                    va.set_attr("w:val", x);
                }
                None => tcpr
                    .children
                    .retain(|ch| !matches!(ch, Node::Elem(e) if e.name == "w:vAlign")),
            }
        })
    }

    /// Cell width (w:tcPr/w:tcW; set forces type=dxa).
    #[getter]
    fn width(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .and_then(|c| c.find("w:tcPr"))
                    .and_then(|p| p.find("w:tcW"))
                    .and_then(|e| {
                        if e.get_attr("w:type") == Some("dxa") {
                            e.get_attr("w:w").and_then(|s| s.parse::<i64>().ok())
                        } else {
                            None
                        }
                    })
            })
            .flatten()
            .pipe_map(to_len)
        })
    }
    #[setter]
    fn set_width(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |c| {
            let tcpr = tcpr_mut(c);
            let tcw = crate::docmodel_fmt::ensure_child(tcpr, "w:tcW");
            match twips {
                Some(t) => {
                    tcw.set_attr("w:type", "dxa");
                    tcw.set_attr("w:w", &t.to_string());
                }
                None => tcpr
                    .children
                    .retain(|ch| !matches!(ch, Node::Elem(e) if e.name == "w:tcW")),
            }
        })
    }

    /// Grid columns spanned by this cell (w:gridSpan; default 1).
    #[getter]
    fn grid_span(&self, py: Python<'_>) -> i64 {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .and_then(|c| c.find("w:tcPr"))
                    .and_then(|p| p.find("w:gridSpan"))
                    .and_then(|e| e.get_attr("w:val"))
                    .and_then(|s| s.parse::<i64>().ok())
                    .unwrap_or(1)
            })
            .unwrap_or(1)
        })
    }

    /// Tables nested inside this cell.
    #[getter]
    fn tables(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyCellTable> {
        let n = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .map(|c| {
                        c.children
                            .iter()
                            .filter(|ch| matches!(ch, Node::Elem(e) if e.name == "w:tbl"))
                            .count()
                    })
            })
            .flatten()
        })
        .unwrap_or(0);
        (0..n)
            .map(|tindex| crate::docmodel_fmt::PyCellTable {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: self.row,
                col: self.col,
                tindex,
            })
            .collect()
    }

    /// Append a rows x cols table to this cell (python-docx cell.add_table).
    fn add_table(&self, py: Python<'_>, rows: usize, cols: usize) -> PyResult<crate::docmodel_fmt::PyCellTable> {
        let tindex = self.edit(py, |c| {
            let n = c
                .children
                .iter()
                .filter(|ch| matches!(ch, Node::Elem(e) if e.name == "w:tbl"))
                .count();
            let mut tbl = Element::new("w:tbl");
            let mut grid = Element::new("w:tblGrid");
            let w = (8640usize / cols.max(1)).to_string();
            for _ in 0..cols {
                let mut gc = Element::new("w:gridCol");
                gc.set_attr("w:w", &w);
                grid.children.push(Node::Elem(gc));
            }
            tbl.children.push(Node::Elem(grid));
            for _ in 0..rows {
                let mut tr = Element::new("w:tr");
                for _ in 0..cols {
                    let mut tc = Element::new("w:tc");
                    tc.children.push(Node::Elem(Element::new("w:p")));
                    tr.children.push(Node::Elem(tc));
                }
                tbl.children.push(Node::Elem(tr));
            }
            c.children.push(Node::Elem(tbl));
            // a cell must end with a paragraph
            let last_is_p = matches!(c.children.last(), Some(Node::Elem(e)) if e.name == "w:p");
            if !last_is_p {
                c.children.push(Node::Elem(Element::new("w:p")));
            }
            n
        })?;
        Ok(crate::docmodel_fmt::PyCellTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
            tindex,
        })
    }

    /// Paragraphs and tables of this cell in document order.
    fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let kinds: Vec<bool> = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .map(|c| {
                        c.children
                            .iter()
                            .filter_map(|ch| match ch {
                                Node::Elem(e) if e.name == "w:p" => Some(true),
                                Node::Elem(e) if e.name == "w:tbl" => Some(false),
                                _ => None,
                            })
                            .collect::<Vec<_>>()
                    })
            })
            .flatten()
        })
        .unwrap_or_default();
        let mut pi = 0usize;
        let mut ti = 0usize;
        let mut out = Vec::new();
        for is_p in kinds {
            if is_p {
                if let Ok(v) = Py::new(py, crate::docmodel_fmt::PyCellParagraph {
                    tpl: self.tpl.clone_ref(py),
                    index: self.index,
                    row: self.row,
                    col: self.col,
                    para: pi,
                }) {
                    out.push(v.into_any());
                }
                pi += 1;
            } else {
                if let Ok(v) = Py::new(py, crate::docmodel_fmt::PyCellTable {
                    tpl: self.tpl.clone_ref(py),
                    index: self.index,
                    row: self.row,
                    col: self.col,
                    tindex: ti,
                }) {
                    out.push(v.into_any());
                }
                ti += 1;
            }
        }
        out
    }

    /// Merge this cell with `other` into one cell spanning the rectangular
    /// region between them (python-docx cell.merge). Returns the merged cell.
    fn merge(&self, py: Python<'_>, other: Bound<'_, PyCell>) -> PyResult<PyCell> {
        let (oidx, orow, ocol) = {
            let o = other.borrow();
            (o.index, o.row, o.col)
        };
        if oidx != self.index {
            return Err(PyValueError::new_err(
                "cannot merge cells of different tables",
            ));
        }
        let (r1, r2) = (self.row.min(orow), self.row.max(orow));
        let (c1, c2) = (self.col.min(ocol), self.col.max(ocol));
        if r1 != r2 || c1 != c2 {
            with_core(&self.tpl, py, |core| {
                let mut found = false;
                mutate_document(core, |body| {
                    if let Some(t) = nth_direct(body, "w:tbl", self.index) {
                        merge_region(t, r1, r2, c1, c2);
                        found = true;
                    }
                })
                .map_err(py_err)?;
                if !found {
                    return Err(PyValueError::new_err("table not found"));
                }
                Ok(())
            })?;
        }
        Ok(PyCell {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: r1,
            col: c1,
        })
    }
}

/// Positions of the direct w:tc children of a table row.
fn tc_positions(row: &Element) -> Vec<usize> {
    row.children
        .iter()
        .enumerate()
        .filter(|(_, c)| matches!(c, Node::Elem(e) if e.name == "w:tc"))
        .map(|(i, _)| i)
        .collect()
}

pub(crate) fn tcpr_mut(tc: &mut Element) -> &mut Element {
    if tc.find("w:tcPr").is_none() {
        tc.children.insert(0, Node::Elem(Element::new("w:tcPr")));
    }
    tc.find_mut("w:tcPr").unwrap()
}

/// Rectangular cell merge over direct tc addressing: horizontal span via
/// w:gridSpan, vertical span via w:vMerge (python-docx semantics: content of
/// every merged-away cell is moved into the top-left cell; vertical
/// continuation cells end up with a single empty paragraph).
fn merge_region(t: &mut Element, r1: usize, r2: usize, c1: usize, c2: usize) {
    // content collected from merged-away cells, in reading order
    let mut tail: Vec<Node> = Vec::new();
    for r in r1..=r2 {
        let Some(row) = nth_direct(t, "w:tr", r) else {
            continue;
        };
        let tpos = tc_positions(row);
        if c1 >= tpos.len() {
            continue;
        }
        let c2c = c2.min(tpos.len() - 1);
        // total grid span of the merged range
        let mut span: i64 = 0;
        for &p in &tpos[c1..=c2c] {
            if let Node::Elem(tc) = &row.children[p] {
                span += tc
                    .find("w:tcPr")
                    .and_then(|pr| pr.find("w:gridSpan"))
                    .and_then(|g| g.get_attr("w:val"))
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(1);
            }
        }
        // remove cells c1+1..=c2c (reverse so positions stay valid),
        // keeping their non-empty content
        let mut moved: Vec<Node> = Vec::new();
        for &p in tpos[c1 + 1..=c2c].iter().rev() {
            if let Node::Elem(tc) = row.children.remove(p) {
                for ch in tc.children {
                    if is_non_empty_block(&ch) {
                        moved.push(ch);
                    }
                }
            }
        }
        moved.reverse();
        let first_pos = tpos[c1];
        if let Node::Elem(tc) = &mut row.children[first_pos] {
            if r > r1 {
                // vertical continuation: move own content up to the origin cell
                let mut own: Vec<Node> = Vec::new();
                for ch in std::mem::take(&mut tc.children) {
                    if is_non_empty_block(&ch) {
                        own.push(ch);
                    } else if matches!(&ch, Node::Elem(e) if e.name == "w:tcPr") {
                        tc.children.push(ch);
                    }
                }
                tail.append(&mut own);
            }
            {
                let tcpr = tcpr_mut(tc);
                if span > 1 {
                    if tcpr.find("w:gridSpan").is_none() {
                        tcpr.children.insert(0, Node::Elem(Element::new("w:gridSpan")));
                    }
                    tcpr
                        .find_mut("w:gridSpan")
                        .unwrap()
                        .set_attr("w:val", &span.to_string());
                }
                if r2 > r1 {
                    if tcpr.find("w:vMerge").is_none() {
                        tcpr.children.insert(0, Node::Elem(Element::new("w:vMerge")));
                    }
                    let vm = tcpr.find_mut("w:vMerge").unwrap();
                    if r == r1 {
                        vm.set_attr("w:val", "restart");
                    } else {
                        vm.attrs.retain(|(k, _)| k != "w:val");
                    }
                }
            }
            tail.append(&mut moved);
            if r > r1 {
                // vertical-continuation cell: exactly one empty paragraph
                tc.children
                    .retain(|ch| matches!(ch, Node::Elem(e) if e.name == "w:tcPr"));
                tc.children.push(Node::Elem(Element::new("w:p")));
            }
        }
    }
    // move the collected content into the origin cell
    if tail.is_empty() {
        return;
    }
    if let Some(row) = nth_direct(t, "w:tr", r1) {
        let tpos = tc_positions(row);
        if c1 >= tpos.len() {
            return;
        }
        if let Node::Elem(tc) = &mut row.children[tpos[c1]] {
            // drop a single trailing empty paragraph (python-docx does this
            // before appending moved content)
            if matches!(
                tc.children.last(),
                Some(Node::Elem(e)) if e.name == "w:p" && element_text(e).is_empty()
            ) {
                tc.children.pop();
            }
            tc.children.append(&mut tail);
            if !tc.children.iter().any(|ch| matches!(ch, Node::Elem(e) if e.name == "w:p")) {
                tc.children.push(Node::Elem(Element::new("w:p")));
            }
        }
    }
}

/// A block element worth moving when merging: non-paragraph blocks and
/// paragraphs with text content.
fn is_non_empty_block(ch: &Node) -> bool {
    match ch {
        Node::Elem(e) if e.name == "w:tcPr" => false,
        Node::Elem(e) => e.name != "w:p" || !element_text(e).is_empty(),
        Node::Text(_) => false,
    }
}

// ---------------- sections ----------------

/// A document section (live proxy).
#[pyclass(name = "Section", unsendable)]
pub struct PySection {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

pub(crate) fn collect_sectprs_mut<'a>(el: &'a mut Element, out: &mut Vec<&'a mut Element>) {
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            if e.name == "w:sectPr" {
                out.push(e);
            } else {
                collect_sectprs_mut(e, out);
            }
        }
    }
}

impl PySection {
    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            let dom = core.document_dom().ok()?;
            let mut sects: Vec<&Element> = Vec::new();
            dom.root.iter_descendants("w:sectPr", &mut sects);
            sects.get(self.index).map(|s| f(s))
        })
    }

    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            mutate_document(core, |body| {
                let mut sects: Vec<&mut Element> = Vec::new();
                collect_sectprs_mut(body, &mut sects);
                if let Some(s) = sects.get_mut(self.index) {
                    result = Some(f(s));
                }
            })
            .map_err(py_err)?;
            result.ok_or_else(|| PyValueError::new_err("section not found"))
        })
    }
}

pub(crate) fn get_twips(sp: &Element, tag: &str, attr: &str) -> Option<i64> {
    sp.find(tag)
        .and_then(|e| e.get_attr(attr))
        .and_then(|v| v.parse::<i64>().ok())
}

pub(crate) fn set_twips(sp: &mut Element, tag: &str, attr: &str, v: Option<i64>, defaults: &[(&str, &str)]) {
    let el = match sp.find_mut(tag) {
        Some(e) => e,
        None => {
            let mut e = Element::new(tag);
            for (k, dv) in defaults {
                e.set_attr(k, dv);
            }
            sp.children.insert(0, Node::Elem(e));
            sp.find_mut(tag).unwrap()
        }
    };
    if let Some(v) = v {
        el.set_attr(attr, &v.to_string());
    }
}

pub(crate) fn to_len(v: Option<i64>) -> Option<crate::pyclasses::PyLength> {
    v.map(|t| crate::pyclasses::PyLength { emu: t * 635 })
}

pub(crate) fn from_len(obj: &Bound<'_, PyAny>) -> PyResult<Option<i64>> {
    crate::pyclasses::extract_length_pub(obj).map(|o| o.map(|emu| emu / 635))
}

trait PipeMap: Sized {
    fn pipe_map<R>(self, f: impl FnOnce(Self) -> R) -> R {
        f(self)
    }
}
impl<T> PipeMap for T {}

#[pymethods]
impl PySection {


    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    #[getter]
    fn page_width(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgSz", "w:w"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_page_width(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgSz", "w:w", twips, &[("w:w", "12240"), ("w:h", "15840")]))
    }
    #[getter]
    fn page_height(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgSz", "w:h"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_page_height(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgSz", "w:h", twips, &[("w:w", "12240"), ("w:h", "15840")]))
    }
    #[getter]
    fn left_margin(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:left"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_left_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:left", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }
    #[getter]
    fn right_margin(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:right"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_right_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:right", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }
    #[getter]
    fn top_margin(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:top"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_top_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:top", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }
    #[getter]
    fn bottom_margin(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:bottom"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_bottom_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:bottom", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }

    #[getter]
    fn orientation(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |sp| {
            sp.find("w:pgSz")
                .and_then(|e| e.get_attr("w:orient").map(|s| s.to_string()))
        })
        .flatten()
    }

    #[setter]
    fn set_orientation(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |sp| {
            // swapping orientation also swaps page dimensions (python-docx)
            let orient = if v.to_lowercase().starts_with("land") { "landscape" } else { "portrait" };
            let (w, h) = (
                get_twips(sp, "w:pgSz", "w:w"),
                get_twips(sp, "w:pgSz", "w:h"),
            );
            let cur = sp
                .find("w:pgSz")
                .and_then(|e| e.get_attr("w:orient"))
                .unwrap_or("portrait")
                .to_string();
            if cur != orient {
                if let Some(el) = sp.find_mut("w:pgSz") {
                    el.set_attr("w:orient", orient);
                    if let (Some(w), Some(h)) = (w, h) {
                        if (orient == "landscape" && w < h) || (orient == "portrait" && w > h) {
                            el.set_attr("w:w", &h.to_string());
                            el.set_attr("w:h", &w.to_string());
                        }
                    }
                }
            }
        })
    }


    #[getter]
    fn header(&self, py: Python<'_>) -> PySectionHdrFtr {
        PySectionHdrFtr {
            tpl: self.tpl.clone_ref(py),
            section: self.index,
            kind: "header".to_string(),
        }
    }

    #[getter]
    fn footer(&self, py: Python<'_>) -> PySectionHdrFtr {
        PySectionHdrFtr {
            tpl: self.tpl.clone_ref(py),
            section: self.index,
            kind: "footer".to_string(),
        }
    }

    /// Header used on even pages (python-docx even_page_header).
    /// Requires settings.odd_and_even_pages_header_footer to take effect.
    #[getter]
    fn even_page_header(&self, py: Python<'_>) -> PySectionHdrFtr {
        PySectionHdrFtr {
            tpl: self.tpl.clone_ref(py),
            section: self.index,
            kind: "even_header".to_string(),
        }
    }

    /// Footer used on even pages (python-docx even_page_footer).
    #[getter]
    fn even_page_footer(&self, py: Python<'_>) -> PySectionHdrFtr {
        PySectionHdrFtr {
            tpl: self.tpl.clone_ref(py),
            section: self.index,
            kind: "even_footer".to_string(),
        }
    }

    /// Header used on the first page (python-docx first_page_header).
    /// Requires different_first_page_header_footer to take effect.
    #[getter]
    fn first_page_header(&self, py: Python<'_>) -> PySectionHdrFtr {
        PySectionHdrFtr {
            tpl: self.tpl.clone_ref(py),
            section: self.index,
            kind: "first_header".to_string(),
        }
    }

    /// Footer used on the first page (python-docx first_page_footer).
    #[getter]
    fn first_page_footer(&self, py: Python<'_>) -> PySectionHdrFtr {
        PySectionHdrFtr {
            tpl: self.tpl.clone_ref(py),
            section: self.index,
            kind: "first_footer".to_string(),
        }
    }

    #[getter]
    fn different_first_page_header_footer(&self, py: Python<'_>) -> bool {
        self.read(py, |sp| sp.find("w:titlePg").is_some())
            .unwrap_or(false)
    }

    #[setter]
    fn set_different_first_page_header_footer(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |sp| {
            if v {
                if sp.find("w:titlePg").is_none() {
                    sp.children.insert(0, Node::Elem(Element::new("w:titlePg")));
                }
            } else {
                sp.children
                    .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:titlePg"));
            }
        })
    }

    /// Section start type as a WD_SECTION_START int (0=continuous,
    /// 1=nextColumn, 2=nextPage, 3=evenPage, 4=oddPage; xml name also
    /// accepted on set). Missing w:type reads as 2 (next page).
    #[getter]
    fn start_type(&self, py: Python<'_>) -> i64 {
        self.read(py, |sp| {
            sp.find("w:type")
                .and_then(|e| e.get_attr("w:val"))
                .map(|s| match s {
                    "continuous" => 0,
                    "nextColumn" => 1,
                    "evenPage" => 3,
                    "oddPage" => 4,
                    _ => 2,
                })
                .unwrap_or(2)
        })
        .unwrap_or(2)
    }
    #[setter]
    fn set_start_type(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let val: Option<&'static str> = if v.is_none() {
            None
        } else if let Ok(i) = v.extract::<i64>() {
            Some(match i {
                0 => "continuous",
                1 => "nextColumn",
                3 => "evenPage",
                4 => "oddPage",
                _ => "nextPage",
            })
        } else {
            let s: String = v.extract()?;
            Some(match s.as_str() {
                "continuous" => "continuous",
                "nextColumn" => "nextColumn",
                "evenPage" => "evenPage",
                "oddPage" => "oddPage",
                _ => "nextPage",
            })
        };
        self.edit(py, |sp| match val {
            // nextPage is the default: drop the element (python-docx)
            None | Some("nextPage") => sp
                .children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:type")),
            Some(x) => {
                let t = crate::docmodel_fmt::ensure_child(sp, "w:type");
                t.set_attr("w:val", x);
            }
        })
    }

    #[getter]
    fn header_distance(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:header"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_header_distance(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:header", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }
    #[getter]
    fn footer_distance(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:footer"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_footer_distance(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:footer", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }
    #[getter]
    fn gutter(&self, py: Python<'_>) -> Option<crate::pyclasses::PyLength> {
        self.read(py, |sp| get_twips(sp, "w:pgMar", "w:gutter"))
            .flatten()
            .pipe_map(to_len)
    }
    #[setter]
    fn set_gutter(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let twips = from_len(v)?;
        self.edit(py, |sp| set_twips(sp, "w:pgMar", "w:gutter", twips, &[("w:left", "1800"), ("w:right", "1800"), ("w:top", "1440"), ("w:bottom", "1440")]))
    }

    /// Paragraphs and tables of this section in document order. Section
    /// boundaries are the paragraphs carrying a paragraph-level sectPr; the
    /// last section ends at the body-level sectPr.
    pub fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let items: Vec<(bool, usize)> = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                // section ranges: (is_paragraph, tag_index) per section
                let mut sections: Vec<Vec<(bool, usize)>> = vec![Vec::new()];
                let mut pi = 0usize;
                let mut ti = 0usize;
                for c in &body.children {
                    let Node::Elem(e) = c else { continue };
                    match e.name.as_str() {
                        "w:p" => {
                            sections.last_mut().unwrap().push((true, pi));
                            pi += 1;
                            let ends_section = e
                                .find("w:pPr")
                                .and_then(|ppr| ppr.find("w:sectPr"))
                                .is_some();
                            if ends_section {
                                sections.push(Vec::new());
                            }
                        }
                        "w:tbl" => {
                            sections.last_mut().unwrap().push((false, ti));
                            ti += 1;
                        }
                        _ => {}
                    }
                }
                sections.get(self.index).cloned().unwrap_or_default()
            })
            .unwrap_or_default()
        });
        let mut out = Vec::new();
        for (is_p, idx) in items {
            if is_p {
                if let Ok(v) = Py::new(py, PyParagraph { tpl: self.tpl.clone_ref(py), index: idx }) {
                    out.push(v.into_any());
                }
            } else if let Ok(v) = Py::new(py, PyTable { tpl: self.tpl.clone_ref(py), index: idx }) {
                out.push(v.into_any());
            }
        }
        out
    }
}

/// Header or footer of a section.
#[pyclass(name = "SectionHdrFtr", unsendable)]
pub struct PySectionHdrFtr {
    pub tpl: Py<PyDocxTemplate>,
    pub section: usize,
    pub kind: String, // "header" | "footer"
}

fn rel_type_for(kind: &str) -> &'static str {
    let (_, base) = split_kind(kind);
    if base == "header" {
        crate::package::rel_type::HEADER
    } else {
        crate::package::rel_type::FOOTER
    }
}

/// Split a hdrftr kind into (w:type value, base "header"|"footer").
/// Kinds: header/footer (default), even_header/even_footer, first_header/first_footer.
fn split_kind(kind: &str) -> (&'static str, &'static str) {
    match kind {
        "footer" => ("default", "footer"),
        "even_header" => ("even", "header"),
        "even_footer" => ("even", "footer"),
        "first_header" => ("first", "header"),
        "first_footer" => ("first", "footer"),
        _ => ("default", "header"),
    }
}

/// find the header/footer part path linked to a section (by headerReference
/// order within its sectPr), if any
fn find_hdrftr_part(
    core: &mut TplCore,
    section_idx: usize,
    kind: &str,
) -> Option<(String, String)> {
    // returns (rid, part_path)
    let rid = {
        let dom = core.document_dom().ok()?;
        let mut sects: Vec<&Element> = Vec::new();
        dom.root.iter_descendants("w:sectPr", &mut sects);
        let sect = sects.get(section_idx)?;
        let (wtype, base) = split_kind(kind);
        let want_tag = if base == "header" {
            "w:headerReference"
        } else {
            "w:footerReference"
        };
        // use the reference with the matching w:type
        let mut rid: Option<String> = None;
        for c in &sect.children {
            if let Node::Elem(e) = c {
                if e.name == want_tag {
                    let t = e.get_attr("w:type").unwrap_or("default");
                    if t == wtype && rid.is_none() {
                        rid = e.get_attr("r:id").map(|s| s.to_string());
                    }
                }
            }
        }
        rid?
    };
    let pkg = core.package.as_ref()?;
    let rels = pkg.rels(DOCUMENT_PART);
    let rel = rels.get(&rid)?;
    Some((
        rid,
        crate::package::resolve_target(DOCUMENT_PART, &rel.target),
    ))
}

#[pymethods]
impl PySectionHdrFtr {
    #[getter]
    fn is_linked_to_previous(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| find_hdrftr_part(core, self.section, &self.kind).is_none())
    }

    #[setter]
    fn set_is_linked_to_previous(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| -> Result<(), String> {
            if v {
                // unlink: remove the reference (and leave the part)
                remove_hdrftr_reference(core, self.section, &self.kind)?;
            } else {
                // link: create a new empty part if none exists
                if find_hdrftr_part(core, self.section, &self.kind).is_none() {
                    create_hdrftr_part(core, self.section, &self.kind)?;
                }
            }
            Ok(())
        })
        .map_err(py_err)
    }

    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Vec<String> {
        with_core(&self.tpl, py, |core| {
            match find_hdrftr_part(core, self.section, &self.kind) {
                Some((_, part)) => {
                    let xml = core
                        .package
                        .as_ref()
                        .and_then(|p| p.get_string(&part));
                    xml.and_then(|x| Document::parse(&x).ok())
                        .map(|dom| {
                            let mut out = Vec::new();
                            let mut paras: Vec<&Element> = Vec::new();
                            dom.root.iter_descendants("w:p", &mut paras);
                            for p in paras {
                                out.push(element_text(p));
                            }
                            out
                        })
                        .unwrap_or_default()
                }
                None => Vec::new(),
            }
        })
    }

    #[pyo3(signature = (text=""))]
    fn add_paragraph(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        with_core(&self.tpl, py, |core| -> Result<(), String> {
            if find_hdrftr_part(core, self.section, &self.kind).is_none() {
                create_hdrftr_part(core, self.section, &self.kind)?;
            }
            let (_, part) = find_hdrftr_part(core, self.section, &self.kind)
                .ok_or("cannot create header/footer part")?;
            let pkg = core.package.as_mut().ok_or("package not loaded")?;
            let mut xml = pkg.get_string(&part).unwrap_or_default();
            let para = format!(
                "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                crate::richtext::html_escape(text)
            );
            let close = format!("</w:{}>", if split_kind(&self.kind).1 == "header" { "hdr" } else { "ftr" });
            if let Some(pos) = xml.rfind(&close) {
                xml.insert_str(pos, &para);
            } else {
                return Err("invalid header/footer part".into());
            }
            let enc = pkg.encoding_of(&part);
            pkg.set(&part, crate::package::encode_part_owned(xml, &enc));
            Ok(())
        })
        .map_err(py_err)
    }
}

fn remove_hdrftr_reference(core: &mut TplCore, section_idx: usize, kind: &str) -> Result<(), String> {
    mutate_document(core, |body| {
        let mut sects: Vec<&mut Element> = Vec::new();
        fn collect<'a>(el: &'a mut Element, out: &mut Vec<&'a mut Element>) {
            for c in el.children.iter_mut() {
                if let Node::Elem(e) = c {
                    if e.name == "w:sectPr" {
                        out.push(e);
                    } else {
                        collect(e, out);
                    }
                }
            }
        }
        collect(body, &mut sects);
        if let Some(sect) = sects.get_mut(section_idx) {
            let (wtype, base) = split_kind(kind);
            let want = if base == "header" {
                "w:headerReference"
            } else {
                "w:footerReference"
            };
            sect.children.retain(|c| {
                !(matches!(c, Node::Elem(e) if e.name == want && e.get_attr("w:type").unwrap_or("default") == wtype))
            });
        }
    })
}

fn create_hdrftr_part(core: &mut TplCore, section_idx: usize, kind: &str) -> Result<(), String> {
    let (wtype, base) = split_kind(kind);
    let pkg = core.package.as_mut().ok_or("package not loaded")?;
    // find next part number
    let prefix = if base == "header" { "word/header" } else { "word/footer" };
    let mut n = 1;
    while pkg.contains(&format!("{}{}.xml", prefix, n)) {
        n += 1;
    }
    let part = format!("{}{}.xml", prefix, n);
    let root = if base == "header" { "hdr" } else { "ftr" };
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:{root} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:{root}>"
    );
    pkg.set(&part, xml.into_bytes());
    let ct = if base == "header" {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    } else {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    };
    pkg.ensure_content_type_override(&part, ct);
    let target = crate::package::relative_target(DOCUMENT_PART, &part);
    let rid = pkg.add_rel(DOCUMENT_PART, rel_type_for(kind), &target, false);

    // add the reference to the section
    let tag = if base == "header" {
        "w:headerReference"
    } else {
        "w:footerReference"
    };
    mutate_document(core, |body| {
        let mut sects: Vec<&mut Element> = Vec::new();
        fn collect<'a>(el: &'a mut Element, out: &mut Vec<&'a mut Element>) {
            for c in el.children.iter_mut() {
                if let Node::Elem(e) = c {
                    if e.name == "w:sectPr" {
                        out.push(e);
                    } else {
                        collect(e, out);
                    }
                }
            }
        }
        collect(body, &mut sects);
        if let Some(sect) = sects.get_mut(section_idx) {
            let mut el = Element::new(tag);
            el.set_attr("w:type", wtype);
            el.set_attr("r:id", &rid);
            sect.children.insert(0, Node::Elem(el));
        }
    })
}

// ---------------- styles ----------------

/// A style in the document (live proxy).
#[pyclass(name = "Style", unsendable)]
pub struct PyStyle {
    pub tpl: Py<PyDocxTemplate>,
    pub style_id: String,
}

pub(crate) fn with_styles<R>(core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
    ensure_styles_part(core)?;
    let r = f(&mut core.part_dom("word/styles.xml")?.root);
    core.mark_part_dirty("word/styles.xml");
    Ok(r)
}

pub(crate) fn ensure_styles_part(core: &mut TplCore) -> Result<(), String> {
    core.init_docx(false)?;
    let pkg = core.package.as_mut().ok_or("package not loaded")?;
    if pkg.contains("word/styles.xml") {
        return Ok(());
    }
    let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>";
    pkg.set("word/styles.xml", xml.as_bytes().to_vec());
    pkg.ensure_content_type_override(
        "word/styles.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    );
    if pkg.rels(DOCUMENT_PART).by_type(crate::package::rel_type::STYLES).next().is_none() {
        pkg.add_rel(DOCUMENT_PART, crate::package::rel_type::STYLES, "styles.xml", false);
    }
    Ok(())
}

pub(crate) fn find_style_el<'a>(root: &'a Element, style_id: &str) -> Option<&'a Element> {
    fn walk<'a>(el: &'a Element, style_id: &str) -> Option<&'a Element> {
        for c in &el.children {
            if let Node::Elem(e) = c {
                if e.name == "w:style" && e.get_attr("w:styleId") == Some(style_id) {
                    return Some(e);
                }
                if let Some(r) = walk(e, style_id) {
                    return Some(r);
                }
            }
        }
        None
    }
    walk(root, style_id)
}

pub(crate) fn find_style_el_mut<'a>(root: &'a mut Element, style_id: &str) -> Option<&'a mut Element> {
    for c in root.children.iter_mut() {
        if let Node::Elem(e) = c {
            if e.name == "w:style" && e.get_attr("w:styleId") == Some(style_id) {
                return Some(e);
            }
            if let Some(r) = find_style_el_mut(e, style_id) {
                return Some(r);
            }
        }
    }
    None
}

impl PyStyle {
    pub(crate) fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            with_styles(core, |root| {
                if let Some(st) = find_style_el_mut(root, &self.style_id) {
                    result = Some(f(st));
                }
            })
            .map_err(py_err)?;
            result.ok_or_else(|| PyValueError::new_err("style not found"))
        })
    }

    pub(crate) fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            let dom = core.part_dom("word/styles.xml").ok()?;
            find_style_el(&dom.root, &self.style_id).map(|e| f(e))
        })
    }
}

fn style_name_of(el: &Element) -> Option<String> {
    el.find("w:name")
        .and_then(|n| n.get_attr("w:val").map(|s| s.to_string()))
}

// ---------------- Document facade ----------------

/// A python-docx-inspired facade over the template's current document.
#[pyclass(name = "Document", unsendable)]
pub struct PyDocument {
    pub tpl: Py<PyDocxTemplate>,
}

#[pymethods]
impl PyDocument {
    /// Raw XML root element of word/document.xml (live proxy), the
    /// python-docx `document.element` escape hatch.
    #[getter]
    pub fn element(&self, py: Python<'_>) -> crate::pyxml::PyXmlElement {
        crate::pyxml::PyXmlElement {
            tpl: self.tpl.clone_ref(py),
            part: DOCUMENT_PART.to_string(),
            path: Vec::new(),
        }
    }


    /// The package part this object belongs to (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: DOCUMENT_PART.to_string(),
        }
    }

    #[getter]
    pub fn paragraphs(&self, py: Python<'_>) -> Vec<PyParagraph> {
        let n = with_core(&self.tpl, py, |core| count_in_body(core, "w:p"));
        (0..n)
            .map(|i| PyParagraph {
                tpl: self.tpl.clone_ref(py),
                index: i,
            })
            .collect()
    }

    #[getter]
    pub fn tables(&self, py: Python<'_>) -> Vec<PyTable> {
        let n = with_core(&self.tpl, py, |core| count_in_body(core, "w:tbl"));
        (0..n)
            .map(|i| PyTable {
                tpl: self.tpl.clone_ref(py),
                index: i,
            })
            .collect()
    }

    /// Paragraphs and tables of the document body in document order
    /// (python-docx iter_inner_content).
    pub fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let kinds: Vec<bool> = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                body.children
                    .iter()
                    .filter_map(|c| match c {
                        Node::Elem(e) if e.name == "w:p" => Some(true),
                        Node::Elem(e) if e.name == "w:tbl" => Some(false),
                        _ => None,
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default()
        });
        let mut pi = 0usize;
        let mut ti = 0usize;
        let mut out = Vec::new();
        for is_p in kinds {
            if is_p {
                if let Ok(v) = Py::new(py, PyParagraph { tpl: self.tpl.clone_ref(py), index: pi }) {
                    out.push(v.into_any());
                }
                pi += 1;
            } else {
                if let Ok(v) = Py::new(py, PyTable { tpl: self.tpl.clone_ref(py), index: ti }) {
                    out.push(v.into_any());
                }
                ti += 1;
            }
        }
        out
    }

    #[getter]
    pub fn sections(&self, py: Python<'_>) -> Vec<PySection> {
        let n = with_core(&self.tpl, py, |core| {
            core.document_dom()
                .map(|dom| {
                    let mut v: Vec<&Element> = Vec::new();
                    dom.root.iter_descendants("w:sectPr", &mut v);
                    v.len()
                })
                .unwrap_or(0)
        });
        (0..n)
            .map(|i| PySection {
                tpl: self.tpl.clone_ref(py),
                index: i,
            })
            .collect()
    }

    #[getter]
    pub fn styles(&self, py: Python<'_>) -> PyStyles {
        PyStyles {
            tpl: self.tpl.clone_ref(py),
        }
    }

    #[getter]
    pub fn settings(&self, py: Python<'_>) -> PySettings {
        PySettings {
            tpl: self.tpl.clone_ref(py),
        }
    }

    #[getter]
    pub fn core_properties(&self, py: Python<'_>) -> PyCoreProperties {
        PyCoreProperties {
            tpl: self.tpl.clone_ref(py),
        }
    }

    #[getter]
    pub fn comments(&self, py: Python<'_>) -> crate::doccomments::PyComments {
        crate::doccomments::PyComments {
            tpl: self.tpl.clone_ref(py),
        }
    }

    #[getter]
    pub fn inline_shapes(&self, py: Python<'_>) -> Vec<PyInlineShape> {
        with_core(&self.tpl, py, |core| {
            core.document_dom()
                .map(|doc| {
                    let mut out = Vec::new();
                    let mut inlines: Vec<&Element> = Vec::new();
                    doc.root.iter_descendants("wp:inline", &mut inlines);
                    for il in inlines {
                        let mut cx = 0;
                        let mut cy = 0;
                        if let Some(ext) = il.find("wp:extent") {
                            cx = ext.get_attr("cx").and_then(|v| v.parse().ok()).unwrap_or(0);
                            cy = ext.get_attr("cy").and_then(|v| v.parse().ok()).unwrap_or(0);
                        }
                        out.push(PyInlineShape {
                            width: crate::pyclasses::PyLength { emu: cx },
                            height: crate::pyclasses::PyLength { emu: cy },
                            kind: "picture".to_string(),
                        });
                    }
                    out
                })
                .unwrap_or_default()
        })
    }

    pub fn save(&self, py: Python<'_>, filename: &Bound<'_, PyAny>) -> PyResult<()> {
        self.tpl.bind(py).borrow().save(py, filename)
    }

    /// Append a paragraph to the document (python-docx add_paragraph).
    #[pyo3(signature = (text="", style=None))]
    pub fn add_paragraph(
        &self,
        py: Python<'_>,
        text: &str,
        style: Option<String>,
    ) -> PyResult<PyParagraph> {
        crate::docmodel_add::doc_add_paragraph(self, py, text, style)
    }

    /// Append a heading paragraph (python-docx add_heading).
    #[pyo3(signature = (text="", level=1))]
    pub fn add_heading(
        &self,
        py: Python<'_>,
        text: &str,
        level: u32,
    ) -> PyResult<PyParagraph> {
        crate::docmodel_add::doc_add_heading(self, py, text, level)
    }

    /// Append a picture paragraph (python-docx add_picture).
    #[pyo3(signature = (image_descriptor, width=None, height=None))]
    pub fn add_picture(
        &self,
        py: Python<'_>,
        image_descriptor: &Bound<'_, PyAny>,
        width: Option<&Bound<'_, PyAny>>,
        height: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let w = width.map(|v| crate::pyclasses::extract_length_pub(v)).transpose()?;
        let h = height.map(|v| crate::pyclasses::extract_length_pub(v)).transpose()?;
        crate::docmodel_add::doc_add_picture(self, py, image_descriptor, w.flatten(), h.flatten())
    }

    /// Append a table (python-docx add_table).
    pub fn add_table(
        &self,
        py: Python<'_>,
        rows: usize,
        cols: usize,
    ) -> PyResult<PyTable> {
        crate::docmodel_add::doc_add_table(self, py, rows, cols)
    }

    /// Append a page break paragraph (python-docx add_page_break).
    pub fn add_page_break(&self, py: Python<'_>) -> PyResult<()> {
        crate::docmodel_add::doc_add_page_break(self, py)
    }

    /// Add a new section (python-docx add_section).
    #[pyo3(signature = (start_type=2))]
    pub fn add_section(&self, py: Python<'_>, start_type: u32) -> PyResult<PySection> {
        crate::docmodel_add::doc_add_section(self, py, start_type)
    }

    /// Add a comment anchored to the given runs (python-docx add_comment).
    #[pyo3(signature = (runs, text="", author="", initials=""))]
    pub fn add_comment(
        &self,
        py: Python<'_>,
        runs: &Bound<'_, PyAny>,
        text: &str,
        author: &str,
        initials: &str,
    ) -> PyResult<crate::doccomments::PyComment> {
        crate::doccomments::doc_add_comment(self, py, runs, text, author, initials)
    }
}

#[pymethods]
impl PyStyle {

    /// The styles package part (minimal facade).
    #[getter]
    fn part(&self, py: Python<'_>) -> crate::docmodel_fmt::PyPart {
        crate::docmodel_fmt::PyPart {
            tpl: self.tpl.clone_ref(py),
            part_name: "word/styles.xml".to_string(),
        }
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |st| style_name_of(st)).flatten()
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |st| {
            if let Some(n) = st.find_mut("w:name") {
                n.set_attr("w:val", &v);
            } else {
                let mut n = Element::new("w:name");
                n.set_attr("w:val", &v);
                st.children.insert(0, Node::Elem(n));
            }
        })
    }

    #[getter]
    fn style_id(&self) -> String {
        self.style_id.clone()
    }

    #[getter]
    fn style_type(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |st| st.get_attr("w:type").map(|s| s.to_string()))
            .flatten()
    }

    #[getter]
    fn base_style(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |st| {
            st.find("w:basedOn")
                .and_then(|b| b.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }

    #[setter]
    fn set_base_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |st| {
            if let Some(b) = st.find_mut("w:basedOn") {
                b.set_attr("w:val", &v);
            } else {
                let mut b = Element::new("w:basedOn");
                b.set_attr("w:val", &v);
                st.children.push(Node::Elem(b));
            }
        })
    }

    /// Full python-docx Font facade over the style's w:rPr.
    #[getter]
    fn font(&self, py: Python<'_>) -> crate::docmodel_fmt::PyFont {
        crate::docmodel_fmt::PyFont {
            tpl: self.tpl.clone_ref(py),
            target: crate::docmodel_fmt::FontTarget::Style {
                style_id: self.style_id.clone(),
            },
        }
    }

    /// Paragraph formatting of the style (python-docx
    /// ParagraphStyle.paragraph_format).
    #[getter]
    fn paragraph_format(&self, py: Python<'_>) -> crate::docmodel_fmt::PyParagraphFormat {
        crate::docmodel_fmt::PyParagraphFormat {
            tpl: self.tpl.clone_ref(py),
            target: crate::docmodel_fmt::PfTarget::Style {
                style_id: self.style_id.clone(),
            },
        }
    }

    /// Hidden in the UI until used (w:semiHidden; python-docx Style.hidden).
    #[getter]
    fn hidden(&self, py: Python<'_>) -> bool {
        self.read(py, |st| st.find("w:semiHidden").is_some())
            .unwrap_or(false)
    }
    #[setter]
    fn set_hidden(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |st| style_flag(st, "w:semiHidden", v))
    }

    /// Locked against editing (w:locked).
    #[getter]
    fn locked(&self, py: Python<'_>) -> bool {
        self.read(py, |st| st.find("w:locked").is_some())
            .unwrap_or(false)
    }
    #[setter]
    fn set_locked(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |st| style_flag(st, "w:locked", v))
    }

    /// Shown in the quick style gallery (w:qFormat).
    #[getter]
    fn quick_style(&self, py: Python<'_>) -> bool {
        self.read(py, |st| st.find("w:qFormat").is_some())
            .unwrap_or(false)
    }
    #[setter]
    fn set_quick_style(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |st| style_flag(st, "w:qFormat", v))
    }

    /// Re-hide when the style is no longer used (w:unhideWhenUsed).
    #[getter]
    fn unhide_when_used(&self, py: Python<'_>) -> bool {
        self.read(py, |st| st.find("w:unhideWhenUsed").is_some())
            .unwrap_or(false)
    }
    #[setter]
    fn set_unhide_when_used(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |st| style_flag(st, "w:unhideWhenUsed", v))
    }

    /// UI priority (w:uiPriority w:val); None removes it.
    #[getter]
    fn priority(&self, py: Python<'_>) -> Option<i64> {
        self.read(py, |st| {
            st.find("w:uiPriority")
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
        })
        .flatten()
    }
    #[setter]
    fn set_priority(&self, py: Python<'_>, v: Option<i64>) -> PyResult<()> {
        self.edit(py, |st| match v {
            Some(n) => {
                let e = crate::docmodel_fmt::ensure_child(st, "w:uiPriority");
                e.set_attr("w:val", &n.to_string());
            }
            None => st
                .children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:uiPriority")),
        })
    }

    /// Builtin styles lack the w:customStyle attribute (read-only).
    #[getter]
    fn builtin(&self, py: Python<'_>) -> bool {
        self.read(py, |st| {
            !matches!(st.get_attr("w:customStyle"), Some("1") | Some("true") | Some("on"))
        })
        .unwrap_or(true)
    }

    /// Style applied to the next paragraph (w:next; paragraph styles).
    #[getter]
    fn next_paragraph_style(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |st| {
            st.find("w:next")
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_next_paragraph_style(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        if v.is_none() {
            return self.edit(py, |st| {
                st.children
                    .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:next"));
            });
        }
        let name: String = v.extract()?;
        let sid = with_core(&self.tpl, py, |core| {
            crate::subdocbuilder::resolve_style_id(core, &name)
        });
        self.edit(py, |st| {
            let e = crate::docmodel_fmt::ensure_child(st, "w:next");
            e.set_attr("w:val", &sid);
        })
    }

    fn delete(&self, py: Python<'_>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| {
            with_styles(core, |root| {
                root.children.retain(|c| {
                    !(matches!(c, Node::Elem(e) if e.name == "w:style" && e.get_attr("w:styleId") == Some(self.style_id.as_str())))
                });
            })
            .map_err(py_err)
        })
    }
}

/// on/off child element of a style (missing == off).
fn style_flag(st: &mut Element, tag: &str, on: bool) {
    let exists = st.find(tag).is_some();
    if on && !exists {
        st.children.push(Node::Elem(Element::new(tag)));
    } else if !on && exists {
        st.children
            .retain(|c| !matches!(c, Node::Elem(e) if e.name == tag));
    }
}

/// Font properties of a style.
#[pyclass(name = "StyleFont", unsendable)]
pub struct PyStyleFont {
    pub tpl: Py<PyDocxTemplate>,
    pub style_id: String,
}

impl PyStyleFont {
    fn edit_rpr<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        let st = PyStyle {
            tpl: self.tpl.clone_ref(py),
            style_id: self.style_id.clone(),
        };
        st.edit(py, |el| {
            let rpr = {
                let has = el.children.iter().any(|c| matches!(c, Node::Elem(e) if e.name == "w:rPr"));
                if !has {
                    el.children.push(Node::Elem(Element::new("w:rPr")));
                }
                el.find_mut("w:rPr").unwrap()
            };
            f(rpr)
        })
    }

    fn read_rpr<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let st = PyStyle {
            tpl: self.tpl.clone_ref(py),
            style_id: self.style_id.clone(),
        };
        st.read(py, |el| el.find("w:rPr").map(|r| f(r))).flatten()
    }

    fn set_rpr_val(rpr: &mut Element, tag: &str, attr: &str, val: &str) {
        if let Some(el) = rpr.find_mut(tag) {
            el.set_attr(attr, val);
        } else {
            let mut el = Element::new(tag);
            el.set_attr(attr, val);
            rpr.children.push(Node::Elem(el));
        }
    }

    fn set_rpr_flag(rpr: &mut Element, tag: &str, on: bool) {
        let exists = rpr.children.iter().any(|c| matches!(c, Node::Elem(e) if e.name == tag));
        if on && !exists {
            rpr.children.push(Node::Elem(Element::new(tag)));
        } else if !on && exists {
            rpr.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == tag));
        }
    }
}

#[pymethods]
impl PyStyleFont {

    #[getter]
    fn bold(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:b").map(|e| {
                !matches!(e.get_attr("w:val"), Some("0") | Some("false") | Some("off"))
            })
        })
        .flatten()
    }
    #[setter]
    fn set_bold(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit_rpr(py, |rpr| Self::set_rpr_flag(rpr, "w:b", v))
    }
    #[getter]
    fn italic(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:i").map(|e| {
                !matches!(e.get_attr("w:val"), Some("0") | Some("false") | Some("off"))
            })
        })
        .flatten()
    }
    #[setter]
    fn set_italic(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit_rpr(py, |rpr| Self::set_rpr_flag(rpr, "w:i", v))
    }
    #[getter]
    fn all_caps(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:caps").map(|e| {
                !matches!(e.get_attr("w:val"), Some("0") | Some("false") | Some("off"))
            })
        })
        .flatten()
    }
    #[setter]
    fn set_all_caps(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit_rpr(py, |rpr| Self::set_rpr_flag(rpr, "w:caps", v))
    }
    #[getter]
    fn small_caps(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:smallCaps").map(|e| {
                !matches!(e.get_attr("w:val"), Some("0") | Some("false") | Some("off"))
            })
        })
        .flatten()
    }
    #[setter]
    fn set_small_caps(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit_rpr(py, |rpr| Self::set_rpr_flag(rpr, "w:smallCaps", v))
    }
    #[getter]
    fn strike(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:strike").map(|e| {
                !matches!(e.get_attr("w:val"), Some("0") | Some("false") | Some("off"))
            })
        })
        .flatten()
    }
    #[setter]
    fn set_strike(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit_rpr(py, |rpr| Self::set_rpr_flag(rpr, "w:strike", v))
    }

    #[getter]
    fn underline(&self, py: Python<'_>) -> Option<String> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:u")
                .and_then(|u| u.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_underline(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        if let Ok(s) = v.extract::<String>() {
            self.edit_rpr(py, |rpr| Self::set_rpr_val(rpr, "w:u", "w:val", &s))
        } else if v.is_truthy()? {
            self.edit_rpr(py, |rpr| Self::set_rpr_val(rpr, "w:u", "w:val", "single"))
        } else {
            self.edit_rpr(py, |rpr| Self::set_rpr_flag(rpr, "w:u", false))
        }
    }

    #[getter]
    fn size(&self, py: Python<'_>) -> Option<u32> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:sz")
                .and_then(|e| e.get_attr("w:val").and_then(|v| v.parse().ok()))
        })
        .flatten()
    }
    #[setter]
    fn set_size(&self, py: Python<'_>, v: u32) -> PyResult<()> {
        let s = v.to_string();
        self.edit_rpr(py, |rpr| {
            Self::set_rpr_val(rpr, "w:sz", "w:val", &s);
            Self::set_rpr_val(rpr, "w:szCs", "w:val", &s);
        })
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> Option<String> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:color")
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_color(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let c = v.strip_prefix('#').unwrap_or(&v).to_string();
        self.edit_rpr(py, |rpr| Self::set_rpr_val(rpr, "w:color", "w:val", &c))
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> Option<String> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:rFonts")
                .and_then(|e| e.get_attr("w:ascii").map(|s| s.to_string()))
        })
        .flatten()
    }
    #[setter]
    fn set_name(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit_rpr(py, |rpr| {
            Self::set_rpr_val(rpr, "w:rFonts", "w:ascii", &v);
            Self::set_rpr_val(rpr, "w:rFonts", "w:hAnsi", &v);
            Self::set_rpr_val(rpr, "w:rFonts", "w:cs", &v);
        })
    }
}

/// The styles collection.
#[pyclass(name = "Styles", unsendable)]
pub struct PyStyles {
    pub tpl: Py<PyDocxTemplate>,
}

#[pymethods]
impl PyStyles {
    /// Raw XML root element of word/styles.xml (live proxy), the
    /// python-docx `styles.element` escape hatch.
    #[getter]
    fn element(&self, py: Python<'_>) -> PyResult<crate::pyxml::PyXmlElement> {
        with_core(&self.tpl, py, |core| ensure_styles_part(core)).map_err(py_err)?;
        Ok(crate::pyxml::PyXmlElement {
            tpl: self.tpl.clone_ref(py),
            part: "word/styles.xml".to_string(),
            path: Vec::new(),
        })
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let styles = self.style_list(py);
        let list = pyo3::types::PyList::new(py, styles)?;
        Ok(list.call_method0("__iter__")?.unbind())
    }

    fn style_list(&self, py: Python<'_>) -> Vec<PyStyle> {
        with_core(&self.tpl, py, |core| {
            core
                .part_dom("word/styles.xml")
                .map(|dom| {
                    let mut out = Vec::new();
                    fn walk(el: &Element, out: &mut Vec<String>) {
                        for c in &el.children {
                            if let Node::Elem(e) = c {
                                if e.name == "w:style" {
                                    if let Some(id) = e.get_attr("w:styleId") {
                                        out.push(id.to_string());
                                    }
                                }
                                walk(e, out);
                            }
                        }
                    }
                    walk(&dom.root, &mut out);
                    out
                })
                .unwrap_or_default()
                .into_iter()
                .map(|id| PyStyle {
                    tpl: self.tpl.clone_ref(py),
                    style_id: id,
                })
                .collect()
        })
    }

    fn __len__(&self, py: Python<'_>) -> usize {
        self.style_list(py).len()
    }

    /// Get a style by name or id (python-docx styles[name]).
    fn __getitem__(&self, py: Python<'_>, key: String) -> PyResult<PyStyle> {
        with_core(&self.tpl, py, |core| {
            core.init_docx(false).map_err(PyRuntimeError::new_err)?;
            let sid = crate::subdocbuilder::resolve_style_id(core, &key);
            // verify it exists (cached DOM)
            let found = core
                .part_dom("word/styles.xml")
                .map(|dom| find_style_el(&dom.root, &sid).is_some())
                .unwrap_or(false);
            if !found {
                return Err(PyValueError::new_err(format!("no style with name '{}'", key)));
            }
            Ok(PyStyle {
                tpl: self.tpl.clone_ref(py),
                style_id: sid,
            })
        })
    }

    /// Create a new style (python-docx styles.add_style).
    /// type: 1=paragraph, 2=character, 3=table, 4=list
    #[pyo3(signature = (name, style_type=1, builtin=false))]
    fn add_style(&self, py: Python<'_>, name: &str, style_type: u32, builtin: bool) -> PyResult<PyStyle> {
        let type_str = match style_type {
            1 => "paragraph",
            2 => "character",
            3 => "table",
            4 => "numbering",
            _ => return Err(PyValueError::new_err("invalid style type")),
        };
        let style_id = with_core(&self.tpl, py, |core| {
            with_styles(core, |root| {
                let mut base = String::new();
                for c in name.chars() {
                    if c.is_alphanumeric() {
                        base.push(c);
                    }
                }
                if base.is_empty() {
                    base = "Style".to_string();
                }
                let mut id = base.clone();
                let mut n = 1;
                while find_style_el(root, &id).is_some() {
                    n += 1;
                    id = format!("{}{}", base, n);
                }
                let mut el = Element::new("w:style");
                el.set_attr("w:type", type_str);
                if builtin {
                    el.set_attr("w:customStyle", "0");
                } else {
                    el.set_attr("w:customStyle", "1");
                }
                el.set_attr("w:styleId", &id);
                let mut nm = Element::new("w:name");
                nm.set_attr("w:val", name);
                el.children.push(Node::Elem(nm));
                root.children.push(Node::Elem(el));
                id
            })
            .map_err(py_err)
        })?;
        Ok(PyStyle {
            tpl: self.tpl.clone_ref(py),
            style_id,
        })
    }
}

// ---------------- settings ----------------

/// Document settings (word/settings.xml).
#[pyclass(name = "Settings", unsendable)]
pub struct PySettings {
    pub tpl: Py<PyDocxTemplate>,
}

pub(crate) fn ensure_settings_part(core: &mut TplCore) -> Result<(), String> {
    core.init_docx(false)?;
    let pkg = core.package.as_mut().ok_or("package not loaded")?;
    if !pkg.contains("word/settings.xml") {
        let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>";
        pkg.set("word/settings.xml", xml.as_bytes().to_vec());
        pkg.ensure_content_type_override(
            "word/settings.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        );
        if pkg.rels(DOCUMENT_PART).by_type(crate::package::rel_type::SETTINGS).next().is_none() {
            pkg.add_rel(DOCUMENT_PART, crate::package::rel_type::SETTINGS, "settings.xml", false);
        }
    }
    Ok(())
}

fn with_settings<R>(core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
    ensure_settings_part(core)?;
    let r = f(&mut core.part_dom("word/settings.xml")?.root);
    core.mark_part_dirty("word/settings.xml");
    Ok(r)
}

#[pymethods]
impl PySettings {
    /// Raw XML root element of word/settings.xml (live proxy), the
    /// python-docx `settings.element` escape hatch.
    #[getter]
    fn element(&self, py: Python<'_>) -> PyResult<crate::pyxml::PyXmlElement> {
        with_core(&self.tpl, py, |core| ensure_settings_part(core)).map_err(py_err)?;
        Ok(crate::pyxml::PyXmlElement {
            tpl: self.tpl.clone_ref(py),
            part: "word/settings.xml".to_string(),
            path: Vec::new(),
        })
    }

    #[getter]
    fn odd_and_even_pages_header_footer(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| {
            core.part_dom("word/settings.xml")
                .map(|dom| dom.root.find("w:evenAndOddHeaders").is_some())
                .unwrap_or(false)
        })
    }

    #[setter]
    fn set_odd_and_even_pages_header_footer(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| {
            with_settings(core, |root| {
                let exists = root.find("w:evenAndOddHeaders").is_some();
                if v && !exists {
                    root.children
                        .push(Node::Elem(Element::new("w:evenAndOddHeaders")));
                } else if !v && exists {
                    root.children
                        .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:evenAndOddHeaders"));
                }
            })
            .map_err(py_err)
        })
    }

    /// Update fields (PAGE/NUMPAGES/TOC/...) when the document is opened in
    /// Word (w:updateFields in settings.xml).
    #[getter]
    fn update_fields_on_open(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| {
            core.part_dom("word/settings.xml")
                .map(|dom| dom.root.find("w:updateFields").is_some())
                .unwrap_or(false)
        })
    }

    #[setter]
    fn set_update_fields_on_open(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| {
            with_settings(core, |root| {
                let exists = root.find("w:updateFields").is_some();
                if v && !exists {
                    root.children
                        .push(Node::Elem(Element::new("w:updateFields")));
                } else if !v && exists {
                    root.children
                        .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:updateFields"));
                }
            })
            .map_err(py_err)
        })
    }
}

// ---------------- inline shapes ----------------

/// An inline shape (read-only snapshot with live lengths).
#[pyclass(name = "InlineShape", unsendable, skip_from_py_object)]
pub struct PyInlineShape {
    #[pyo3(get)]
    pub width: crate::pyclasses::PyLength,
    #[pyo3(get)]
    pub height: crate::pyclasses::PyLength,
    #[pyo3(get, name = "type")]
    pub kind: String,
}

// ---------------- core properties ----------------

pub const CORE_PROPS: &[(&str, &str)] = &[
    ("author", "dc:creator"),
    ("category", "cp:category"),
    ("comments", "dc:description"),
    ("content_status", "cp:contentStatus"),
    ("identifier", "dc:identifier"),
    ("keywords", "cp:keywords"),
    ("language", "dc:language"),
    ("last_modified_by", "cp:lastModifiedBy"),
    ("revision", "cp:revision"),
    ("subject", "dc:subject"),
    ("title", "dc:title"),
    ("created", "dcterms:created"),
    ("modified", "dcterms:modified"),
];

pub fn get_core_property(core: &mut TplCore, tag: &str) -> String {
    core
        .part_dom("docProps/core.xml")
        .ok()
        .and_then(|dom| dom.root.find(tag).map(|e| e.text_content()))
        .unwrap_or_default()
}

pub fn set_core_property(core: &mut TplCore, tag: &str, value: &str) -> Result<(), String> {
    core.init_docx(false)?;
    if core
        .package
        .as_ref()
        .map(|p| !p.contains("docProps/core.xml"))
        .unwrap_or(false)
    {
        crate::package::ensure_core_part(core.package.as_mut().unwrap());
    }
    {
        let dom = core.part_dom("docProps/core.xml")?;
        match dom.root.find_mut(tag) {
            Some(el) => {
                el.children = vec![Node::Text(value.to_string())];
            }
            None => {
                let mut el = Element::new(tag);
                el.children.push(Node::Text(value.to_string()));
                let pos = dom
                    .root
                    .children
                    .iter()
                    .position(|c| matches!(c, Node::Elem(e) if e.name.starts_with("dcterms:")));
                match pos {
                    Some(i) => dom.root.children.insert(i, Node::Elem(el)),
                    None => dom.root.children.push(Node::Elem(el)),
                }
            }
        }
    }
    core.mark_part_dirty("docProps/core.xml");
    Ok(())
}

/// Core properties of the document (read/write).
#[pyclass(name = "CoreProperties", unsendable)]
pub struct PyCoreProperties {
    pub tpl: Py<PyDocxTemplate>,
}

fn prop_get(tpl: &Py<PyDocxTemplate>, py: Python<'_>, attr: &str) -> String {
    let tag = CORE_PROPS
        .iter()
        .find(|(a, _)| *a == attr)
        .map(|(_, t)| *t)
        .unwrap_or(attr);
    with_core(tpl, py, |core| get_core_property(core, tag))
}

fn prop_set(tpl: &Py<PyDocxTemplate>, py: Python<'_>, attr: &str, value: &str) -> PyResult<()> {
    let tag = CORE_PROPS
        .iter()
        .find(|(a, _)| *a == attr)
        .map(|(_, t)| *t)
        .unwrap_or(attr);
    with_core(tpl, py, |core| set_core_property(core, tag, value))
        .map_err(PyRuntimeError::new_err)
}

#[pymethods]
impl PyCoreProperties {
    #[getter]
    fn author(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "author")
    }
    #[setter]
    fn set_author(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "author", &value)
    }
    #[getter]
    fn category(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "category")
    }
    #[setter]
    fn set_category(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "category", &value)
    }
    #[getter]
    fn comments(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "comments")
    }
    #[setter]
    fn set_comments(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "comments", &value)
    }
    #[getter]
    fn content_status(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "content_status")
    }
    #[setter]
    fn set_content_status(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "content_status", &value)
    }
    #[getter]
    fn identifier(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "identifier")
    }
    #[setter]
    fn set_identifier(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "identifier", &value)
    }
    #[getter]
    fn keywords(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "keywords")
    }
    #[setter]
    fn set_keywords(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "keywords", &value)
    }
    #[getter]
    fn language(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "language")
    }
    #[setter]
    fn set_language(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "language", &value)
    }
    #[getter]
    fn last_modified_by(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "last_modified_by")
    }
    #[setter]
    fn set_last_modified_by(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "last_modified_by", &value)
    }
    #[getter]
    fn revision(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "revision")
    }
    #[setter]
    fn set_revision(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "revision", &value)
    }
    #[getter]
    fn subject(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "subject")
    }
    #[setter]
    fn set_subject(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "subject", &value)
    }
    #[getter]
    fn title(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "title")
    }
    #[setter]
    fn set_title(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "title", &value)
    }
    #[getter]
    fn created(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "created")
    }
    #[setter]
    fn set_created(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "created", &value)
    }
    #[getter]
    fn modified(&self, py: Python<'_>) -> String {
        prop_get(&self.tpl, py, "modified")
    }
    #[setter]
    fn set_modified(&self, py: Python<'_>, value: String) -> PyResult<()> {
        prop_set(&self.tpl, py, "modified", &value)
    }
}
