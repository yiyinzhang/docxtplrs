//! Mutation support for the Document facade: add_paragraph / add_heading /
//! add_picture / add_table / add_page_break / add_section.

use crate::docmodel::{with_core, PyDocument, PyParagraph, PySection, PyTable};
use crate::richtext::{richtext_run, TextProps};
use crate::template::{TplCore, DOCUMENT_PART};
use crate::xmldom::{Document, Element, Node};
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyBytes;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

// ---------------- core helpers ----------------

/// Append a body-level fragment before the trailing sectPr (or </w:body>).
pub fn append_to_body(core: &mut TplCore, fragment: &str) -> Result<(), String> {
    // parse the fragment through a wrapper root; xmldom is name-agnostic
    let wrap = Document::parse(&format!("<w:__wrap>{}</w:__wrap>", fragment))?;
    let mut nodes = wrap.root.children;
    mutate_document(core, |body| {
        let pos = body
            .children
            .iter()
            .rposition(|c| matches!(c, Node::Elem(e) if e.name == "w:sectPr"))
            .unwrap_or(body.children.len());
        for (i, child) in nodes.drain(..).enumerate() {
            body.children.insert(pos + i, child);
        }
    })
}

/// Mutate the body element of the cached document DOM in place; the change
/// is serialized back into the package on the next flush (render/save/etc).
pub fn mutate_document(
    core: &mut TplCore,
    f: impl FnOnce(&mut Element),
) -> Result<(), String> {
    {
        let dom = core.document_dom()?;
        let body = dom
            .root
            .find_mut("w:body")
            .ok_or_else(|| "no w:body".to_string())?;
        f(body);
    }
    core.mark_doc_dirty();
    Ok(())
}

fn count_direct(core: &mut TplCore, name: &str) -> usize {
    crate::docmodel::read_body(core, |b| {
        b.children
            .iter()
            .filter(|c| matches!(c, Node::Elem(e) if e.name == name))
            .count()
    })
    .unwrap_or(0)
}

pub fn nth_direct<'a>(el: &'a mut Element, name: &str, n: usize) -> Option<&'a mut Element> {
    el.children
        .iter_mut()
        .filter_map(|c| match c {
            Node::Elem(e) if e.name == name => Some(e),
            _ => None,
        })
        .nth(n)
}

// ---------------- add_* implementations ----------------

pub fn doc_add_paragraph(
    doc: &PyDocument,
    py: Python<'_>,
    text: &str,
    style: Option<String>,
) -> PyResult<PyParagraph> {
    let index = with_core(&doc.tpl, py, |core| {
        let sid = style.map(|s| crate::subdocbuilder::resolve_style_id(core, &s));
        let mut p = String::from("<w:p>");
        if let Some(sid) = &sid {
            p.push_str(&format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", sid));
        }
        if !text.is_empty() {
            p.push_str(&richtext_run(text, &TextProps::default()));
        }
        p.push_str("</w:p>");
        append_to_body(core, &p).map(|_| count_direct(core, "w:p"))
    });
    let index = index.map_err(py_err)?;
    Ok(PyParagraph {
        tpl: doc.tpl.clone_ref(py),
        index: index - 1,
    })
}

pub fn doc_add_heading(
    doc: &PyDocument,
    py: Python<'_>,
    text: &str,
    level: u32,
) -> PyResult<PyParagraph> {
    doc_add_paragraph(doc, py, text, Some(format!("Heading {}", level.max(1))))
}

pub fn doc_add_picture(
    doc: &PyDocument,
    py: Python<'_>,
    image_descriptor: &Bound<'_, PyAny>,
    width: Option<i64>,
    height: Option<i64>,
) -> PyResult<()> {
    let (blob, filename) = read_image_source(image_descriptor)?;
    with_core(&doc.tpl, py, |core| -> Result<(), String> {
        core.init_docx(false)?;
        let drawing = crate::inline_image::drawing_xml(
            core,
            DOCUMENT_PART,
            &blob,
            filename.as_deref(),
            width,
            height,
            None,
            None,
            None,
        )?;
        append_to_body(core, &format!("<w:p><w:r>{}</w:r></w:p>", drawing))
    })
    .map_err(py_err)
}

fn read_image_source(obj: &Bound<'_, PyAny>) -> PyResult<(Vec<u8>, Option<String>)> {
    let filename = obj.extract::<String>().ok().and_then(|p| {
        std::path::Path::new(&p)
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
    });
    if let Ok(s) = obj.extract::<String>() {
        let data = std::fs::read(&s)
            .map_err(|e| PyValueError::new_err(format!("cannot read {}: {}", s, e)))?;
        return Ok((data, filename));
    }
    if let Ok(b) = obj.cast::<PyBytes>() {
        return Ok((b.as_bytes().to_vec(), None));
    }
    if let Ok(fspath) = obj.call_method0("__fspath__") {
        if let Ok(s) = fspath.extract::<String>() {
            let data = std::fs::read(&s)
                .map_err(|e| PyValueError::new_err(format!("cannot read {}: {}", s, e)))?;
            return Ok((data, filename));
        }
    }
    if let Ok(data) = obj.call_method0("read") {
        if let Ok(b) = data.cast::<PyBytes>() {
            return Ok((b.as_bytes().to_vec(), None));
        }
    }
    Err(PyValueError::new_err(
        "expected a path, bytes, or file-like object",
    ))
}

pub fn doc_add_table(
    doc: &PyDocument,
    py: Python<'_>,
    rows: usize,
    cols: usize,
) -> PyResult<PyTable> {
    let index = with_core(&doc.tpl, py, |core| {
        let usable = crate::subdocbuilder::master_usable_width_twips(core);
        let xml = crate::subdocbuilder::table_xml(
            &vec![vec![String::new(); cols]; rows],
            usable,
        );
        append_to_body(core, &xml).map(|_| count_direct(core, "w:tbl"))
    });
    let index = index.map_err(py_err)?;
    Ok(PyTable {
        tpl: doc.tpl.clone_ref(py),
        index: index - 1,
    })
}

pub fn doc_add_page_break(doc: &PyDocument, py: Python<'_>) -> PyResult<()> {
    with_core(&doc.tpl, py, |core| {
        append_to_body(core, "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>")
    })
    .map_err(py_err)
}

/// python-docx add_section: close the current section with a paragraph-level
/// sectPr copy and set the body sectPr's start type.
pub fn doc_add_section(doc: &PyDocument, py: Python<'_>, start_type: u32) -> PyResult<PySection> {
    let type_str = match start_type {
        0 => "continuous",
        1 => "nextColumn",
        2 => "nextPage",
        3 => "evenPage",
        4 => "oddPage",
        _ => return Err(PyValueError::new_err("invalid section start type")),
    };
    let index = with_core(&doc.tpl, py, |core| -> Result<usize, String> {
        mutate_document(core, |body| {
            // clone the body-level sectPr (last direct w:sectPr child)
            let body_sectpr = body
                .children
                .iter()
                .rev()
                .find_map(|c| match c {
                    Node::Elem(e) if e.name == "w:sectPr" => Some(e.clone()),
                    _ => None,
                })
                .unwrap_or_else(|| Element::new("w:sectPr"));
            // paragraph carrying the OLD section properties
            let mut p = Element::new("w:p");
            let mut ppr = Element::new("w:pPr");
            ppr.children.push(Node::Elem(body_sectpr));
            p.children.push(Node::Elem(ppr));
            // remove any body-level sectPr from it stays as the new section's
            // properties; the paragraph copy preserves the previous section
            // find the current body-level sectPr and set its type
            let mut sects: Vec<&mut Element> = Vec::new();
            crate::docmodel::collect_sectprs_mut(body, &mut sects);
            if let Some(last) = sects.last_mut() {
                if let Some(t) = last.find_mut("w:type") {
                    t.set_attr("w:val", type_str);
                } else {
                    let mut t = Element::new("w:type");
                    t.set_attr("w:val", type_str);
                    last.children.insert(0, Node::Elem(t));
                }
            }
            // insert the closing paragraph before the body sectPr
            let insert_pos = body
                .children
                .iter()
                .rposition(|c| matches!(c, Node::Elem(e) if e.name == "w:sectPr"))
                .unwrap_or(body.children.len());
            body.children.insert(insert_pos, Node::Elem(p));
        })?;
        // index of the new section = previous count
        let n = {
            let dom = core.document_dom()?;
            let mut v: Vec<&crate::xmldom::Element> = Vec::new();
            dom.root.iter_descendants("w:sectPr", &mut v);
            v.len()
        };
        Ok(n - 1)
    });
    let index = index.map_err(py_err)?;
    Ok(PySection {
        tpl: doc.tpl.clone_ref(py),
        index,
    })
}
