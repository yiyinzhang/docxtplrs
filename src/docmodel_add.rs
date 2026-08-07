//! Mutation support for the Document facade: add_paragraph / add_heading /
//! add_picture / add_table / add_page_break / add_section.
//!
//! Thin forwarding wrappers: the core logic lives in [`crate::doc`].

use crate::docmodel::{with_core, PyDocument, PyParagraph, PySection, PyTable};
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyBytes;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

// ---------------- add_* implementations ----------------

pub fn doc_add_paragraph(
    doc: &PyDocument,
    py: Python<'_>,
    text: &str,
    style: Option<String>,
) -> PyResult<PyParagraph> {
    let p = with_core(&doc.tpl, py, |core| {
        crate::doc::add_paragraph(core, text, style.as_deref())
    })
    .map_err(py_err)?;
    Ok(PyParagraph {
        tpl: doc.tpl.clone_ref(py),
        index: p.index,
    })
}

pub fn doc_add_heading(
    doc: &PyDocument,
    py: Python<'_>,
    text: &str,
    level: u32,
) -> PyResult<PyParagraph> {
    let p = with_core(&doc.tpl, py, |core| crate::doc::add_heading(core, text, level))
        .map_err(py_err)?;
    Ok(PyParagraph {
        tpl: doc.tpl.clone_ref(py),
        index: p.index,
    })
}

pub fn doc_add_picture(
    doc: &PyDocument,
    py: Python<'_>,
    image_descriptor: &Bound<'_, PyAny>,
    width: Option<i64>,
    height: Option<i64>,
) -> PyResult<()> {
    let (blob, filename) = read_image_source(image_descriptor)?;
    with_core(&doc.tpl, py, |core| {
        crate::doc::add_picture(core, &blob, filename.as_deref(), width, height)
    })
    .map_err(py_err)
}

pub(crate) fn read_image_source(obj: &Bound<'_, PyAny>) -> PyResult<(Vec<u8>, Option<String>)> {
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
    let t = with_core(&doc.tpl, py, |core| crate::doc::add_table(core, rows, cols))
        .map_err(py_err)?;
    Ok(PyTable {
        tpl: doc.tpl.clone_ref(py),
        index: t.index,
    })
}

pub fn doc_add_page_break(doc: &PyDocument, py: Python<'_>) -> PyResult<()> {
    with_core(&doc.tpl, py, |core| crate::doc::add_page_break(core)).map_err(py_err)
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
    let s = with_core(&doc.tpl, py, |core| crate::doc::add_section(core, type_str))
        .map_err(py_err)?;
    Ok(PySection {
        tpl: doc.tpl.clone_ref(py),
        index: s.index,
    })
}
