//! Comments support (python-docx 1.2 comments API).
//!
//! Thin forwarding wrappers: the DOM logic lives in [`crate::doc`].

use crate::docmodel::{with_core, PyDocument};
use crate::pyclasses::PyDocxTemplate;
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
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
        crate::doc::append_comment(core, text, author, initials)
    })
    .map_err(py_err)?;

    with_core(&doc.tpl, py, |core| {
        crate::doc::anchor_comment(core, first, last, comment_id)
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

#[pymethods]
impl PyComment {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| {
            crate::doc::comment_read(core, self.comment_id, crate::doc::element_text)
        })
        .unwrap_or_default()
    }
    #[getter]
    fn author(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| {
            crate::doc::comment_read(core, self.comment_id, |c| {
                c.get_attr("w:author").unwrap_or("").to_string()
            })
        })
        .unwrap_or_default()
    }
    #[getter]
    fn initials(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| {
            crate::doc::comment_read(core, self.comment_id, |c| {
                c.get_attr("w:initials").unwrap_or("").to_string()
            })
        })
        .unwrap_or_default()
    }
    #[getter]
    fn timestamp(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| {
            crate::doc::comment_read(core, self.comment_id, |c| {
                c.get_attr("w:date").unwrap_or("").to_string()
            })
        })
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
        with_core(&self.tpl, py, |core| crate::doc::comment_ids(core))
            .into_iter()
            .map(|id| PyComment {
                tpl: self.tpl.clone_ref(py),
                comment_id: id,
            })
            .collect()
    }

    fn __len__(&self, py: Python<'_>) -> usize {
        self.comment_list(py).len()
    }

    /// Add an (unanchored) comment (python-docx comments.add_comment).
    #[pyo3(signature = (text="", author="", initials=""))]
    fn add_comment(&self, py: Python<'_>, text: &str, author: &str, initials: &str) -> PyResult<PyComment> {
        let id = with_core(&self.tpl, py, |core| {
            crate::doc::append_comment(core, text, author, initials)
        })
        .map_err(py_err)?;
        Ok(PyComment {
            tpl: self.tpl.clone_ref(py),
            comment_id: id,
        })
    }
}
