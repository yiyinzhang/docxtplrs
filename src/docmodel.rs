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
                    out.push_str(&e.text_content());
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
            mutate_document(core, |body| {
                if let Some(p) = nth_direct(body, "w:p", self.index) {
                    result = Some(f(p));
                }
            })
            .map_err(py_err)?;
            result.ok_or_else(|| PyValueError::new_err("paragraph not found"))
        })
    }

    pub(crate) fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                body.children
                    .iter()
                    .filter_map(|c| match c {
                        Node::Elem(e) if e.name == "w:p" => Some(e),
                        _ => None,
                    })
                    .nth(self.index)
                    .map(|p| f(p))
            })
            .flatten()
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
            mutate_document(core, |body| {
                if let Some(p) = nth_direct(body, "w:p", self.para) {
                    if let Some(r) = nth_direct(p, "w:r", self.index) {
                        result = Some(f(r));
                    }
                }
            })
            .map_err(py_err)?;
            result.ok_or_else(|| PyValueError::new_err("run not found"))
        })
    }

    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:p", self.para)
                    .and_then(|p| nth_direct_ref(p, "w:r", self.index))
                    .map(|r| f(r))
            })
            .flatten()
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

fn read_flag(el: &Element, tag: &str) -> Option<bool> {
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

    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        self.read(py, |r| element_text(r)).unwrap_or_default()
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| {
            r.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:t"));
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v));
            r.children.push(Node::Elem(t));
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
}

/// A table in the document (live proxy).
#[pyclass(name = "Table", unsendable)]
pub struct PyTable {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

impl PyTable {
    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| nth_direct_ref(body, "w:tbl", self.index).map(|t| f(t)))
                .flatten()
        })
    }

    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            mutate_document(core, |body| {
                if let Some(t) = nth_direct(body, "w:tbl", self.index) {
                    result = Some(f(t));
                }
            })
            .map_err(py_err)?;
            result.ok_or_else(|| PyValueError::new_err("table not found"))
        })
    }
}

#[pymethods]
impl PyTable {
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
}

/// A table row (live proxy).
#[pyclass(name = "TableRow", unsendable)]
pub struct PyTableRow {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
}

#[pymethods]
impl PyTableRow {
    #[getter]
    fn cells(&self, py: Python<'_>) -> Vec<PyCell> {
        let n = with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .map(|r| {
                        r.children
                            .iter()
                            .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                            .count()
                    })
            })
            .flatten()
            .unwrap_or(0)
        });
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
    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            let mut result = None;
            mutate_document(core, |body| {
                if let Some(t) = nth_direct(body, "w:tbl", self.index) {
                    if let Some(r) = nth_direct(t, "w:tr", self.row) {
                        if let Some(c) = nth_direct(r, "w:tc", self.col) {
                            result = Some(f(c));
                        }
                    }
                }
            })
            .map_err(py_err)?;
            result.ok_or_else(|| PyValueError::new_err("cell not found"))
        })
    }
}

#[pymethods]
impl PyCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| {
            read_body(core, |body| {
                nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| nth_direct_ref(r, "w:tc", self.col))
                    .map(|c| element_text(c))
            })
            .flatten()
            .unwrap_or_default()
        })
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

fn tcpr_mut(tc: &mut Element) -> &mut Element {
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

fn get_twips(sp: &Element, tag: &str, attr: &str) -> Option<i64> {
    sp.find(tag)
        .and_then(|e| e.get_attr(attr))
        .and_then(|v| v.parse::<i64>().ok())
}

fn set_twips(sp: &mut Element, tag: &str, attr: &str, v: Option<i64>, defaults: &[(&str, &str)]) {
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

fn to_len(v: Option<i64>) -> Option<crate::pyclasses::PyLength> {
    v.map(|t| crate::pyclasses::PyLength { emu: t * 635 })
}

fn from_len(obj: &Bound<'_, PyAny>) -> PyResult<Option<i64>> {
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
            pkg.set(&part, crate::package::encode_part(&xml, &enc));
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

fn find_style_el_mut<'a>(root: &'a mut Element, style_id: &str) -> Option<&'a mut Element> {
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
    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
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

    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
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

    #[getter]
    fn font(&self, py: Python<'_>) -> PyStyleFont {
        PyStyleFont {
            tpl: self.tpl.clone_ref(py),
            style_id: self.style_id.clone(),
        }
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

const DEFAULT_CORE_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type=\"dcterms:W3CDTF\">2000-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2000-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>";

/// Create the core properties part if missing (python-docx always has one).
pub fn ensure_core_part(pkg: &mut crate::package::Package) {
    if pkg.contains("docProps/core.xml") {
        return;
    }
    pkg.set("docProps/core.xml", DEFAULT_CORE_XML.as_bytes().to_vec());
    pkg.ensure_content_type_override(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
    );
    let rels_path = "_rels/.rels";
    let mut rels = pkg
        .get_string(rels_path)
        .map(|x| crate::package::Rels::from_xml(&x))
        .unwrap_or_default();
    rels.add(
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "docProps/core.xml",
        false,
    );
    pkg.set(rels_path, rels.to_xml().into_bytes());
}

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
        ensure_core_part(core.package.as_mut().unwrap());
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
