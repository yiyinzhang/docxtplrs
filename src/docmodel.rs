//! A python-docx-inspired facade over the current document:
//! live, writable proxies (Paragraph/Run/Table/Cell), sections, styles,
//! inline shapes and core properties.
//!
//! All DOM logic lives in the pure-Rust [`crate::doc`] handles; the pyclasses
//! here are thin forwarding wrappers.

use crate::doc::{self, BlockItem};
use crate::pyclasses::{PyDocxTemplate, PyLength};
use crate::template::{TplCore, DOCUMENT_PART};
use crate::xmldom::{Element, Node};
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

fn val_err(e: String) -> PyErr {
    PyValueError::new_err(e)
}

pub(crate) fn with_core<R>(tpl: &Py<PyDocxTemplate>, py: Python<'_>, f: impl FnOnce(&mut TplCore) -> R) -> R {
    f(&mut tpl.bind(py).borrow().core.borrow_mut())
}

fn to_pylen(l: doc::Length) -> PyLength {
    PyLength { emu: l.emu }
}

fn from_pylen(v: &Bound<'_, PyAny>) -> PyResult<Option<doc::Length>> {
    Ok(crate::pyclasses::extract_length_pub(v)?.map(|emu| doc::Length { emu }))
}

// ---------------- proxies ----------------

/// A paragraph in the document (live proxy).
#[pyclass(name = "Paragraph", unsendable)]
pub struct PyParagraph {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

impl PyParagraph {
    fn handle(&self) -> doc::Paragraph {
        doc::Paragraph { index: self.index }
    }
}

#[pymethods]
impl PyParagraph {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_text(core, &v)).map_err(val_err)
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().style(core))
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_style(core, &v)).map_err(val_err)
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
        let n = with_core(&self.tpl, py, |core| self.handle().run_count(core));
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
        let r = with_core(&self.tpl, py, |core| self.handle().add_run(core, text))
            .map_err(val_err)?;
        Ok(PyRun {
            tpl: self.tpl.clone_ref(py),
            para: r.para,
            index: r.index,
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
        with_core(&self.tpl, py, |core| self.handle().clear(core)).map_err(val_err)
    }

    /// Field codes in this paragraph (w:fldSimple and complex fields).
    #[getter]
    fn fields(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyField> {
        let n = with_core(&self.tpl, py, |core| self.handle().field_count(core));
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
        let f = with_core(&self.tpl, py, |core| self.handle().add_field(core, instr, cached))
            .map_err(val_err)?;
        Ok(crate::docmodel_fmt::PyField {
            tpl: self.tpl.clone_ref(py),
            para: f.para,
            index: f.index,
        })
    }

    /// Hyperlinks in this paragraph (read-only proxies).
    #[getter]
    fn hyperlinks(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyHyperlink> {
        let n = with_core(&self.tpl, py, |core| self.handle().hyperlink_count(core));
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
        with_core(&self.tpl, py, |core| self.handle().contains_page_break(core))
    }

    /// Rendered page breaks (w:lastRenderedPageBreak markers written by Word
    /// at save time) in this paragraph.
    #[getter]
    fn rendered_page_breaks(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyRenderedPageBreak> {
        let n = with_core(&self.tpl, py, |core| self.handle().rendered_page_break_count(core));
        (0..n)
            .map(|_| crate::docmodel_fmt::PyRenderedPageBreak {})
            .collect()
    }

    /// Runs and hyperlinks of this paragraph in document order.
    fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let items = with_core(&self.tpl, py, |core| self.handle().iter_inner_content(core));
        let mut out = Vec::new();
        for item in items {
            match item {
                doc::ParaItem::Run(ri) => {
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
                }
                doc::ParaItem::Hyperlink(hi) => {
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
                }
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
        let p = with_core(&self.tpl, py, |core| {
            self.handle().insert_paragraph_before(core, text, style)
        })
        .map_err(val_err)?;
        Ok(PyParagraph {
            tpl: self.tpl.clone_ref(py),
            index: p.index,
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
    fn handle(&self) -> doc::Run {
        doc::Run {
            para: self.para,
            index: self.index,
        }
    }
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
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_text(core, &v)).map_err(val_err)
    }

    #[getter]
    fn bold(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().bold(core))
    }
    #[setter]
    fn set_bold(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_bold(core, v)).map_err(val_err)
    }

    #[getter]
    fn italic(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().italic(core))
    }
    #[setter]
    fn set_italic(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_italic(core, v)).map_err(val_err)
    }

    #[getter]
    fn strike(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().strike(core))
    }
    #[setter]
    fn set_strike(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_strike(core, v)).map_err(val_err)
    }

    #[getter]
    fn underline(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().underline(core))
    }
    #[setter]
    fn set_underline(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let val = if let Ok(s) = v.extract::<String>() {
            Some(s)
        } else if v.is_truthy()? {
            Some("single".to_string())
        } else {
            None
        };
        with_core(&self.tpl, py, |core| self.handle().set_underline(core, val.as_deref()))
            .map_err(val_err)
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().style(core))
    }
    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_style(core, &v)).map_err(val_err)
    }

    #[getter]
    fn font_name(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().font_name(core))
    }
    #[setter]
    fn set_font(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_font_name(core, &v))
            .map_err(val_err)
    }

    #[getter]
    fn size(&self, py: Python<'_>) -> Option<u32> {
        with_core(&self.tpl, py, |core| self.handle().size(core))
    }
    #[setter]
    fn set_size(&self, py: Python<'_>, v: u32) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_size(core, v)).map_err(val_err)
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().color(core))
    }
    #[setter]
    fn set_color(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_color(core, &v)).map_err(val_err)
    }

    #[getter]
    fn highlight(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().highlight(core))
    }
    #[setter]
    fn set_highlight(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_highlight(core, &v))
            .map_err(val_err)
    }

    #[getter]
    fn subscript(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().subscript(core))
    }
    #[setter]
    fn set_subscript(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_subscript(core, v))
            .map_err(val_err)
    }

    #[getter]
    fn superscript(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().superscript(core))
    }
    #[setter]
    fn set_superscript(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_superscript(core, v))
            .map_err(val_err)
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
        with_core(&self.tpl, py, |core| self.handle().add_break(core, break_type))
            .map_err(val_err)
    }

    /// Append a tab character (w:tab).
    fn add_tab(&self, py: Python<'_>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().add_tab(core)).map_err(val_err)
    }

    /// Append text (w:t), preserving leading/trailing whitespace.
    fn add_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().add_text(core, text)).map_err(val_err)
    }

    /// Remove all content, keeping run properties (python-docx clear).
    fn clear(&self, py: Python<'_>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().clear(core)).map_err(val_err)
    }

    /// True when a rendered page break (w:lastRenderedPageBreak) occurs in
    /// this run (hard breaks are not counted, python-docx semantics).
    #[getter]
    fn contains_page_break(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().contains_page_break(core))
    }

    /// Content items of this run in order (python-docx iter_inner_content):
    /// contiguous text-ish ranges as strings, drawings as live XmlElement
    /// proxies, rendered page breaks as RenderedPageBreak markers.
    fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let items = with_core(&self.tpl, py, |core| self.handle().iter_inner_content(core));
        let mut out = Vec::new();
        for item in items {
            match item {
                doc::RunItem::Text(text) => {
                    out.push(pyo3::types::PyString::new(py, &text).into_any().unbind());
                }
                doc::RunItem::RenderedPageBreak => {
                    if let Ok(v) = Py::new(py, crate::docmodel_fmt::PyRenderedPageBreak {}) {
                        out.push(v.into_any());
                    }
                }
                doc::RunItem::Drawing(path) => {
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
        with_core(&self.tpl, py, |core| {
            self.handle()
                .add_picture(core, &blob, filename.as_deref(), width, height)
        })
        .map_err(py_err)
    }

    /// Mark the range from this run to `last_run` as belonging to the
    /// comment `comment_id` (python-docx run.mark_comment_range).
    fn mark_comment_range(&self, py: Python<'_>, last_run: Bound<'_, PyRun>, comment_id: i64) -> PyResult<()> {
        let (lpara, lindex) = (last_run.borrow().para, last_run.borrow().index);
        with_core(&self.tpl, py, |core| {
            self.handle().mark_comment_range(
                core,
                doc::Run {
                    para: lpara,
                    index: lindex,
                },
                comment_id,
            )
        })
        .map_err(py_err)
    }
}

/// A table in the document (live proxy).
#[pyclass(name = "Table", unsendable)]
pub struct PyTable {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

impl PyTable {
    fn handle(&self) -> doc::Table {
        doc::Table { index: self.index }
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
        let n = with_core(&self.tpl, py, |core| self.handle().row_count(core));
        (0..n)
            .map(|row| PyTableRow {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row,
            })
            .collect()
    }

    fn add_row(&self, py: Python<'_>) -> PyResult<PyTableRow> {
        let r = with_core(&self.tpl, py, |core| self.handle().add_row(core)).map_err(val_err)?;
        Ok(PyTableRow {
            tpl: self.tpl.clone_ref(py),
            index: r.index,
            row: r.row,
        })
    }

    /// Logical-grid cell access (python-docx table.cell semantics).
    fn cell(&self, py: Python<'_>, i: usize, j: usize) -> PyResult<PyCell> {
        let c = with_core(&self.tpl, py, |core| self.handle().cell(core, i, j))
            .ok_or_else(|| {
                pyo3::exceptions::PyIndexError::new_err("cell index out of range")
            })?;
        Ok(PyCell {
            tpl: self.tpl.clone_ref(py),
            index: c.index,
            row: c.row,
            col: c.col,
        })
    }

    /// The table style id (python-docx table.style; accepts a style name or
    /// id on assignment, returns the style id).
    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().style(core))
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_style(core, &v)).map_err(val_err)
    }

    /// Table alignment as a WD_TABLE_ALIGNMENT int (xml name also accepted).
    #[getter]
    fn alignment(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().alignment(core))
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
        with_core(&self.tpl, py, |core| self.handle().set_alignment(core, val))
            .map_err(val_err)
    }

    /// Autofit (tblLayout type=autofit vs fixed; missing -> True).
    #[getter]
    fn autofit(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().autofit(core))
    }
    #[setter]
    fn set_autofit(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_autofit(core, v)).map_err(val_err)
    }

    /// Append a column of the given width (gridCol + one cell per row).
    fn add_column(&self, py: Python<'_>, width: &Bound<'_, PyAny>) -> PyResult<crate::docmodel_fmt::PyTableColumn> {
        let emu = crate::pyclasses::extract_length_pub(width)?
            .ok_or_else(|| PyValueError::new_err("width is required"))?;
        let c = with_core(&self.tpl, py, |core| self.handle().add_column(core, emu))
            .map_err(val_err)?;
        Ok(crate::docmodel_fmt::PyTableColumn {
            tpl: self.tpl.clone_ref(py),
            index: c.index,
            col: c.col,
        })
    }

    #[getter]
    fn columns(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyTableColumn> {
        let n = with_core(&self.tpl, py, |core| self.handle().column_count(core));
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
        with_core(&self.tpl, py, |core| self.handle().table_direction(core))
    }
    #[setter]
    fn set_table_direction(&self, py: Python<'_>, v: i64) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_table_direction(core, v))
            .map_err(val_err)
    }

    /// Cells of column `i` (one per row, logical grid).
    fn column_cells(&self, py: Python<'_>, i: usize) -> Vec<PyCell> {
        with_core(&self.tpl, py, |core| self.handle().column_cells(core, i))
            .into_iter()
            .map(|c| PyCell {
                tpl: self.tpl.clone_ref(py),
                index: c.index,
                row: c.row,
                col: c.col,
            })
            .collect()
    }

    /// Cells of row `i` (logical grid).
    fn row_cells(&self, py: Python<'_>, i: usize) -> Vec<PyCell> {
        with_core(&self.tpl, py, |core| self.handle().row_cells(core, i))
            .into_iter()
            .map(|c| PyCell {
                tpl: self.tpl.clone_ref(py),
                index: c.index,
                row: c.row,
                col: c.col,
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
    fn handle(&self) -> doc::TableRow {
        doc::TableRow {
            index: self.index,
            row: self.row,
        }
    }
}

#[pymethods]
impl PyTableRow {
    /// Row height (w:trPr/w:trHeight w:val).
    #[getter]
    fn height(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().height(core)).map(to_pylen)
    }
    #[setter]
    fn set_height(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_height(core, l)).map_err(val_err)
    }

    /// Row height rule as a WD_ROW_HEIGHT_RULE int (0=auto, 1=atLeast,
    /// 2=exact; xml name also accepted on set).
    #[getter]
    fn height_rule(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().height_rule(core))
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
        with_core(&self.tpl, py, |core| self.handle().set_height_rule(core, val))
            .map_err(val_err)
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
        with_core(&self.tpl, py, |core| self.handle().grid_cols_before(core))
    }
    /// Grid columns after this row (trPr/gridAfter; default 0).
    #[getter]
    fn grid_cols_after(&self, py: Python<'_>) -> i64 {
        with_core(&self.tpl, py, |core| self.handle().grid_cols_after(core))
    }

    /// Cells of this row in logical-grid order (python-docx row.cells:
    /// gridSpan/vMerge covered coordinates resolve to the merged origin).
    #[getter]
    fn cells(&self, py: Python<'_>) -> Vec<PyCell> {
        with_core(&self.tpl, py, |core| self.handle().cells(core))
            .into_iter()
            .map(|c| PyCell {
                tpl: self.tpl.clone_ref(py),
                index: c.index,
                row: c.row,
                col: c.col,
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
    fn handle(&self) -> doc::Cell {
        doc::Cell {
            index: self.index,
            row: self.row,
            col: self.col,
        }
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
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_text(core, &v)).map_err(val_err)
    }

    /// Paragraphs in this cell.
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyCellParagraph> {
        let n = with_core(&self.tpl, py, |core| self.handle().paragraph_count(core));
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
        let p = with_core(&self.tpl, py, |core| self.handle().add_paragraph(core, text, style))
            .map_err(val_err)?;
        Ok(crate::docmodel_fmt::PyCellParagraph {
            tpl: self.tpl.clone_ref(py),
            index: p.index,
            row: p.row,
            col: p.col,
            para: p.para,
        })
    }

    /// Vertical alignment as a WD_CELL_VERTICAL_ALIGNMENT int (0=top,
    /// 1=center, 3=bottom, 101=both; xml name also accepted on set).
    #[getter]
    fn vertical_alignment(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().vertical_alignment(core))
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
        with_core(&self.tpl, py, |core| self.handle().set_vertical_alignment(core, val))
            .map_err(val_err)
    }

    /// Cell width (w:tcPr/w:tcW; set forces type=dxa).
    #[getter]
    fn width(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().width(core)).map(to_pylen)
    }
    #[setter]
    fn set_width(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_width(core, l)).map_err(val_err)
    }

    /// Grid columns spanned by this cell (w:gridSpan; default 1).
    #[getter]
    fn grid_span(&self, py: Python<'_>) -> i64 {
        with_core(&self.tpl, py, |core| self.handle().grid_span(core))
    }

    /// Tables nested inside this cell.
    #[getter]
    fn tables(&self, py: Python<'_>) -> Vec<crate::docmodel_fmt::PyCellTable> {
        let n = with_core(&self.tpl, py, |core| self.handle().table_count(core));
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
        let t = with_core(&self.tpl, py, |core| self.handle().add_table(core, rows, cols))
            .map_err(val_err)?;
        Ok(crate::docmodel_fmt::PyCellTable {
            tpl: self.tpl.clone_ref(py),
            index: t.index,
            row: t.row,
            col: t.col,
            tindex: t.tindex,
        })
    }

    /// Paragraphs and tables of this cell in document order.
    fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let items = with_core(&self.tpl, py, |core| self.handle().iter_inner_content(core));
        let mut out = Vec::new();
        for item in items {
            match item {
                BlockItem::Paragraph(pi) => {
                    if let Ok(v) = Py::new(py, crate::docmodel_fmt::PyCellParagraph {
                        tpl: self.tpl.clone_ref(py),
                        index: self.index,
                        row: self.row,
                        col: self.col,
                        para: pi,
                    }) {
                        out.push(v.into_any());
                    }
                }
                BlockItem::Table(ti) => {
                    if let Ok(v) = Py::new(py, crate::docmodel_fmt::PyCellTable {
                        tpl: self.tpl.clone_ref(py),
                        index: self.index,
                        row: self.row,
                        col: self.col,
                        tindex: ti,
                    }) {
                        out.push(v.into_any());
                    }
                }
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
        let merged = with_core(&self.tpl, py, |core| {
            self.handle().merge(
                core,
                doc::Cell {
                    index: oidx,
                    row: orow,
                    col: ocol,
                },
            )
        })
        .map_err(val_err)?;
        Ok(PyCell {
            tpl: self.tpl.clone_ref(py),
            index: merged.index,
            row: merged.row,
            col: merged.col,
        })
    }
}

// ---------------- sections ----------------

/// A document section (live proxy).
#[pyclass(name = "Section", unsendable)]
pub struct PySection {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
}

impl PySection {
    fn handle(&self) -> doc::Section {
        doc::Section { index: self.index }
    }
}

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
    fn page_width(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().page_width(core)).map(to_pylen)
    }
    #[setter]
    fn set_page_width(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_page_width(core, l)).map_err(val_err)
    }
    #[getter]
    fn page_height(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().page_height(core)).map(to_pylen)
    }
    #[setter]
    fn set_page_height(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_page_height(core, l)).map_err(val_err)
    }
    #[getter]
    fn left_margin(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().left_margin(core)).map(to_pylen)
    }
    #[setter]
    fn set_left_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_left_margin(core, l)).map_err(val_err)
    }
    #[getter]
    fn right_margin(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().right_margin(core)).map(to_pylen)
    }
    #[setter]
    fn set_right_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_right_margin(core, l)).map_err(val_err)
    }
    #[getter]
    fn top_margin(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().top_margin(core)).map(to_pylen)
    }
    #[setter]
    fn set_top_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_top_margin(core, l)).map_err(val_err)
    }
    #[getter]
    fn bottom_margin(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().bottom_margin(core)).map(to_pylen)
    }
    #[setter]
    fn set_bottom_margin(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_bottom_margin(core, l)).map_err(val_err)
    }
    #[getter]
    fn header_distance(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().header_distance(core)).map(to_pylen)
    }
    #[setter]
    fn set_header_distance(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_header_distance(core, l)).map_err(val_err)
    }
    #[getter]
    fn footer_distance(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().footer_distance(core)).map(to_pylen)
    }
    #[setter]
    fn set_footer_distance(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_footer_distance(core, l)).map_err(val_err)
    }
    #[getter]
    fn gutter(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().gutter(core)).map(to_pylen)
    }
    #[setter]
    fn set_gutter(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let l = from_pylen(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_gutter(core, l)).map_err(val_err)
    }

    #[getter]
    fn orientation(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().orientation(core))
    }

    #[setter]
    fn set_orientation(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_orientation(core, &v))
            .map_err(val_err)
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
        with_core(&self.tpl, py, |core| {
            self.handle().different_first_page_header_footer(core)
        })
    }

    #[setter]
    fn set_different_first_page_header_footer(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| {
            self.handle().set_different_first_page_header_footer(core, v)
        })
        .map_err(val_err)
    }

    /// Section start type as a WD_SECTION_START int (0=continuous,
    /// 1=nextColumn, 2=nextPage, 3=evenPage, 4=oddPage; xml name also
    /// accepted on set). Missing w:type reads as 2 (next page).
    #[getter]
    fn start_type(&self, py: Python<'_>) -> i64 {
        with_core(&self.tpl, py, |core| self.handle().start_type(core))
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
        with_core(&self.tpl, py, |core| self.handle().set_start_type(core, val))
            .map_err(val_err)
    }

    /// Paragraphs and tables of this section in document order. Section
    /// boundaries are the paragraphs carrying a paragraph-level sectPr; the
    /// last section ends at the body-level sectPr.
    pub fn iter_inner_content(&self, py: Python<'_>) -> Vec<Py<PyAny>> {
        let items = with_core(&self.tpl, py, |core| self.handle().iter_inner_content(core));
        let mut out = Vec::new();
        for item in items {
            match item {
                BlockItem::Paragraph(idx) => {
                    if let Ok(v) = Py::new(py, PyParagraph { tpl: self.tpl.clone_ref(py), index: idx }) {
                        out.push(v.into_any());
                    }
                }
                BlockItem::Table(idx) => {
                    if let Ok(v) = Py::new(py, PyTable { tpl: self.tpl.clone_ref(py), index: idx }) {
                        out.push(v.into_any());
                    }
                }
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
    pub kind: String, // "header" | "footer" | even_/first_ variants
}

impl PySectionHdrFtr {
    fn handle(&self) -> doc::HdrFtr {
        doc::HdrFtr {
            section: self.section,
            kind: self.kind.clone(),
        }
    }
}

#[pymethods]
impl PySectionHdrFtr {
    #[getter]
    fn is_linked_to_previous(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().is_linked_to_previous(core))
    }

    #[setter]
    fn set_is_linked_to_previous(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_is_linked_to_previous(core, v))
            .map_err(py_err)
    }

    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Vec<String> {
        with_core(&self.tpl, py, |core| self.handle().paragraphs(core))
    }

    #[pyo3(signature = (text=""))]
    fn add_paragraph(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().add_paragraph(core, text))
            .map_err(py_err)
    }
}

// ---------------- styles ----------------

/// A style in the document (live proxy).
#[pyclass(name = "Style", unsendable)]
pub struct PyStyle {
    pub tpl: Py<PyDocxTemplate>,
    pub style_id: String,
}

impl PyStyle {
    fn handle(&self) -> doc::Style {
        doc::Style {
            style_id: self.style_id.clone(),
        }
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
        with_core(&self.tpl, py, |core| self.handle().name(core))
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_name(core, &v)).map_err(val_err)
    }

    #[getter]
    fn style_id(&self) -> String {
        self.style_id.clone()
    }

    #[getter]
    fn style_type(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().style_type(core))
    }

    #[getter]
    fn base_style(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().base_style(core))
    }

    #[setter]
    fn set_base_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_base_style(core, &v))
            .map_err(val_err)
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
        with_core(&self.tpl, py, |core| self.handle().hidden(core))
    }
    #[setter]
    fn set_hidden(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_hidden(core, v)).map_err(val_err)
    }

    /// Locked against editing (w:locked).
    #[getter]
    fn locked(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().locked(core))
    }
    #[setter]
    fn set_locked(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_locked(core, v)).map_err(val_err)
    }

    /// Shown in the quick style gallery (w:qFormat).
    #[getter]
    fn quick_style(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().quick_style(core))
    }
    #[setter]
    fn set_quick_style(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_quick_style(core, v))
            .map_err(val_err)
    }

    /// Re-hide when the style is no longer used (w:unhideWhenUsed).
    #[getter]
    fn unhide_when_used(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().unhide_when_used(core))
    }
    #[setter]
    fn set_unhide_when_used(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_unhide_when_used(core, v))
            .map_err(val_err)
    }

    /// UI priority (w:uiPriority w:val); None removes it.
    #[getter]
    fn priority(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().priority(core))
    }
    #[setter]
    fn set_priority(&self, py: Python<'_>, v: Option<i64>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_priority(core, v)).map_err(val_err)
    }

    /// Builtin styles lack the w:customStyle attribute (read-only).
    #[getter]
    fn builtin(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().builtin(core))
    }

    /// Style applied to the next paragraph (w:next; paragraph styles).
    #[getter]
    fn next_paragraph_style(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().next_paragraph_style(core))
    }
    #[setter]
    fn set_next_paragraph_style(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let name: Option<String> = if v.is_none() { None } else { Some(v.extract()?) };
        with_core(&self.tpl, py, |core| {
            self.handle().set_next_paragraph_style(core, name.as_deref())
        })
        .map_err(val_err)
    }

    fn delete(&self, py: Python<'_>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().delete(core)).map_err(py_err)
    }
}

/// Font properties of a style (legacy two-state facade).
#[pyclass(name = "StyleFont", unsendable)]
pub struct PyStyleFont {
    pub tpl: Py<PyDocxTemplate>,
    pub style_id: String,
}

impl PyStyleFont {
    fn edit_rpr<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        let sid = self.style_id.clone();
        with_core(&self.tpl, py, |core| {
            doc::style_edit(core, &sid, |el| {
                if el.find("w:rPr").is_none() {
                    el.children.push(Node::Elem(Element::new("w:rPr")));
                }
                f(el.find_mut("w:rPr").unwrap())
            })
        })
        .map_err(val_err)
    }

    fn read_rpr<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let sid = self.style_id.clone();
        with_core(&self.tpl, py, |core| {
            doc::style_read(core, &sid, |el| el.find("w:rPr").map(|r| f(r))).flatten()
        })
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
        with_core(&self.tpl, py, |core| doc::ensure_styles_part(core)).map_err(py_err)?;
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
        with_core(&self.tpl, py, |core| doc::style_ids(core))
            .into_iter()
            .map(|id| PyStyle {
                tpl: self.tpl.clone_ref(py),
                style_id: id,
            })
            .collect()
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
            let found = doc::style_read(core, &sid, |_| ()).is_some();
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
        let style_id = with_core(&self.tpl, py, |core| doc::add_style(core, name, type_str, builtin))
            .map_err(py_err)?;
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

#[pymethods]
impl PySettings {
    /// Raw XML root element of word/settings.xml (live proxy), the
    /// python-docx `settings.element` escape hatch.
    #[getter]
    fn element(&self, py: Python<'_>) -> PyResult<crate::pyxml::PyXmlElement> {
        with_core(&self.tpl, py, |core| doc::ensure_settings_part(core)).map_err(py_err)?;
        Ok(crate::pyxml::PyXmlElement {
            tpl: self.tpl.clone_ref(py),
            part: "word/settings.xml".to_string(),
            path: Vec::new(),
        })
    }

    #[getter]
    fn odd_and_even_pages_header_footer(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| doc::settings_flag(core, "w:evenAndOddHeaders"))
    }

    #[setter]
    fn set_odd_and_even_pages_header_footer(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| {
            doc::set_settings_flag(core, "w:evenAndOddHeaders", v)
        })
        .map_err(py_err)
    }

    /// Update fields (PAGE/NUMPAGES/TOC/...) when the document is opened in
    /// Word (w:updateFields in settings.xml).
    #[getter]
    fn update_fields_on_open(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| doc::settings_flag(core, "w:updateFields"))
    }

    #[setter]
    fn set_update_fields_on_open(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        with_core(&self.tpl, py, |core| doc::set_settings_flag(core, "w:updateFields", v))
            .map_err(py_err)
    }
}

// ---------------- inline shapes ----------------

/// An inline shape (read-only snapshot with live lengths).
#[pyclass(name = "InlineShape", unsendable, skip_from_py_object)]
pub struct PyInlineShape {
    #[pyo3(get)]
    pub width: PyLength,
    #[pyo3(get)]
    pub height: PyLength,
    #[pyo3(get, name = "type")]
    pub kind: String,
}

// ---------------- core properties ----------------

/// Core properties of the document (read/write).
#[pyclass(name = "CoreProperties", unsendable)]
pub struct PyCoreProperties {
    pub tpl: Py<PyDocxTemplate>,
}

fn prop_get(tpl: &Py<PyDocxTemplate>, py: Python<'_>, attr: &str) -> String {
    let tag = doc::CORE_PROPS
        .iter()
        .find(|(a, _)| *a == attr)
        .map(|(_, t)| *t)
        .unwrap_or(attr);
    with_core(tpl, py, |core| doc::get_core_property(core, tag))
}

fn prop_set(tpl: &Py<PyDocxTemplate>, py: Python<'_>, attr: &str, value: &str) -> PyResult<()> {
    let tag = doc::CORE_PROPS
        .iter()
        .find(|(a, _)| *a == attr)
        .map(|(_, t)| *t)
        .unwrap_or(attr);
    with_core(tpl, py, |core| doc::set_core_property(core, tag, value))
        .map_err(PyRuntimeError::new_err)
}

macro_rules! core_prop {
    ($($name:ident / $setter:ident => $attr:literal),* $(,)?) => {
        #[pymethods]
        impl PyCoreProperties {
            $(
            #[getter]
            fn $name(&self, py: Python<'_>) -> String {
                prop_get(&self.tpl, py, $attr)
            }
            #[setter]
            fn $setter(&self, py: Python<'_>, value: String) -> PyResult<()> {
                prop_set(&self.tpl, py, $attr, &value)
            }
            )*
        }
    };
}

core_prop! {
    author / set_author => "author",
    category / set_category => "category",
    comments / set_comments => "comments",
    content_status / set_content_status => "content_status",
    identifier / set_identifier => "identifier",
    keywords / set_keywords => "keywords",
    language / set_language => "language",
    last_modified_by / set_last_modified_by => "last_modified_by",
    revision / set_revision => "revision",
    subject / set_subject => "subject",
    title / set_title => "title",
    created / set_created => "created",
    modified / set_modified => "modified",
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
        let n = with_core(&self.tpl, py, |core| doc::count_in_body(core, "w:p"));
        (0..n)
            .map(|i| PyParagraph {
                tpl: self.tpl.clone_ref(py),
                index: i,
            })
            .collect()
    }

    #[getter]
    pub fn tables(&self, py: Python<'_>) -> Vec<PyTable> {
        let n = with_core(&self.tpl, py, |core| doc::count_in_body(core, "w:tbl"));
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
            doc::read_body(core, |body| {
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
        let n = with_core(&self.tpl, py, |core| doc::section_count(core));
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
        with_core(&self.tpl, py, |core| doc::inline_shapes(core))
            .into_iter()
            .map(|(cx, cy, kind)| PyInlineShape {
                width: PyLength { emu: cx },
                height: PyLength { emu: cy },
                kind,
            })
            .collect()
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
