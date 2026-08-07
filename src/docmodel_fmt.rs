//! python-docx parity formatting proxies: Font, ColorFormat, ParagraphFormat,
//! TabStops/TabStop, table Column, and cell paragraphs.
//!
//! Enum values follow python-docx: getters return ints (python-docx enums are
//! int subclasses, so `== WD_*` comparisons hold); setters accept ints or the
//! xml vocabulary strings. Tri-state booleans follow python-docx CT_OnOff
//! semantics: get -> None when the element is missing; set None removes the
//! element, True writes a bare element, False writes `w:val="0"`.

use crate::doc::{
    int_of, xml_of, Font, LineSpacing, ParagraphFormat, TabStop, TabStops,
    ALIGN, HIGHLIGHT, TAB_ALIGN, TAB_LEADER, UNDERLINE,
};
// re-exported so the pyclass field types keep their docmodel_fmt::* paths
pub use crate::doc::{FontTarget, PfTarget};
use crate::docmodel::with_core;
use crate::pyclasses::{PyDocxTemplate, PyLength};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyFloat;

// ---------------------------------------------------------------- helpers

/// Accept int or xml-vocabulary string; None passthrough.
fn extract_enum(v: &Bound<'_, PyAny>, table: &'static [(i64, &'static str)]) -> PyResult<Option<i64>> {
    if v.is_none() {
        return Ok(None);
    }
    if let Ok(i) = v.extract::<i64>() {
        if xml_of(table, i).is_some() {
            return Ok(Some(i));
        }
        return Err(PyValueError::new_err(format!("invalid enum value {}", i)));
    }
    if let Ok(s) = v.extract::<String>() {
        if let Some(i) = int_of(table, &s) {
            return Ok(Some(i));
        }
        return Err(PyValueError::new_err(format!("invalid enum value {:?}", s)));
    }
    Err(PyValueError::new_err("expected int, str, or None"))
}

fn extract_emu(v: &Bound<'_, PyAny>) -> PyResult<Option<i64>> {
    crate::pyclasses::extract_length_pub(v)
}

// ---------------------------------------------------------------- Font

/// python-docx Font: full run/style character properties (tri-state booleans).
/// Thin wrapper over [`crate::doc::Font`].
#[pyclass(name = "Font", unsendable)]
pub struct PyFont {
    pub tpl: Py<PyDocxTemplate>,
    pub target: FontTarget,
}

impl PyFont {
    fn handle(&self) -> Font {
        Font {
            target: self.target.clone(),
        }
    }
}

/// Generate tri-state bool property pairs on PyFont (plus extra methods).
macro_rules! font_tri {
    ($($name:ident / $setter:ident),* $(,)?; $($extra:item)*) => {
        #[pymethods]
        impl PyFont {
            $(
            #[getter]
            fn $name(&self, py: Python<'_>) -> Option<bool> {
                with_core(&self.tpl, py, |core| self.handle().$name(core))
            }
            #[setter]
            fn $setter(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
                with_core(&self.tpl, py, |core| self.handle().$setter(core, v))
                    .map_err(PyValueError::new_err)
            }
            )*
            $($extra)*
        }
    };
}

font_tri! {
    bold / set_bold,
    cs_bold / set_cs_bold,
    italic / set_italic,
    cs_italic / set_cs_italic,
    all_caps / set_all_caps,
    small_caps / set_small_caps,
    strike / set_strike,
    double_strike / set_double_strike,
    outline / set_outline,
    shadow / set_shadow,
    emboss / set_emboss,
    imprint / set_imprint,
    no_proof / set_no_proof,
    snap_to_grid / set_snap_to_grid,
    hidden / set_hidden,
    web_hidden / set_web_hidden,
    spec_vanish / set_spec_vanish,
    rtl / set_rtl,
    complex_script / set_complex_script,
    math / set_math,
;
    #[getter]
    fn subscript(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().subscript(core))
    }
    #[setter]
    fn set_subscript(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_subscript(core, v))
            .map_err(PyValueError::new_err)
    }
    #[getter]
    fn superscript(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().superscript(core))
    }
    #[setter]
    fn set_superscript(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_superscript(core, v))
            .map_err(PyValueError::new_err)
    }

    /// Font size as a Length (EMU); w:sz stores half-points.
    #[getter]
    fn size(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().size(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_size(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_size(core, emu))
            .map_err(PyValueError::new_err)
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().name(core))
    }
    #[setter]
    fn set_name(&self, py: Python<'_>, v: Option<String>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_name(core, v))
            .map_err(PyValueError::new_err)
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> PyColorFormat {
        PyColorFormat {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
        }
    }

    /// Highlight color as a WD_COLOR_INDEX int (or xml name on set).
    #[getter]
    fn highlight_color(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().highlight_color(core))
    }
    #[setter]
    fn set_highlight_color(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = extract_enum(v, HIGHLIGHT)?;
        with_core(&self.tpl, py, |core| self.handle().set_highlight_color(core, v))
            .map_err(PyValueError::new_err)
    }

    /// Underline as a WD_UNDERLINE int; True -> single, False -> none.
    #[getter]
    fn underline(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().underline(core))
    }
    #[setter]
    fn set_underline(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = if let Ok(b) = v.extract::<bool>() {
            Some(if b { 1 } else { 0 })
        } else {
            extract_enum(v, UNDERLINE)?
        };
        with_core(&self.tpl, py, |core| self.handle().set_underline(core, v))
            .map_err(PyValueError::new_err)
    }
}

// ---------------------------------------------------------------- ColorFormat

/// python-docx ColorFormat (rgb only).
#[pyclass(name = "ColorFormat", unsendable)]
pub struct PyColorFormat {
    pub tpl: Py<PyDocxTemplate>,
    pub target: FontTarget,
}

#[pymethods]
impl PyColorFormat {
    /// RGB as a 6-digit hex string (e.g. "FF0000"), None when not set.
    #[getter]
    fn rgb(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| {
            Font {
                target: self.target.clone(),
            }
            .color_rgb(core)
        })
    }
    /// Accepts "FF0000" / "#FF0000" / RGBColor-like (str() = hex); None clears.
    #[setter]
    fn set_rgb(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let hex: Option<String> = if v.is_none() {
            None
        } else {
            let s: String = v
                .call_method0("__str__")
                .and_then(|o| o.extract())
                .map_err(|_| PyValueError::new_err("rgb must be a 6-digit hex string"))?;
            let s = s.strip_prefix('#').unwrap_or(&s).to_uppercase();
            if s.len() != 6 || !s.chars().all(|c| c.is_ascii_hexdigit()) {
                return Err(PyValueError::new_err(format!("invalid rgb value {:?}", s)));
            }
            Some(s)
        };
        with_core(&self.tpl, py, |core| {
            Font {
                target: self.target.clone(),
            }
            .set_color_rgb(core, hex)
        })
        .map_err(PyValueError::new_err)
    }
}

// ---------------------------------------------------------------- ParagraphFormat

/// python-docx ParagraphFormat (w:pPr properties).
/// Thin wrapper over [`crate::doc::ParagraphFormat`].
#[pyclass(name = "ParagraphFormat", unsendable)]
pub struct PyParagraphFormat {
    pub tpl: Py<PyDocxTemplate>,
    pub target: PfTarget,
}

impl PyParagraphFormat {
    fn handle(&self) -> ParagraphFormat {
        ParagraphFormat {
            target: self.target.clone(),
        }
    }
}

/// tri-state bool properties on ParagraphFormat (plus extra methods).
macro_rules! pf_tri {
    ($($name:ident / $setter:ident),* $(,)?; $($extra:item)*) => {
        #[pymethods]
        impl PyParagraphFormat {
            $(
            #[getter]
            fn $name(&self, py: Python<'_>) -> Option<bool> {
                with_core(&self.tpl, py, |core| self.handle().$name(core))
            }
            #[setter]
            fn $setter(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
                with_core(&self.tpl, py, |core| self.handle().$setter(core, v))
                    .map_err(PyValueError::new_err)
            }
            )*
            $($extra)*
        }
    };
}

pf_tri! {
    keep_together / set_keep_together,
    keep_with_next / set_keep_with_next,
    page_break_before / set_page_break_before,
;
    /// Alignment as a WD_ALIGN_PARAGRAPH int (xml name also accepted on set).
    #[getter]
    pub(crate) fn alignment(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().alignment(core))
    }
    #[setter]
    pub(crate) fn set_alignment(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = extract_enum(v, ALIGN)?;
        with_core(&self.tpl, py, |core| self.handle().set_alignment(core, v))
            .map_err(PyValueError::new_err)
    }

    #[getter]
    fn left_indent(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().left_indent(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_left_indent(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_left_indent(core, emu))
            .map_err(PyValueError::new_err)
    }
    #[getter]
    fn right_indent(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().right_indent(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_right_indent(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_right_indent(core, emu))
            .map_err(PyValueError::new_err)
    }

    /// First-line indent; negative values become a hanging indent.
    #[getter]
    fn first_line_indent(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().first_line_indent(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_first_line_indent(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_first_line_indent(core, emu))
            .map_err(PyValueError::new_err)
    }

    #[getter]
    fn space_before(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().space_before(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_space_before(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_space_before(core, emu))
            .map_err(PyValueError::new_err)
    }
    #[getter]
    fn space_after(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().space_after(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_space_after(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_space_after(core, emu))
            .map_err(PyValueError::new_err)
    }

    /// Line spacing: a float multiple (2.0 = double) when the rule is auto,
    /// otherwise a Length (exact / at-least).
    #[getter]
    fn line_spacing(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        let ls = with_core(&self.tpl, py, |core| self.handle().line_spacing(core))?;
        let v: Py<PyAny> = match ls {
            LineSpacing::Multiple(f) => PyFloat::new(py, f).unbind().into_any(),
            LineSpacing::Exact(l) => Py::new(py, PyLength { emu: l.emu }).ok()?.into_any(),
        };
        Some(v)
    }
    #[setter]
    fn set_line_spacing(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        // float -> multiple (lineRule=auto); Length/int -> exact twips
        let ls = if v.is_none() {
            None
        } else if let Ok(f) = v.extract::<f64>() {
            Some(LineSpacing::Multiple(f))
        } else {
            extract_emu(v)?.map(|emu| LineSpacing::Exact(crate::doc::Length { emu }))
        };
        with_core(&self.tpl, py, |core| self.handle().set_line_spacing(core, ls))
            .map_err(PyValueError::new_err)
    }

    /// Line spacing rule as a WD_LINE_SPACING int.
    #[getter]
    fn line_spacing_rule(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().line_spacing_rule(core))
    }
    #[setter]
    fn set_line_spacing_rule(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        const RULE: &[(i64, &'static str)] = &[
            (0, "auto"),
            (1, "auto"),
            (2, "auto"),
            (3, "atLeast"),
            (4, "exact"),
            (5, "auto"),
        ];
        let v = extract_enum(v, RULE)?;
        with_core(&self.tpl, py, |core| self.handle().set_line_spacing_rule(core, v))
            .map_err(PyValueError::new_err)
    }

    /// Widow control; defaults to True when the element is missing
    /// (python-docx semantics).
    #[getter]
    fn widow_control(&self, py: Python<'_>) -> Option<bool> {
        with_core(&self.tpl, py, |core| self.handle().widow_control(core))
    }
    #[setter]
    fn set_widow_control(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_widow_control(core, v))
            .map_err(PyValueError::new_err)
    }

    #[getter]
    fn tab_stops(&self, py: Python<'_>) -> PyTabStops {
        PyTabStops {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
        }
    }
}

// ---------------------------------------------------------------- TabStops

/// python-docx TabStops (w:pPr/w:tabs).
/// Thin wrapper over [`crate::doc::TabStops`].
#[pyclass(name = "TabStops", unsendable)]
pub struct PyTabStops {
    pub tpl: Py<PyDocxTemplate>,
    pub target: PfTarget,
}

impl PyTabStops {
    fn handle(&self) -> TabStops {
        TabStops {
            target: self.target.clone(),
        }
    }
}

#[pymethods]
impl PyTabStops {
    fn __len__(&self, py: Python<'_>) -> usize {
        with_core(&self.tpl, py, |core| self.handle().len(core))
    }

    fn __getitem__(&self, py: Python<'_>, i: usize) -> PyResult<PyTabStop> {
        if i >= self.__len__(py) {
            return Err(PyValueError::new_err("tab stop index out of range"));
        }
        Ok(PyTabStop {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
            index: i,
        })
    }

    /// Add a tab stop. position: Length; alignment/leader: WD_TAB_ALIGNMENT /
    /// WD_TAB_LEADER ints (or xml names).
    #[pyo3(signature = (position, alignment=None, leader=None))]
    fn add_tab_stop(
        &self,
        py: Python<'_>,
        position: &Bound<'_, PyAny>,
        alignment: Option<&Bound<'_, PyAny>>,
        leader: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyTabStop> {
        let emu = extract_emu(position)?
            .ok_or_else(|| PyValueError::new_err("position is required"))?;
        let al = match alignment {
            Some(a) => extract_enum(a, TAB_ALIGN)?
                .ok_or_else(|| PyValueError::new_err("invalid alignment"))?,
            None => 0,
        };
        let ld = match leader {
            Some(l) => extract_enum(l, TAB_LEADER)?
                .ok_or_else(|| PyValueError::new_err("invalid leader"))?,
            None => 0,
        };
        let n = with_core(&self.tpl, py, |core| {
            self.handle().add_tab_stop(core, emu, al, ld)
        })
        .map_err(PyValueError::new_err)?;
        Ok(PyTabStop {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
            index: n,
        })
    }

    fn clear_all(&self, py: Python<'_>) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().clear_all(core))
            .map_err(PyValueError::new_err)
    }
}

/// A single tab stop (live proxy).
/// Thin wrapper over [`crate::doc::TabStop`].
#[pyclass(name = "TabStop", unsendable)]
pub struct PyTabStop {
    pub tpl: Py<PyDocxTemplate>,
    pub target: PfTarget,
    pub index: usize,
}

impl PyTabStop {
    fn handle(&self) -> TabStop {
        TabStop {
            target: self.target.clone(),
            index: self.index,
        }
    }
}

#[pymethods]
impl PyTabStop {
    #[getter]
    fn position(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().position(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[getter]
    fn alignment(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().alignment(core))
    }
    #[getter]
    fn leader(&self, py: Python<'_>) -> Option<i64> {
        with_core(&self.tpl, py, |core| self.handle().leader(core))
    }
}

// ---------------------------------------------------------------- Column

/// A table column (w:tblGrid/w:gridCol proxy).
/// Thin wrapper over [`crate::doc::Column`].
#[pyclass(name = "Column", unsendable)]
pub struct PyTableColumn {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub col: usize,
}

impl PyTableColumn {
    fn handle(&self) -> crate::doc::Column {
        crate::doc::Column {
            index: self.index,
            col: self.col,
        }
    }
}

#[pymethods]
impl PyTableColumn {
    #[getter]
    fn width(&self, py: Python<'_>) -> Option<PyLength> {
        with_core(&self.tpl, py, |core| self.handle().width(core))
            .map(|l| PyLength { emu: l.emu })
    }
    #[setter]
    fn set_width(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        with_core(&self.tpl, py, |core| self.handle().set_width(core, emu))
            .map_err(PyValueError::new_err)
    }
}

// ---------------------------------------------------------------- cell paragraphs

/// A paragraph inside a table cell (live proxy).
/// Thin wrapper over [`crate::doc::CellParagraph`].
#[pyclass(name = "CellParagraph", unsendable)]
pub struct PyCellParagraph {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub para: usize,
}

impl PyCellParagraph {
    fn handle(&self) -> crate::doc::CellParagraph {
        crate::doc::CellParagraph {
            index: self.index,
            row: self.row,
            col: self.col,
            para: self.para,
        }
    }
}

#[pymethods]
impl PyCellParagraph {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }
    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_text(core, &v))
            .map_err(PyValueError::new_err)
    }
    #[getter]
    fn style(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().style(core))
    }
    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_style(core, &v))
            .map_err(PyValueError::new_err)
    }
}

// ---------------------------------------------------------------- Hyperlink

/// A hyperlink inside a paragraph (live proxy, read-only).
/// Thin wrapper over [`crate::doc::Hyperlink`].
#[pyclass(name = "Hyperlink", unsendable)]
pub struct PyHyperlink {
    pub tpl: Py<PyDocxTemplate>,
    pub para: usize,
    /// index among the paragraph's w:hyperlink children
    pub index: usize,
}

impl PyHyperlink {
    fn handle(&self) -> crate::doc::Hyperlink {
        crate::doc::Hyperlink {
            para: self.para,
            index: self.index,
        }
    }
}

#[pymethods]
impl PyHyperlink {
    /// Visible text of the hyperlink.
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }

    /// The URL the hyperlink points to ("" for internal jumps).
    #[getter]
    fn address(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| self.handle().address(core))
    }

    /// Fragment reference (w:anchor), e.g. a bookmark name.
    #[getter]
    fn fragment(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| self.handle().fragment(core))
    }

    /// True when the hyperlink text is broken across pages
    /// (w:lastRenderedPageBreak present).
    #[getter]
    fn contains_page_break(&self, py: Python<'_>) -> bool {
        with_core(&self.tpl, py, |core| self.handle().contains_page_break(core))
    }
}

// ---------------------------------------------------------------- nested tables

/// A table nested inside a table cell (live proxy).
/// Thin wrapper over [`crate::doc::CellTable`].
#[pyclass(name = "CellTable", unsendable)]
pub struct PyCellTable {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
    /// index among the cell's direct w:tbl children
    pub tindex: usize,
}

impl PyCellTable {
    fn handle(&self) -> crate::doc::CellTable {
        crate::doc::CellTable {
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
        }
    }
}

#[pymethods]
impl PyCellTable {
    #[getter]
    fn rows(&self, py: Python<'_>) -> Vec<PyNestedRow> {
        let n = with_core(&self.tpl, py, |core| self.handle().row_count(core));
        (0..n)
            .map(|row| PyNestedRow {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: self.row,
                col: self.col,
                tindex: self.tindex,
                nrow: row,
            })
            .collect()
    }

    fn cell(&self, py: Python<'_>, i: usize, j: usize) -> PyNestedCell {
        PyNestedCell {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
            nrow: i,
            ncol: j,
        }
    }
}

/// A row of a nested table (live proxy).
/// Thin wrapper over [`crate::doc::NestedRow`].
#[pyclass(name = "NestedRow", unsendable)]
pub struct PyNestedRow {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub tindex: usize,
    pub nrow: usize,
}

#[pymethods]
impl PyNestedRow {
    #[getter]
    fn cells(&self, py: Python<'_>) -> Vec<PyNestedCell> {
        let h = crate::doc::NestedRow {
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
            nrow: self.nrow,
        };
        let n = with_core(&self.tpl, py, |core| h.cell_count(core));
        (0..n)
            .map(|ncol| PyNestedCell {
                tpl: self.tpl.clone_ref(py),
                index: self.index,
                row: self.row,
                col: self.col,
                tindex: self.tindex,
                nrow: self.nrow,
                ncol,
            })
            .collect()
    }
}

/// A cell of a nested table (live proxy).
/// Thin wrapper over [`crate::doc::NestedCell`].
#[pyclass(name = "NestedCell", unsendable)]
pub struct PyNestedCell {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub tindex: usize,
    pub nrow: usize,
    pub ncol: usize,
}

impl PyNestedCell {
    fn handle(&self) -> crate::doc::NestedCell {
        crate::doc::NestedCell {
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
            nrow: self.nrow,
            ncol: self.ncol,
        }
    }
}

#[pymethods]
impl PyNestedCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_text(core, &v))
            .map_err(PyValueError::new_err)
    }
}

// ---------------------------------------------------------------- markers & part

/// Marker for a rendered page break (w:lastRenderedPageBreak), written by
/// Word at save time. Mirrors python-docx's RenderedPageBreak (no data).
#[pyclass(name = "RenderedPageBreak", unsendable)]
pub struct PyRenderedPageBreak {}

/// Minimal package-part facade (python-docx Part parity is intentionally
/// shallow: partname + blob).
#[pyclass(name = "Part", unsendable)]
pub struct PyPart {
    pub tpl: Py<PyDocxTemplate>,
    pub part_name: String,
}

#[pymethods]
impl PyPart {
    /// PackURI-style name, e.g. "/word/document.xml".
    #[getter]
    fn partname(&self) -> String {
        format!("/{}", self.part_name)
    }

    /// Raw bytes of the part.
    #[getter]
    fn blob(&self, py: Python<'_>) -> Option<Vec<u8>> {
        crate::docmodel::with_core(&self.tpl, py, |core| {
            core.init_docx(false).ok();
            core.package
                .as_ref()
                .and_then(|pkg| pkg.get(&self.part_name).map(|b| b.to_vec()))
        })
    }

    /// Relationships of the part, as a list of dicts with keys "rId",
    /// "type", "target", "is_external".
    #[getter]
    fn rels(&self, py: Python<'_>) -> Vec<Py<pyo3::types::PyDict>> {
        with_core(&self.tpl, py, |core| crate::doc::part_rels(core, &self.part_name))
            .into_iter()
            .map(|(id, rel_type, target, is_external)| {
                let d = pyo3::types::PyDict::new(py);
                let _ = d.set_item("rId", id);
                let _ = d.set_item("type", rel_type);
                let _ = d.set_item("target", target);
                let _ = d.set_item("is_external", is_external);
                d.unbind()
            })
            .collect()
    }

    /// Content type of the part (Override first, then Default by extension).
    #[getter]
    fn content_type(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| {
            crate::doc::part_content_type(core, &self.part_name)
        })
    }
}

// ---------------------------------------------------------------- fields

/// A field code in a paragraph (live proxy). Covers both w:fldSimple and
/// complex (fldChar begin/.../end) fields.
/// Thin wrapper over [`crate::doc::Field`].
#[pyclass(name = "Field", unsendable)]
pub struct PyField {
    pub tpl: Py<PyDocxTemplate>,
    pub para: usize,
    /// index among the paragraph's fields, in document order
    pub index: usize,
}

impl PyField {
    fn handle(&self) -> crate::doc::Field {
        crate::doc::Field {
            para: self.para,
            index: self.index,
        }
    }
}

#[pymethods]
impl PyField {
    /// "simple" (w:fldSimple) or "complex" (fldChar-based).
    #[getter]
    fn kind(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().kind(core))
    }

    /// The field instruction (e.g. `PAGE \* MERGEFORMAT`), trimmed.
    #[getter]
    fn instr(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().instr(core))
    }

    #[setter]
    fn set_instr(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_instr(core, &v))
            .map_err(PyValueError::new_err)
    }

    /// The cached (last-rendered) field result text.
    #[getter]
    fn text(&self, py: Python<'_>) -> Option<String> {
        with_core(&self.tpl, py, |core| self.handle().text(core))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        with_core(&self.tpl, py, |core| self.handle().set_text(core, &v))
            .map_err(PyValueError::new_err)
    }
}
