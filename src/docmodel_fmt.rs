//! python-docx parity formatting proxies: Font, ColorFormat, ParagraphFormat,
//! TabStops/TabStop, table Column, and cell paragraphs.
//!
//! Enum values follow python-docx: getters return ints (python-docx enums are
//! int subclasses, so `== WD_*` comparisons hold); setters accept ints or the
//! xml vocabulary strings. Tri-state booleans follow python-docx CT_OnOff
//! semantics: get -> None when the element is missing; set None removes the
//! element, True writes a bare element, False writes `w:val="0"`.

use crate::docmodel::{ensure_rpr, PyParagraph, PyRun, PyStyle};
use crate::pyclasses::{PyDocxTemplate, PyLength};
use crate::xmldom::{Element, Node};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyFloat;

// ---------------------------------------------------------------- helpers

/// get-or-create a direct child element (appended at the end).
pub(crate) fn ensure_child<'a>(parent: &'a mut Element, tag: &str) -> &'a mut Element {
    if parent.find(tag).is_none() {
        parent.children.push(Node::Elem(Element::new(tag)));
    }
    parent.find_mut(tag).unwrap()
}

/// w:pPr of a paragraph (must be the first child).
pub(crate) fn ensure_ppr(p: &mut Element) -> &mut Element {
    // lifetime elided: single input reference
    if p.find("w:pPr").is_none() {
        p.children.insert(0, Node::Elem(Element::new("w:pPr")));
    }
    p.find_mut("w:pPr").unwrap()
}

/// w:trPr of a table row (must be the first child).
pub(crate) fn ensure_trpr(tr: &mut Element) -> &mut Element {
    if tr.find("w:trPr").is_none() {
        tr.children.insert(0, Node::Elem(Element::new("w:trPr")));
    }
    tr.find_mut("w:trPr").unwrap()
}

/// w:tblPr of a table (must be the first child).
pub(crate) fn ensure_tblpr(tbl: &mut Element) -> &mut Element {
    if tbl.find("w:tblPr").is_none() {
        tbl.children.insert(0, Node::Elem(Element::new("w:tblPr")));
    }
    tbl.find_mut("w:tblPr").unwrap()
}

fn remove_attr(el: &mut Element, name: &str) {
    el.attrs.retain(|(k, _)| k != name);
}

fn remove_child(parent: &mut Element, tag: &str) {
    parent
        .children
        .retain(|c| !matches!(c, Node::Elem(e) if e.name == tag));
}

/// tri-state bool read: missing element -> None; missing w:val -> true.
pub(crate) fn tri_get(container: &Element, tag: &str) -> Option<bool> {
    container.find(tag).map(|e| {
        !matches!(
            e.get_attr("w:val"),
            Some("0") | Some("false") | Some("off") | Some("none")
        )
    })
}

/// tri-state bool write: None -> remove, Some(true) -> bare, Some(false) ->
/// `w:val="0"` (python-docx semantics).
pub(crate) fn tri_set(container: &mut Element, tag: &str, v: Option<bool>) {
    match v {
        None => remove_child(container, tag),
        Some(on) => {
            let el = ensure_child(container, tag);
            if on {
                remove_attr(el, "w:val");
            } else {
                el.set_attr("w:val", "0");
            }
        }
    }
}

fn attr_get(container: &Element, tag: &str, attr: &str) -> Option<String> {
    container
        .find(tag)
        .and_then(|e| e.get_attr(attr))
        .map(|s| s.to_string())
}

/// set an attribute on a child element; None removes the whole element.
fn attr_set(container: &mut Element, tag: &str, attr: &str, val: Option<&str>) {
    match val {
        None => remove_child(container, tag),
        Some(v) => ensure_child(container, tag).set_attr(attr, v),
    }
}

// ---- enum tables (int <-> xml vocabulary) ----

fn xml_of(table: &'static [(i64, &'static str)], v: i64) -> Option<&'static str> {
    table.iter().find(|(i, _)| *i == v).map(|(_, s)| *s)
}

fn int_of(table: &[(i64, &str)], s: &str) -> Option<i64> {
    table.iter().find(|(_, x)| *x == s).map(|(i, _)| *i)
}

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

fn enum_get(container: &Element, tag: &str, attr: &str, table: &[(i64, &str)]) -> Option<i64> {
    attr_get(container, tag, attr).and_then(|s| int_of(table, &s))
}

fn enum_set(
    container: &mut Element,
    tag: &str,
    attr: &str,
    table: &'static [(i64, &'static str)],
    v: Option<i64>,
) {
    match v.and_then(|i| xml_of(table, i)) {
        None => remove_child(container, tag),
        Some(s) => attr_set(container, tag, attr, Some(s)),
    }
}

// ---- length (PyLength EMU <-> twips) ----

fn len_get(container: &Element, tag: &str, attr: &str) -> Option<PyLength> {
    attr_get(container, tag, attr)
        .and_then(|s| s.parse::<i64>().ok())
        .map(|twips| PyLength { emu: twips * 635 })
}

/// v: EMU. None removes the attribute (element kept if it has other attrs).
fn len_set(container: &mut Element, tag: &str, attr: &str, v: Option<i64>) {
    match v {
        Some(emu) => {
            let twips = (emu / 635).to_string();
            ensure_child(container, tag).set_attr(attr, &twips);
        }
        None => {
            if let Some(el) = container.find_mut(tag) {
                remove_attr(el, attr);
            }
        }
    }
}

fn extract_emu(v: &Bound<'_, PyAny>) -> PyResult<Option<i64>> {
    crate::pyclasses::extract_length_pub(v)
}

// alignment: WD_ALIGN_PARAGRAPH
pub(crate) const ALIGN: &[(i64, &str)] = &[
    (0, "left"),
    (1, "center"),
    (2, "right"),
    (3, "both"),
    (4, "distribute"),
    (5, "mediumKashida"),
    (7, "highKashida"),
    (8, "lowKashida"),
    (9, "thaiDistribute"),
];

// WD_UNDERLINE
const UNDERLINE: &[(i64, &str)] = &[
    (0, "none"),
    (1, "single"),
    (2, "words"),
    (3, "double"),
    (4, "dotted"),
    (6, "thick"),
    (7, "dash"),
    (9, "dotDash"),
    (10, "dotDotDash"),
    (11, "wave"),
    (20, "dottedHeavy"),
    (23, "dashedHeavy"),
    (25, "dashDotHeavy"),
    (26, "dashDotDotHeavy"),
    (27, "wavyHeavy"),
    (39, "dashLong"),
    (43, "wavyDouble"),
    (55, "dashLongHeavy"),
];

// WD_COLOR_INDEX (highlight)
const HIGHLIGHT: &[(i64, &str)] = &[
    (0, "default"),
    (1, "black"),
    (2, "blue"),
    (3, "cyan"),
    (4, "green"),
    (5, "magenta"),
    (6, "red"),
    (7, "yellow"),
    (8, "white"),
    (9, "darkBlue"),
    (10, "darkCyan"),
    (11, "darkGreen"),
    (12, "darkMagenta"),
    (13, "darkRed"),
    (14, "darkYellow"),
    (15, "darkGray"),
    (16, "lightGray"),
];

// WD_TAB_ALIGNMENT
const TAB_ALIGN: &[(i64, &str)] = &[
    (0, "left"),
    (1, "center"),
    (2, "right"),
    (3, "decimal"),
    (4, "bar"),
    (6, "list"),
    (101, "clear"),
];

// WD_TAB_LEADER
const TAB_LEADER: &[(i64, &str)] = &[
    (0, "none"),
    (1, "dot"),
    (2, "hyphen"),
    (3, "underscore"),
    (4, "heavy"),
    (5, "middleDot"),
];

// ---------------------------------------------------------------- Font

/// What a Font/ColorFormat is attached to.
#[derive(Clone)]
pub enum FontTarget {
    Run { para: usize, index: usize },
    Style { style_id: String },
}

/// python-docx Font: full run/style character properties (tri-state booleans).
#[pyclass(name = "Font", unsendable)]
pub struct PyFont {
    pub tpl: Py<PyDocxTemplate>,
    pub target: FontTarget,
}

impl PyFont {
    pub(crate) fn edit_rpr<R>(
        &self,
        py: Python<'_>,
        f: impl FnOnce(&mut Element) -> R,
    ) -> PyResult<R> {
        match &self.target {
            FontTarget::Run { para, index } => {
                let run = PyRun {
                    tpl: self.tpl.clone_ref(py),
                    para: *para,
                    index: *index,
                };
                run.edit(py, |r| f(ensure_rpr(r)))
            }
            FontTarget::Style { style_id } => {
                let st = PyStyle {
                    tpl: self.tpl.clone_ref(py),
                    style_id: style_id.clone(),
                };
                st.edit(py, |el| {
                    if el.find("w:rPr").is_none() {
                        el.children.push(Node::Elem(Element::new("w:rPr")));
                    }
                    f(el.find_mut("w:rPr").unwrap())
                })
            }
        }
    }

    pub(crate) fn read_rpr<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        match &self.target {
            FontTarget::Run { para, index } => {
                let run = PyRun {
                    tpl: self.tpl.clone_ref(py),
                    para: *para,
                    index: *index,
                };
                run.read(py, |r| r.find("w:rPr").map(|e| f(e))).flatten()
            }
            FontTarget::Style { style_id } => {
                let st = PyStyle {
                    tpl: self.tpl.clone_ref(py),
                    style_id: style_id.clone(),
                };
                st.read(py, |el| el.find("w:rPr").map(|e| f(e))).flatten()
            }
        }
    }
}

/// Generate tri-state bool property pairs on PyFont (plus extra methods).
macro_rules! font_tri {
    ($($name:ident / $setter:ident => $tag:literal),* $(,)?; $($extra:item)*) => {
        #[pymethods]
        impl PyFont {
            $(
            #[getter]
            fn $name(&self, py: Python<'_>) -> Option<bool> {
                self.read_rpr(py, |rpr| tri_get(rpr, $tag)).flatten()
            }
            #[setter]
            fn $setter(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
                self.edit_rpr(py, |rpr| tri_set(rpr, $tag, v))
            }
            )*
            $($extra)*
        }
    };
}

font_tri! {
    bold / set_bold => "w:b",
    cs_bold / set_cs_bold => "w:bCs",
    italic / set_italic => "w:i",
    cs_italic / set_cs_italic => "w:iCs",
    all_caps / set_all_caps => "w:caps",
    small_caps / set_small_caps => "w:smallCaps",
    strike / set_strike => "w:strike",
    double_strike / set_double_strike => "w:dstrike",
    outline / set_outline => "w:outline",
    shadow / set_shadow => "w:shadow",
    emboss / set_emboss => "w:emboss",
    imprint / set_imprint => "w:imprint",
    no_proof / set_no_proof => "w:noProof",
    snap_to_grid / set_snap_to_grid => "w:snapToGrid",
    hidden / set_hidden => "w:vanish",
    web_hidden / set_web_hidden => "w:webHidden",
    spec_vanish / set_spec_vanish => "w:specVanish",
    rtl / set_rtl => "w:rtl",
    complex_script / set_complex_script => "w:cs",
    math / set_math => "w:oMath",
;
    #[getter]
    fn subscript(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:vertAlign")
                .map(|e| e.get_attr("w:val") == Some("subscript"))
        })
        .flatten()
    }
    #[setter]
    fn set_subscript(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
        self.edit_rpr(py, |rpr| match v {
            Some(true) => attr_set(rpr, "w:vertAlign", "w:val", Some("subscript")),
            Some(false) => {
                if attr_get(rpr, "w:vertAlign", "w:val").as_deref() == Some("subscript") {
                    remove_child(rpr, "w:vertAlign");
                }
            }
            None => remove_child(rpr, "w:vertAlign"),
        })
    }
    #[getter]
    fn superscript(&self, py: Python<'_>) -> Option<bool> {
        self.read_rpr(py, |rpr| {
            rpr.find("w:vertAlign")
                .map(|e| e.get_attr("w:val") == Some("superscript"))
        })
        .flatten()
    }
    #[setter]
    fn set_superscript(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
        self.edit_rpr(py, |rpr| match v {
            Some(true) => attr_set(rpr, "w:vertAlign", "w:val", Some("superscript")),
            Some(false) => {
                if attr_get(rpr, "w:vertAlign", "w:val").as_deref() == Some("superscript") {
                    remove_child(rpr, "w:vertAlign");
                }
            }
            None => remove_child(rpr, "w:vertAlign"),
        })
    }

    /// Font size as a Length (EMU); w:sz stores half-points.
    #[getter]
    fn size(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_rpr(py, |rpr| {
            attr_get(rpr, "w:sz", "w:val")
                .and_then(|s| s.parse::<i64>().ok())
                .map(|hp| PyLength { emu: hp * 12700 / 2 })
        })
        .flatten()
    }
    #[setter]
    fn set_size(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        self.edit_rpr(py, |rpr| match emu {
            Some(e) => {
                let hp = ((e * 2) / 12700).to_string();
                attr_set(rpr, "w:sz", "w:val", Some(&hp));
            }
            None => remove_child(rpr, "w:sz"),
        })
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> Option<String> {
        self.read_rpr(py, |rpr| attr_get(rpr, "w:rFonts", "w:ascii"))
            .flatten()
    }
    #[setter]
    fn set_name(&self, py: Python<'_>, v: Option<String>) -> PyResult<()> {
        self.edit_rpr(py, |rpr| match v {
            Some(name) => {
                let rf = ensure_child(rpr, "w:rFonts");
                rf.set_attr("w:ascii", &name);
                rf.set_attr("w:hAnsi", &name);
            }
            None => remove_child(rpr, "w:rFonts"),
        })
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
        self.read_rpr(py, |rpr| enum_get(rpr, "w:highlight", "w:val", HIGHLIGHT))
            .flatten()
    }
    #[setter]
    fn set_highlight_color(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = extract_enum(v, HIGHLIGHT)?;
        self.edit_rpr(py, |rpr| enum_set(rpr, "w:highlight", "w:val", HIGHLIGHT, v))
    }

    /// Underline as a WD_UNDERLINE int; True -> single, False -> none.
    #[getter]
    fn underline(&self, py: Python<'_>) -> Option<i64> {
        self.read_rpr(py, |rpr| enum_get(rpr, "w:u", "w:val", UNDERLINE))
            .flatten()
    }
    #[setter]
    fn set_underline(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = if let Ok(b) = v.extract::<bool>() {
            Some(if b { 1 } else { 0 })
        } else {
            extract_enum(v, UNDERLINE)?
        };
        self.edit_rpr(py, |rpr| enum_set(rpr, "w:u", "w:val", UNDERLINE, v))
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
        let font = PyFont {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
        };
        font.read_rpr(py, |rpr| attr_get(rpr, "w:color", "w:val"))
            .flatten()
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
        let font = PyFont {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
        };
        font.edit_rpr(py, |rpr| attr_set(rpr, "w:color", "w:val", hex.as_deref()))
    }
}

// ---------------------------------------------------------------- ParagraphFormat

/// What a ParagraphFormat/TabStops is attached to.
#[derive(Clone)]
pub enum PfTarget {
    Para { index: usize },
    Style { style_id: String },
}

/// python-docx ParagraphFormat (w:pPr properties).
#[pyclass(name = "ParagraphFormat", unsendable)]
pub struct PyParagraphFormat {
    pub tpl: Py<PyDocxTemplate>,
    pub target: PfTarget,
}

impl PyParagraphFormat {
    fn edit_ppr<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        match &self.target {
            PfTarget::Para { index } => {
                let p = PyParagraph {
                    tpl: self.tpl.clone_ref(py),
                    index: *index,
                };
                p.edit(py, |el| f(ensure_ppr(el)))
            }
            PfTarget::Style { style_id } => {
                let st = PyStyle {
                    tpl: self.tpl.clone_ref(py),
                    style_id: style_id.clone(),
                };
                st.edit(py, |el| {
                    if el.find("w:pPr").is_none() {
                        el.children.push(Node::Elem(Element::new("w:pPr")));
                    }
                    f(el.find_mut("w:pPr").unwrap())
                })
            }
        }
    }

    fn read_ppr<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        match &self.target {
            PfTarget::Para { index } => {
                let p = PyParagraph {
                    tpl: self.tpl.clone_ref(py),
                    index: *index,
                };
                p.read(py, |el| el.find("w:pPr").map(|e| f(e))).flatten()
            }
            PfTarget::Style { style_id } => {
                let st = PyStyle {
                    tpl: self.tpl.clone_ref(py),
                    style_id: style_id.clone(),
                };
                st.read(py, |el| el.find("w:pPr").map(|e| f(e))).flatten()
            }
        }
    }
}

/// tri-state bool properties on ParagraphFormat (plus extra methods).
macro_rules! pf_tri {
    ($($name:ident / $setter:ident => $tag:literal),* $(,)?; $($extra:item)*) => {
        #[pymethods]
        impl PyParagraphFormat {
            $(
            #[getter]
            fn $name(&self, py: Python<'_>) -> Option<bool> {
                self.read_ppr(py, |ppr| tri_get(ppr, $tag)).flatten()
            }
            #[setter]
            fn $setter(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
                self.edit_ppr(py, |ppr| tri_set(ppr, $tag, v))
            }
            )*
            $($extra)*
        }
    };
}

pf_tri! {
    keep_together / set_keep_together => "w:keepLines",
    keep_with_next / set_keep_with_next => "w:keepNext",
    page_break_before / set_page_break_before => "w:pageBreakBefore",
;
    /// Alignment as a WD_ALIGN_PARAGRAPH int (xml name also accepted on set).
    #[getter]
    pub(crate) fn alignment(&self, py: Python<'_>) -> Option<i64> {
        self.read_ppr(py, |ppr| enum_get(ppr, "w:jc", "w:val", ALIGN))
            .flatten()
    }
    #[setter]
    pub(crate) fn set_alignment(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = extract_enum(v, ALIGN)?;
        self.edit_ppr(py, |ppr| enum_set(ppr, "w:jc", "w:val", ALIGN, v))
    }

    #[getter]
    fn left_indent(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_ppr(py, |ppr| len_get(ppr, "w:ind", "w:left")).flatten()
    }
    #[setter]
    fn set_left_indent(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        self.edit_ppr(py, |ppr| len_set(ppr, "w:ind", "w:left", emu))
    }
    #[getter]
    fn right_indent(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_ppr(py, |ppr| len_get(ppr, "w:ind", "w:right")).flatten()
    }
    #[setter]
    fn set_right_indent(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        self.edit_ppr(py, |ppr| len_set(ppr, "w:ind", "w:right", emu))
    }

    /// First-line indent; negative values become a hanging indent.
    #[getter]
    fn first_line_indent(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_ppr(py, |ppr| {
            if let Some(l) = len_get(ppr, "w:ind", "w:firstLine") {
                return Some(l);
            }
            len_get(ppr, "w:ind", "w:hanging").map(|l| PyLength { emu: -l.emu })
        })
        .flatten()
    }
    #[setter]
    fn set_first_line_indent(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        self.edit_ppr(py, |ppr| {
            let ind = ensure_child(ppr, "w:ind");
            remove_attr(ind, "w:firstLine");
            remove_attr(ind, "w:hanging");
            match emu {
                Some(e) if e >= 0 => ind.set_attr("w:firstLine", &(e / 635).to_string()),
                Some(e) => ind.set_attr("w:hanging", &(-e / 635).to_string()),
                None => {}
            }
        })
    }

    #[getter]
    fn space_before(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_ppr(py, |ppr| len_get(ppr, "w:spacing", "w:before"))
            .flatten()
    }
    #[setter]
    fn set_space_before(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        self.edit_ppr(py, |ppr| len_set(ppr, "w:spacing", "w:before", emu))
    }
    #[getter]
    fn space_after(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_ppr(py, |ppr| len_get(ppr, "w:spacing", "w:after"))
            .flatten()
    }
    #[setter]
    fn set_space_after(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        self.edit_ppr(py, |ppr| len_set(ppr, "w:spacing", "w:after", emu))
    }

    /// Line spacing: a float multiple (2.0 = double) when the rule is auto,
    /// otherwise a Length (exact / at-least).
    #[getter]
    fn line_spacing(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        let (line, rule) = self.read_ppr(py, |ppr| {
            let sp = ppr.find("w:spacing")?;
            let line = sp.get_attr("w:line")?.parse::<i64>().ok()?;
            let rule = sp.get_attr("w:lineRule").unwrap_or("auto").to_string();
            Some((line, rule))
        })??;
        let v: Py<PyAny> = if rule == "auto" {
            PyFloat::new(py, line as f64 / 240.0).unbind().into_any()
        } else {
            Py::new(
                py,
                PyLength {
                    emu: line * 635,
                },
            )
            .ok()?
            .into_any()
        };
        Some(v)
    }
    #[setter]
    fn set_line_spacing(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        if v.is_none() {
            return self.edit_ppr(py, |ppr| {
                if let Some(sp) = ppr.find_mut("w:spacing") {
                    remove_attr(sp, "w:line");
                    remove_attr(sp, "w:lineRule");
                }
            });
        }
        // float -> multiple (lineRule=auto); Length/int -> exact twips
        if let Ok(f) = v.extract::<f64>() {
            let line = (f * 240.0).round() as i64;
            return self.edit_ppr(py, |ppr| {
                let sp = ensure_child(ppr, "w:spacing");
                sp.set_attr("w:line", &line.to_string());
                sp.set_attr("w:lineRule", "auto");
            });
        }
        let emu = extract_emu(v)?;
        self.edit_ppr(py, |ppr| {
            let sp = ensure_child(ppr, "w:spacing");
            match emu {
                Some(e) => {
                    sp.set_attr("w:line", &(e / 635).to_string());
                    // keep an existing atLeast rule, else exact
                    if sp.get_attr("w:lineRule") != Some("atLeast") {
                        sp.set_attr("w:lineRule", "exact");
                    }
                }
                None => {
                    remove_attr(sp, "w:line");
                    remove_attr(sp, "w:lineRule");
                }
            }
        })
    }

    /// Line spacing rule as a WD_LINE_SPACING int.
    #[getter]
    fn line_spacing_rule(&self, py: Python<'_>) -> Option<i64> {
        self.read_ppr(py, |ppr| {
            let sp = ppr.find("w:spacing")?;
            let rule = sp.get_attr("w:lineRule").unwrap_or("auto");
            let line = sp
                .get_attr("w:line")
                .and_then(|s| s.parse::<i64>().ok())
                .unwrap_or(240);
            Some(match rule {
                "exact" => 4,
                "atLeast" => 3,
                _ => match line {
                    240 => 0,
                    360 => 1,
                    480 => 2,
                    _ => 5,
                },
            })
        })
        .flatten()
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
        self.edit_ppr(py, |ppr| {
            let sp = ensure_child(ppr, "w:spacing");
            match v {
                None => {
                    remove_attr(sp, "w:lineRule");
                }
                Some(0) => {
                    sp.set_attr("w:lineRule", "auto");
                    sp.set_attr("w:line", "240");
                }
                Some(1) => {
                    sp.set_attr("w:lineRule", "auto");
                    sp.set_attr("w:line", "360");
                }
                Some(2) => {
                    sp.set_attr("w:lineRule", "auto");
                    sp.set_attr("w:line", "480");
                }
                Some(3) => sp.set_attr("w:lineRule", "atLeast"),
                Some(4) => sp.set_attr("w:lineRule", "exact"),
                Some(_) => sp.set_attr("w:lineRule", "auto"),
            }
        })
    }

    /// Widow control; defaults to True when the element is missing
    /// (python-docx semantics).
    #[getter]
    fn widow_control(&self, py: Python<'_>) -> Option<bool> {
        Some(
            self.read_ppr(py, |ppr| tri_get(ppr, "w:widowControl"))
                .flatten()
                .unwrap_or(true),
        )
    }
    #[setter]
    fn set_widow_control(&self, py: Python<'_>, v: Option<bool>) -> PyResult<()> {
        self.edit_ppr(py, |ppr| tri_set(ppr, "w:widowControl", v))
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
#[pyclass(name = "TabStops", unsendable)]
pub struct PyTabStops {
    pub tpl: Py<PyDocxTemplate>,
    pub target: PfTarget,
}

impl PyTabStops {
    fn pf(&self, py: Python<'_>) -> PyParagraphFormat {
        PyParagraphFormat {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
        }
    }
}

#[pymethods]
impl PyTabStops {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.pf(py)
            .read_ppr(py, |ppr| {
                ppr.find("w:tabs")
                    .map(|t| {
                        t.children
                            .iter()
                            .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tab"))
                            .count()
                    })
                    .unwrap_or(0)
            })
            .unwrap_or(0)
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
        let n = self.pf(py).edit_ppr(py, |ppr| {
            let tabs = ensure_child(ppr, "w:tabs");
            let n = tabs
                .children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tab"))
                .count();
            let mut tab = Element::new("w:tab");
            tab.set_attr("w:val", xml_of(TAB_ALIGN, al).unwrap());
            tab.set_attr("w:leader", xml_of(TAB_LEADER, ld).unwrap());
            tab.set_attr("w:pos", &(emu / 635).to_string());
            tabs.children.push(Node::Elem(tab));
            n
        })?;
        Ok(PyTabStop {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
            index: n,
        })
    }

    fn clear_all(&self, py: Python<'_>) -> PyResult<()> {
        self.pf(py).edit_ppr(py, |ppr| remove_child(ppr, "w:tabs"))
    }
}

/// A single tab stop (live proxy).
#[pyclass(name = "TabStop", unsendable)]
pub struct PyTabStop {
    pub tpl: Py<PyDocxTemplate>,
    pub target: PfTarget,
    pub index: usize,
}

impl PyTabStop {
    fn read_tab<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let pf = PyParagraphFormat {
            tpl: self.tpl.clone_ref(py),
            target: self.target.clone(),
        };
        let index = self.index;
        pf.read_ppr(py, |ppr| {
            ppr.find("w:tabs").and_then(|tabs| {
                tabs.children
                    .iter()
                    .filter_map(|c| match c {
                        Node::Elem(e) if e.name == "w:tab" => Some(e),
                        _ => None,
                    })
                    .nth(index)
                    .map(|t| f(t))
            })
        })
        .flatten()
    }
}

#[pymethods]
impl PyTabStop {
    #[getter]
    fn position(&self, py: Python<'_>) -> Option<PyLength> {
        self.read_tab(py, |t| {
            t.get_attr("w:pos")
                .and_then(|s| s.parse::<i64>().ok())
                .map(|tw| PyLength { emu: tw * 635 })
        })
        .flatten()
    }
    #[getter]
    fn alignment(&self, py: Python<'_>) -> Option<i64> {
        self.read_tab(py, |t| t.get_attr("w:val").and_then(|s| int_of(TAB_ALIGN, s)))
            .flatten()
    }
    #[getter]
    fn leader(&self, py: Python<'_>) -> Option<i64> {
        self.read_tab(py, |t| t.get_attr("w:leader").and_then(|s| int_of(TAB_LEADER, s)))
            .flatten()
    }
}

// ---------------------------------------------------------------- Column

/// A table column (w:tblGrid/w:gridCol proxy).
#[pyclass(name = "Column", unsendable)]
pub struct PyTableColumn {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub col: usize,
}

#[pymethods]
impl PyTableColumn {
    #[getter]
    fn width(&self, py: Python<'_>) -> Option<PyLength> {
        let tbl = crate::docmodel::PyTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
        };
        let col = self.col;
        tbl.read(py, |t| {
            t.find("w:tblGrid")
                .and_then(|g| crate::docmodel::nth_direct_ref(g, "w:gridCol", col))
                .and_then(|gc| gc.get_attr("w:w"))
                .and_then(|s| s.parse::<i64>().ok())
                .map(|tw| PyLength { emu: tw * 635 })
        })
        .flatten()
    }
    #[setter]
    fn set_width(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let emu = extract_emu(v)?;
        let tbl = crate::docmodel::PyTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
        };
        let col = self.col;
        tbl.edit(py, |t| {
            if t.find("w:tblGrid").is_none() {
                // tblGrid must follow tblPr
                let pos = if t.find("w:tblPr").is_some() { 1 } else { 0 };
                t.children.insert(pos, Node::Elem(Element::new("w:tblGrid")));
            }
            let grid = t.find_mut("w:tblGrid").unwrap();
            while grid
                .children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:gridCol"))
                .count()
                <= col
            {
                grid.children.push(Node::Elem(Element::new("w:gridCol")));
            }
            if let Some(gc) = grid
                .children
                .iter_mut()
                .filter_map(|c| match c {
                    Node::Elem(e) if e.name == "w:gridCol" => Some(e),
                    _ => None,
                })
                .nth(col)
            {
                match emu {
                    Some(e) => gc.set_attr("w:w", &(e / 635).to_string()),
                    None => remove_attr(gc, "w:w"),
                }
            }
        })
    }
}

// ---------------------------------------------------------------- cell paragraphs

/// A paragraph inside a table cell (live proxy).
#[pyclass(name = "CellParagraph", unsendable)]
pub struct PyCellParagraph {
    pub tpl: Py<PyDocxTemplate>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub para: usize,
}

impl PyCellParagraph {
    fn cell(&self, py: Python<'_>) -> crate::docmodel::PyCell {
        crate::docmodel::PyCell {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
        }
    }

    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        let cell = self.cell(py);
        let para = self.para;
        cell.edit(py, |tc| {
            let mut out = None;
            let mut n = 0usize;
            for c in tc.children.iter_mut() {
                if let Node::Elem(e) = c {
                    if e.name == "w:p" {
                        if n == para {
                            out = Some(f(e));
                            break;
                        }
                        n += 1;
                    }
                }
            }
            out.ok_or_else(|| PyValueError::new_err("paragraph not found"))
        })?
    }

    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let cell = self.cell(py);
        let para = self.para;
        crate::docmodel::with_core(&cell.tpl, py, |core| {
            crate::docmodel::read_body(core, |body| {
                crate::docmodel::nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| crate::docmodel::nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| crate::docmodel::nth_direct_ref(r, "w:tc", self.col))
                    .and_then(|tc| crate::docmodel::nth_direct_ref(tc, "w:p", para))
                    .map(|p| f(p))
            })
            .flatten()
        })
    }
}

#[pymethods]
impl PyCellParagraph {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        self.read(py, |p| crate::docmodel::element_text(p))
            .unwrap_or_default()
    }
    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |p| {
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
        let sid = crate::docmodel::with_core(&self.tpl, py, |core| {
            crate::subdocbuilder::resolve_style_id(core, &v)
        });
        self.edit(py, |p| {
            let ppr = ensure_ppr(p);
            if let Some(ps) = ppr.find_mut("w:pStyle") {
                ps.set_attr("w:val", &sid);
            } else {
                let mut ps = Element::new("w:pStyle");
                ps.set_attr("w:val", &sid);
                ppr.children.insert(0, Node::Elem(ps));
            }
        })
    }
}

// ---------------------------------------------------------------- Hyperlink

/// A hyperlink inside a paragraph (live proxy, read-only).
#[pyclass(name = "Hyperlink", unsendable)]
pub struct PyHyperlink {
    pub tpl: Py<PyDocxTemplate>,
    pub para: usize,
    /// index among the paragraph's w:hyperlink children
    pub index: usize,
}

impl PyHyperlink {
    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let p = PyParagraph {
            tpl: self.tpl.clone_ref(py),
            index: self.para,
        };
        let index = self.index;
        p.read(py, |para| {
            crate::docmodel::nth_direct_ref(para, "w:hyperlink", index).map(|h| f(h))
        })
        .flatten()
    }
}

#[pymethods]
impl PyHyperlink {
    /// Visible text of the hyperlink.
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        self.read(py, |h| crate::docmodel::element_text(h))
            .unwrap_or_default()
    }

    /// The URL the hyperlink points to ("" for internal jumps).
    #[getter]
    fn address(&self, py: Python<'_>) -> String {
        let rid = self
            .read(py, |h| h.get_attr("r:id").map(|s| s.to_string()))
            .flatten();
        let Some(rid) = rid else { return String::new() };
        crate::docmodel::with_core(&self.tpl, py, |core| {
            core.init_docx(false).ok();
            core.package
                .as_ref()
                .map(|pkg| {
                    pkg.rels(crate::template::DOCUMENT_PART)
                        .rels
                        .iter()
                        .find(|r| r.id == rid)
                        .map(|r| r.target.clone())
                        .unwrap_or_default()
                })
                .unwrap_or_default()
        })
    }

    /// Fragment reference (w:anchor), e.g. a bookmark name.
    #[getter]
    fn fragment(&self, py: Python<'_>) -> String {
        self.read(py, |h| {
            h.get_attr("w:anchor").map(|s| s.to_string())
        })
        .flatten()
        .unwrap_or_default()
    }

    /// True when the hyperlink text is broken across pages
    /// (w:lastRenderedPageBreak present).
    #[getter]
    fn contains_page_break(&self, py: Python<'_>) -> bool {
        self.read(py, |h| {
            let mut out = Vec::new();
            h.iter_descendants("w:lastRenderedPageBreak", &mut out);
            !out.is_empty()
        })
        .unwrap_or(false)
    }
}

// ---------------------------------------------------------------- nested tables

/// A table nested inside a table cell (live proxy).
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
    fn parent_cell(&self, py: Python<'_>) -> crate::docmodel::PyCell {
        crate::docmodel::PyCell {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
        }
    }

    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        let cell = self.parent_cell(py);
        let tindex = self.tindex;
        cell.edit(py, |tc| {
            let mut out = None;
            let mut n = 0usize;
            for c in tc.children.iter_mut() {
                if let Node::Elem(e) = c {
                    if e.name == "w:tbl" {
                        if n == tindex {
                            out = Some(f(e));
                            break;
                        }
                        n += 1;
                    }
                }
            }
            out.ok_or_else(|| PyValueError::new_err("table not found"))
        })?
    }

    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let cell = self.parent_cell(py);
        let tindex = self.tindex;
        crate::docmodel::with_core(&cell.tpl, py, |core| {
            crate::docmodel::read_body(core, |body| {
                crate::docmodel::nth_direct_ref(body, "w:tbl", self.index)
                    .and_then(|t| crate::docmodel::nth_direct_ref(t, "w:tr", self.row))
                    .and_then(|r| crate::docmodel::nth_direct_ref(r, "w:tc", self.col))
                    .and_then(|tc| crate::docmodel::nth_direct_ref(tc, "w:tbl", tindex))
                    .map(|t| f(t))
            })
            .flatten()
        })
    }
}

#[pymethods]
impl PyCellTable {
    #[getter]
    fn rows(&self, py: Python<'_>) -> Vec<PyNestedRow> {
        let n = self
            .read(py, |t| {
                t.children
                    .iter()
                    .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tr"))
                    .count()
            })
            .unwrap_or(0);
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
        let t = PyCellTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
        };
        let nrow = self.nrow;
        let n = t
            .read(py, |tbl| {
                crate::docmodel::nth_direct_ref(tbl, "w:tr", nrow).map(|tr| {
                    tr.children
                        .iter()
                        .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                        .count()
                })
            })
            .flatten()
            .unwrap_or(0);
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

#[pymethods]
impl PyNestedCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> String {
        let t = PyCellTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
        };
        let (nrow, ncol) = (self.nrow, self.ncol);
        t.read(py, |tbl| {
            crate::docmodel::nth_direct_ref(tbl, "w:tr", nrow)
                .and_then(|tr| crate::docmodel::nth_direct_ref(tr, "w:tc", ncol))
                .map(|tc| crate::docmodel::element_text(tc))
        })
        .flatten()
        .unwrap_or_default()
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        let t = PyCellTable {
            tpl: self.tpl.clone_ref(py),
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
        };
        let (nrow, ncol) = (self.nrow, self.ncol);
        t.edit(py, |tbl| {
            if let Some(tr) = crate::docmodel_add::nth_direct(tbl, "w:tr", nrow) {
                if let Some(tc) = crate::docmodel_add::nth_direct(tr, "w:tc", ncol) {
                    tc.children
                        .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:p"));
                    let mut p = Element::new("w:p");
                    let mut r = Element::new("w:r");
                    let mut wt = Element::new("w:t");
                    wt.set_attr("xml:space", "preserve");
                    wt.children.push(Node::Text(v.clone()));
                    r.children.push(Node::Elem(wt));
                    p.children.push(Node::Elem(r));
                    tc.children.push(Node::Elem(p));
                }
            }
        })
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
}

// ---------------------------------------------------------------- fields

/// Span of a field within a paragraph's children.
#[derive(Debug)]
pub(crate) enum FieldSpan {
    /// w:fldSimple child position
    Simple(usize),
    /// (begin run, separate run, end run) child positions
    Complex(usize, Option<usize>, usize),
}

/// Locate all fields of a paragraph, in document order.
pub(crate) fn field_spans(p: &Element) -> Vec<FieldSpan> {
    let mut spans = Vec::new();
    let mut cur: Option<(usize, Option<usize>)> = None;
    for (i, c) in p.children.iter().enumerate() {
        let Node::Elem(e) = c else { continue };
        match e.name.as_str() {
            "w:fldSimple" => spans.push(FieldSpan::Simple(i)),
            "w:r" => {
                if let Some(fc) = e.find("w:fldChar") {
                    match fc.get_attr("w:fldCharType") {
                        Some("begin") => cur = Some((i, None)),
                        Some("separate") => {
                            if let Some((b, _)) = cur {
                                cur = Some((b, Some(i)));
                            }
                        }
                        Some("end") => {
                            if let Some((b, s)) = cur.take() {
                                spans.push(FieldSpan::Complex(b, s, i));
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    spans
}

/// A field code in a paragraph (live proxy). Covers both w:fldSimple and
/// complex (fldChar begin/.../end) fields.
#[pyclass(name = "Field", unsendable)]
pub struct PyField {
    pub tpl: Py<PyDocxTemplate>,
    pub para: usize,
    /// index among the paragraph's fields, in document order
    pub index: usize,
}

impl PyField {
    fn para_proxy(&self, py: Python<'_>) -> PyParagraph {
        PyParagraph {
            tpl: self.tpl.clone_ref(py),
            index: self.para,
        }
    }

    fn with_span<R>(&self, py: Python<'_>, f: impl FnOnce(&Element, &FieldSpan) -> R) -> Option<R> {
        let index = self.index;
        self.para_proxy(py).read(py, move |p| {
            field_spans(p).get(index).map(|s| f(p, s))
        })
        .flatten()
    }

    fn with_span_mut<R>(
        &self,
        py: Python<'_>,
        f: impl FnOnce(&mut Element, &FieldSpan) -> R,
    ) -> PyResult<R> {
        let index = self.index;
        let mut out = None;
        self.para_proxy(py).edit(py, |p| {
            let spans = field_spans(p);
            if let Some(s) = spans.get(index) {
                // field spans were computed on the unmodified tree; f applies
                // them in a single mutation
                out = Some(f(p, s));
            }
        })?;
        out.ok_or_else(|| PyValueError::new_err("field not found"))
    }
}

#[pymethods]
impl PyField {
    /// "simple" (w:fldSimple) or "complex" (fldChar-based).
    #[getter]
    fn kind(&self, py: Python<'_>) -> Option<String> {
        self.with_span(py, |_, s| match s {
            FieldSpan::Simple(_) => "simple".to_string(),
            FieldSpan::Complex(..) => "complex".to_string(),
        })
    }

    /// The field instruction (e.g. `PAGE \* MERGEFORMAT`), trimmed.
    #[getter]
    fn instr(&self, py: Python<'_>) -> Option<String> {
        self.with_span(py, |p, s| match s {
            FieldSpan::Simple(i) => match &p.children[*i] {
                Node::Elem(e) => e
                    .get_attr("w:instr")
                    .map(|s| s.trim().to_string())
                    .unwrap_or_default(),
                _ => String::new(),
            },
            FieldSpan::Complex(b, _, e) => {
                let mut instr = String::new();
                for c in &p.children[*b..=*e] {
                    if let Node::Elem(r) = c {
                        if let Some(it) = r.find("w:instrText") {
                            instr.push_str(&it.text_content());
                        }
                    }
                }
                instr.trim().to_string()
            }
        })
    }

    #[setter]
    fn set_instr(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.with_span_mut(py, |p, s| match s {
            FieldSpan::Simple(i) => {
                if let Node::Elem(e) = &mut p.children[*i] {
                    e.set_attr("w:instr", &format!(" {} ", v.trim()));
                }
            }
            FieldSpan::Complex(b, _, e) => {
                let mut written = false;
                for c in &mut p.children[*b..=*e] {
                    if let Node::Elem(r) = c {
                        if let Some(pos) = r
                            .children
                            .iter()
                            .position(|ch| matches!(ch, Node::Elem(x) if x.name == "w:instrText"))
                        {
                            if !written {
                                let mut it = Element::new("w:instrText");
                                it.set_attr("xml:space", "preserve");
                                it.children
                                    .push(Node::Text(format!(" {} ", v.trim())));
                                r.children[pos] = Node::Elem(it);
                                written = true;
                            } else {
                                // clear extra instrText content
                                if let Node::Elem(x) = &mut r.children[pos] {
                                    x.children.clear();
                                }
                            }
                        }
                    }
                }
            }
        })
    }

    /// The cached (last-rendered) field result text.
    #[getter]
    fn text(&self, py: Python<'_>) -> Option<String> {
        self.with_span(py, |p, s| match s {
            FieldSpan::Simple(i) => match &p.children[*i] {
                Node::Elem(e) => crate::docmodel::element_text(e),
                _ => String::new(),
            },
            FieldSpan::Complex(_, sep, e) => {
                let mut text = String::new();
                if let Some(sep) = sep {
                    for c in &p.children[*sep + 1..*e] {
                        if let Node::Elem(r) = c {
                            text.push_str(&crate::docmodel::element_text(r));
                        }
                    }
                }
                text
            }
        })
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.with_span_mut(py, |p, s| {
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v));
            r.children.push(Node::Elem(t));
            match s {
                FieldSpan::Simple(i) => {
                    if let Node::Elem(e) = &mut p.children[*i] {
                        e.children.retain(|c| {
                            !matches!(c, Node::Elem(x) if x.name == "w:r")
                        });
                        e.children.push(Node::Elem(r));
                    }
                }
                FieldSpan::Complex(_, sep, e) => {
                    if let Some(sep) = sep {
                        // replace the cached-result runs between separate and end
                        let mut tail = p.children.split_off(*e);
                        p.children.truncate(*sep + 1);
                        p.children.push(Node::Elem(r));
                        p.children.append(&mut tail);
                    }
                }
            }
        })
    }
}
