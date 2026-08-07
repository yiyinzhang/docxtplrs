//! Pure-Rust document object model: live, index-based handles over
//! [`TplCore`] (paragraphs/runs/tables/sections/styles/...), shared by the
//! Python bindings (docmodel*.rs are thin wrappers over this module).
//!
//! Handles are `Copy` index proxies; every operation takes `&mut TplCore`
//! (or reads through the cached DOM). Enum values follow python-docx: ints
//! matching the WD_* enumerations; xml vocabulary strings are mapped via the
//! tables below.

use crate::template::{TplCore, DOCUMENT_PART};
use crate::xmldom::{Document, Element, Node};

// ---------------------------------------------------------------- Length

/// A length in EMU (914400 per inch). Pure-Rust counterpart of PyLength.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Length {
    pub emu: i64,
}

impl Length {
    pub fn from_twips(twips: i64) -> Length {
        Length { emu: twips * 635 }
    }
    pub fn twips(self) -> i64 {
        self.emu / 635
    }
}

// ---------------------------------------------------------------- text

/// Concatenated text of an element: w:t text, w:tab -> \t, w:br -> \n.
pub fn element_text(el: &Element) -> String {
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

/// Append text to a run, expanding \t -> w:tab, \n -> w:br, \r -> w:cr
/// (python-docx Run.text semantics).
pub fn run_append_text(r: &mut Element, text: &str) {
    let flush = |r: &mut Element, buf: &mut String| {
        if buf.is_empty() {
            return;
        }
        let mut t = Element::new("w:t");
        t.set_attr("xml:space", "preserve");
        t.children.push(Node::Text(std::mem::take(buf)));
        r.children.push(Node::Elem(t));
    };
    let mut buf = String::new();
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

// ---------------------------------------------------------------- dom access

pub fn read_body<R>(core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
    let dom = core.document_dom().ok()?;
    let body = dom.root.find("w:body")?;
    Some(f(body))
}

/// Mutate the body element of the cached document DOM in place; the change
/// is serialized back into the package on the next flush (render/save/etc).
pub fn mutate_document(core: &mut TplCore, f: impl FnOnce(&mut Element)) -> Result<(), String> {
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

pub fn count_in_body(core: &mut TplCore, name: &str) -> usize {
    read_body(core, |b| {
        b.children
            .iter()
            .filter(|c| matches!(c, Node::Elem(e) if e.name == name))
            .count()
    })
    .unwrap_or(0)
}

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

// ---------------------------------------------------------------- locators

pub fn nth_direct<'a>(el: &'a mut Element, name: &str, n: usize) -> Option<&'a mut Element> {
    el.children
        .iter_mut()
        .filter_map(|c| match c {
            Node::Elem(e) if e.name == name => Some(e),
            _ => None,
        })
        .nth(n)
}

pub fn nth_direct_ref<'a>(el: &'a Element, name: &str, n: usize) -> Option<&'a Element> {
    el.children
        .iter()
        .filter_map(|c| match c {
            Node::Elem(e) if e.name == name => Some(e),
            _ => None,
        })
        .nth(n)
}

pub fn nth_cursor_ref<'a>(
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

pub fn nth_cursor_mut<'a>(
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

/// Read the nth w:p of the body.
pub fn para_read<R>(core: &mut TplCore, index: usize, f: impl FnOnce(&Element) -> R) -> Option<R> {
    let gen = core.doc_gen;
    let mut cur = core.para_cursor;
    let r = read_body(core, |body| {
        nth_cursor_ref(body, "w:p", index, &mut cur, gen).map(|p| f(p))
    })
    .flatten();
    core.para_cursor = cur;
    r
}

/// Mutate the nth w:p of the body.
pub fn para_edit<R>(
    core: &mut TplCore,
    index: usize,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    let mut result = None;
    let gen = core.doc_gen;
    let mut cur = core.para_cursor;
    mutate_document(core, |body| {
        if let Some(p) = nth_cursor_mut(body, "w:p", index, &mut cur, gen) {
            result = Some(f(p));
        }
    })?;
    core.para_cursor = cur;
    result.ok_or_else(|| "paragraph not found".to_string())
}

pub fn run_read<R>(
    core: &mut TplCore,
    para: usize,
    index: usize,
    f: impl FnOnce(&Element) -> R,
) -> Option<R> {
    para_read(core, para, |p| nth_direct_ref(p, "w:r", index).map(|r| f(r))).flatten()
}

pub fn run_edit<R>(
    core: &mut TplCore,
    para: usize,
    index: usize,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    para_edit(core, para, |p| {
        nth_direct(p, "w:r", index)
            .map(|r| f(r))
            .ok_or_else(|| "run not found".to_string())
    })?
}

pub fn tbl_read<R>(core: &mut TplCore, index: usize, f: impl FnOnce(&Element) -> R) -> Option<R> {
    let gen = core.doc_gen;
    let mut cur = core.tbl_cursor;
    let r = read_body(core, |body| {
        nth_cursor_ref(body, "w:tbl", index, &mut cur, gen).map(|t| f(t))
    })
    .flatten();
    core.tbl_cursor = cur;
    r
}

pub fn tbl_edit<R>(
    core: &mut TplCore,
    index: usize,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    let mut result = None;
    let gen = core.doc_gen;
    let mut cur = core.tbl_cursor;
    mutate_document(core, |body| {
        if let Some(t) = nth_cursor_mut(body, "w:tbl", index, &mut cur, gen) {
            result = Some(f(t));
        }
    })?;
    core.tbl_cursor = cur;
    result.ok_or_else(|| "table not found".to_string())
}

pub fn row_read<R>(
    core: &mut TplCore,
    tbl: usize,
    row: usize,
    f: impl FnOnce(&Element) -> R,
) -> Option<R> {
    tbl_read(core, tbl, |t| nth_direct_ref(t, "w:tr", row).map(|r| f(r))).flatten()
}

pub fn row_edit<R>(
    core: &mut TplCore,
    tbl: usize,
    row: usize,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    tbl_edit(core, tbl, |t| {
        nth_direct(t, "w:tr", row)
            .map(|r| f(r))
            .ok_or_else(|| "row not found".to_string())
    })?
}

pub fn cell_read<R>(
    core: &mut TplCore,
    tbl: usize,
    row: usize,
    col: usize,
    f: impl FnOnce(&Element) -> R,
) -> Option<R> {
    row_read(core, tbl, row, |r| nth_direct_ref(r, "w:tc", col).map(|c| f(c))).flatten()
}

pub fn cell_edit<R>(
    core: &mut TplCore,
    tbl: usize,
    row: usize,
    col: usize,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    row_edit(core, tbl, row, |r| {
        nth_direct(r, "w:tc", col)
            .map(|c| f(c))
            .ok_or_else(|| "cell not found".to_string())
    })?
}

// sections

pub fn collect_sectprs_mut<'a>(el: &'a mut Element, out: &mut Vec<&'a mut Element>) {
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

pub fn sect_read<R>(core: &mut TplCore, index: usize, f: impl FnOnce(&Element) -> R) -> Option<R> {
    let dom = core.document_dom().ok()?;
    let mut sects: Vec<&Element> = Vec::new();
    dom.root.iter_descendants("w:sectPr", &mut sects);
    sects.get(index).map(|s| f(s))
}

pub fn sect_edit<R>(
    core: &mut TplCore,
    index: usize,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    let mut result = None;
    mutate_document(core, |body| {
        let mut sects: Vec<&mut Element> = Vec::new();
        collect_sectprs_mut(body, &mut sects);
        if let Some(s) = sects.get_mut(index) {
            result = Some(f(s));
        }
    })?;
    result.ok_or_else(|| "section not found".to_string())
}

// styles / settings parts

pub fn ensure_styles_part(core: &mut TplCore) -> Result<(), String> {
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

pub fn with_styles<R>(core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
    ensure_styles_part(core)?;
    let r = f(&mut core.part_dom("word/styles.xml")?.root);
    core.mark_part_dirty("word/styles.xml");
    Ok(r)
}

pub fn find_style_el<'a>(root: &'a Element, style_id: &str) -> Option<&'a Element> {
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

pub fn find_style_el_mut<'a>(root: &'a mut Element, style_id: &str) -> Option<&'a mut Element> {
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

pub fn style_read<R>(
    core: &mut TplCore,
    style_id: &str,
    f: impl FnOnce(&Element) -> R,
) -> Option<R> {
    let dom = core.part_dom("word/styles.xml").ok()?;
    find_style_el(&dom.root, style_id).map(|e| f(e))
}

pub fn style_edit<R>(
    core: &mut TplCore,
    style_id: &str,
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    let mut result = None;
    with_styles(core, |root| {
        if let Some(st) = find_style_el_mut(root, style_id) {
            result = Some(f(st));
        }
    })?;
    result.ok_or_else(|| "style not found".to_string())
}

pub fn ensure_settings_part(core: &mut TplCore) -> Result<(), String> {
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

pub fn with_settings<R>(core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
    ensure_settings_part(core)?;
    let r = f(&mut core.part_dom("word/settings.xml")?.root);
    core.mark_part_dirty("word/settings.xml");
    Ok(r)
}

// ---------------------------------------------------------------- element helpers

/// get-or-create a direct child element (appended at the end).
pub fn ensure_child<'a>(parent: &'a mut Element, tag: &str) -> &'a mut Element {
    if parent.find(tag).is_none() {
        parent.children.push(Node::Elem(Element::new(tag)));
    }
    parent.find_mut(tag).unwrap()
}

/// w:rPr of a run (must be the first child).
pub fn ensure_rpr(run: &mut Element) -> &mut Element {
    ensure_child_at(run, "w:rPr", 0)
}

/// w:pPr of a paragraph (must be the first child).
pub fn ensure_ppr(p: &mut Element) -> &mut Element {
    ensure_child_at(p, "w:pPr", 0)
}

/// w:trPr of a table row (must be the first child).
pub fn ensure_trpr(tr: &mut Element) -> &mut Element {
    ensure_child_at(tr, "w:trPr", 0)
}

/// w:tblPr of a table (must be the first child).
pub fn ensure_tblpr(tbl: &mut Element) -> &mut Element {
    ensure_child_at(tbl, "w:tblPr", 0)
}

fn ensure_child_at<'a>(parent: &'a mut Element, tag: &str, pos: usize) -> &'a mut Element {
    if parent.find(tag).is_none() {
        parent
            .children
            .insert(pos.min(parent.children.len()), Node::Elem(Element::new(tag)));
    }
    parent.find_mut(tag).unwrap()
}

/// w:tcPr of a table cell (must be the first child).
pub fn tcpr_mut(tc: &mut Element) -> &mut Element {
    ensure_child_at(tc, "w:tcPr", 0)
}

pub fn remove_attr(el: &mut Element, name: &str) {
    el.attrs.retain(|(k, _)| k != name);
}

pub fn remove_child(parent: &mut Element, tag: &str) {
    parent
        .children
        .retain(|c| !matches!(c, Node::Elem(e) if e.name == tag));
}

/// on/off child element (missing == off); set false removes the element.
pub fn set_flag(r: &mut Element, tag: &str, on: bool) {
    let rpr = ensure_rpr(r);
    let exists = rpr.find(tag).is_some();
    if on && !exists {
        rpr.children.push(Node::Elem(Element::new(tag)));
    } else if !on && exists {
        remove_child(rpr, tag);
    }
}

/// on/off child element of any container (missing == off).
pub fn flag_on(container: &mut Element, tag: &str, on: bool) {
    let exists = container.find(tag).is_some();
    if on && !exists {
        container.children.push(Node::Elem(Element::new(tag)));
    } else if !on && exists {
        remove_child(container, tag);
    }
}

/// set attr on a run's rPr child element (get-or-create).
pub fn set_val_tag(r: &mut Element, tag: &str, attr: &str, val: &str) {
    let rpr = ensure_rpr(r);
    attr_set(rpr, tag, attr, Some(val));
}

/// read an on/off flag from a run's rPr (missing -> None).
pub fn read_flag(el: &Element, tag: &str) -> Option<bool> {
    let rpr = el.find("w:rPr")?;
    tri_get(rpr, tag)
}

/// tri-state bool read: missing element -> None; missing w:val -> true.
pub fn tri_get(container: &Element, tag: &str) -> Option<bool> {
    container.find(tag).map(|e| {
        !matches!(
            e.get_attr("w:val"),
            Some("0") | Some("false") | Some("off") | Some("none")
        )
    })
}

/// tri-state bool write: None -> remove, Some(true) -> bare, Some(false) ->
/// `w:val="0"` (python-docx semantics).
pub fn tri_set(container: &mut Element, tag: &str, v: Option<bool>) {
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

pub fn attr_get(container: &Element, tag: &str, attr: &str) -> Option<String> {
    container
        .find(tag)
        .and_then(|e| e.get_attr(attr))
        .map(|s| s.to_string())
}

/// set an attribute on a child element; None removes the whole element.
pub fn attr_set(container: &mut Element, tag: &str, attr: &str, val: Option<&str>) {
    match val {
        None => remove_child(container, tag),
        Some(v) => ensure_child(container, tag).set_attr(attr, v),
    }
}

// twips helpers (sections / spacing)

pub fn get_twips(sp: &Element, tag: &str, attr: &str) -> Option<i64> {
    sp.find(tag)
        .and_then(|e| e.get_attr(attr))
        .and_then(|v| v.parse::<i64>().ok())
}

pub fn set_twips(sp: &mut Element, tag: &str, attr: &str, v: Option<i64>, defaults: &[(&str, &str)]) {
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

pub fn len_get(container: &Element, tag: &str, attr: &str) -> Option<Length> {
    attr_get(container, tag, attr)
        .and_then(|s| s.parse::<i64>().ok())
        .map(Length::from_twips)
}

/// v: EMU. None removes the attribute (element kept if it has other attrs).
pub fn len_set(container: &mut Element, tag: &str, attr: &str, v: Option<i64>) {
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

// ---------------------------------------------------------------- enum tables

pub fn xml_of(table: &'static [(i64, &'static str)], v: i64) -> Option<&'static str> {
    table.iter().find(|(i, _)| *i == v).map(|(_, s)| *s)
}

pub fn int_of(table: &[(i64, &str)], s: &str) -> Option<i64> {
    table.iter().find(|(_, x)| *x == s).map(|(i, _)| *i)
}

pub fn enum_get(container: &Element, tag: &str, attr: &str, table: &[(i64, &str)]) -> Option<i64> {
    attr_get(container, tag, attr).and_then(|s| int_of(table, &s))
}

pub fn enum_set(
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

/// WD_ALIGN_PARAGRAPH
pub const ALIGN: &[(i64, &str)] = &[
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

/// WD_UNDERLINE
pub const UNDERLINE: &[(i64, &str)] = &[
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

/// WD_COLOR_INDEX (highlight)
pub const HIGHLIGHT: &[(i64, &str)] = &[
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

/// WD_TAB_ALIGNMENT
pub const TAB_ALIGN: &[(i64, &str)] = &[
    (0, "left"),
    (1, "center"),
    (2, "right"),
    (3, "decimal"),
    (4, "bar"),
    (6, "list"),
    (101, "clear"),
];

/// WD_TAB_LEADER
pub const TAB_LEADER: &[(i64, &str)] = &[
    (0, "none"),
    (1, "dot"),
    (2, "hyphen"),
    (3, "underscore"),
    (4, "heavy"),
    (5, "middleDot"),
];

/// WD_CELL_VERTICAL_ALIGNMENT
pub const VALIGN: &[(i64, &str)] = &[(0, "top"), (1, "center"), (3, "bottom"), (101, "both")];

/// WD_ROW_HEIGHT_RULE
pub const HRULE: &[(i64, &str)] = &[(0, "auto"), (1, "atLeast"), (2, "exact")];

/// WD_SECTION_START
pub const SECTION_START: &[(i64, &str)] = &[
    (0, "continuous"),
    (1, "nextColumn"),
    (2, "nextPage"),
    (3, "evenPage"),
    (4, "oddPage"),
];

/// WD_TABLE_ALIGNMENT
pub const TBL_ALIGN: &[(i64, &str)] = &[(0, "left"), (1, "center"), (2, "right")];

// ---------------------------------------------------------------- targets

/// What a Font/ColorFormat is attached to.
#[derive(Debug, Clone)]
pub enum FontTarget {
    Run { para: usize, index: usize },
    Style { style_id: String },
}

/// What a ParagraphFormat/TabStops is attached to.
#[derive(Debug, Clone)]
pub enum PfTarget {
    Para { index: usize },
    Style { style_id: String },
}

/// Mixed block-level content (document/cell/section iteration).
#[derive(Debug, Clone, Copy)]
pub enum BlockItem {
    Paragraph(usize),
    Table(usize),
}

/// on/off child element of a style (missing == off).
pub fn style_flag(st: &mut Element, tag: &str, on: bool) {
    flag_on(st, tag, on);
}

fn tc_positions(row: &Element) -> Vec<usize> {
    row.children
        .iter()
        .enumerate()
        .filter(|(_, c)| matches!(c, Node::Elem(e) if e.name == "w:tc"))
        .map(|(i, _)| i)
        .collect()
}

fn is_non_empty_block(ch: &Node) -> bool {
    match ch {
        Node::Elem(e) if e.name == "w:tcPr" => false,
        Node::Elem(e) => e.name != "w:p" || !element_text(e).is_empty(),
        Node::Text(_) => false,
    }
}

/// Merge the rectangular cell region (r1,c1)-(r2,c2) of a table
/// (python-docx cell.merge semantics: gridSpan/vMerge, content collection).
pub fn merge_region(t: &mut Element, r1: usize, r2: usize, c1: usize, c2: usize) {
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

// ---------------------------------------------------------------- Font handle

/// Line spacing value (python-docx ParagraphFormat.line_spacing semantics):
/// a float multiple when the rule is auto, otherwise an exact/at-least Length.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacing {
    /// float multiple (w:lineRule="auto"); 2.0 = double
    Multiple(f64),
    /// exact / at-least height (w:line in twips)
    Exact(Length),
}

/// Pure-Rust python-docx Font: full run/style character properties.
pub struct Font {
    pub target: FontTarget,
}

impl Font {
    pub fn edit_rpr<R>(
        &self,
        core: &mut TplCore,
        f: impl FnOnce(&mut Element) -> R,
    ) -> Result<R, String> {
        match &self.target {
            FontTarget::Run { para, index } => run_edit(core, *para, *index, |r| f(ensure_rpr(r))),
            FontTarget::Style { style_id } => style_edit(core, style_id, |el| {
                if el.find("w:rPr").is_none() {
                    el.children.push(Node::Elem(Element::new("w:rPr")));
                }
                f(el.find_mut("w:rPr").unwrap())
            }),
        }
    }

    pub fn read_rpr<R>(&self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        match &self.target {
            FontTarget::Run { para, index } => {
                run_read(core, *para, *index, |r| r.find("w:rPr").map(|e| f(e))).flatten()
            }
            FontTarget::Style { style_id } => {
                style_read(core, style_id, |el| el.find("w:rPr").map(|e| f(e))).flatten()
            }
        }
    }
}

/// Generate tri-state bool property pairs on Font.
macro_rules! font_tri {
    ($($name:ident / $setter:ident => $tag:literal),* $(,)?) => {
        impl Font {
            $(
            pub fn $name(&self, core: &mut TplCore) -> Option<bool> {
                self.read_rpr(core, |rpr| tri_get(rpr, $tag)).flatten()
            }
            pub fn $setter(&self, core: &mut TplCore, v: Option<bool>) -> Result<(), String> {
                self.edit_rpr(core, |rpr| tri_set(rpr, $tag, v))
            }
            )*
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
}

impl Font {
    pub fn subscript(&self, core: &mut TplCore) -> Option<bool> {
        self.read_rpr(core, |rpr| {
            rpr.find("w:vertAlign")
                .map(|e| e.get_attr("w:val") == Some("subscript"))
        })
        .flatten()
    }
    pub fn set_subscript(&self, core: &mut TplCore, v: Option<bool>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| match v {
            Some(true) => attr_set(rpr, "w:vertAlign", "w:val", Some("subscript")),
            Some(false) => {
                if attr_get(rpr, "w:vertAlign", "w:val").as_deref() == Some("subscript") {
                    remove_child(rpr, "w:vertAlign");
                }
            }
            None => remove_child(rpr, "w:vertAlign"),
        })
    }
    pub fn superscript(&self, core: &mut TplCore) -> Option<bool> {
        self.read_rpr(core, |rpr| {
            rpr.find("w:vertAlign")
                .map(|e| e.get_attr("w:val") == Some("superscript"))
        })
        .flatten()
    }
    pub fn set_superscript(&self, core: &mut TplCore, v: Option<bool>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| match v {
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
    pub fn size(&self, core: &mut TplCore) -> Option<Length> {
        self.read_rpr(core, |rpr| {
            attr_get(rpr, "w:sz", "w:val")
                .and_then(|s| s.parse::<i64>().ok())
                .map(|hp| Length { emu: hp * 12700 / 2 })
        })
        .flatten()
    }
    /// emu: EMU (None clears).
    pub fn set_size(&self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| match emu {
            Some(e) => {
                let hp = ((e * 2) / 12700).to_string();
                attr_set(rpr, "w:sz", "w:val", Some(&hp));
            }
            None => remove_child(rpr, "w:sz"),
        })
    }

    pub fn name(&self, core: &mut TplCore) -> Option<String> {
        self.read_rpr(core, |rpr| attr_get(rpr, "w:rFonts", "w:ascii"))
            .flatten()
    }
    pub fn set_name(&self, core: &mut TplCore, v: Option<String>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| match v {
            Some(name) => {
                let rf = ensure_child(rpr, "w:rFonts");
                rf.set_attr("w:ascii", &name);
                rf.set_attr("w:hAnsi", &name);
            }
            None => remove_child(rpr, "w:rFonts"),
        })
    }

    /// Highlight color as a WD_COLOR_INDEX int.
    pub fn highlight_color(&self, core: &mut TplCore) -> Option<i64> {
        self.read_rpr(core, |rpr| enum_get(rpr, "w:highlight", "w:val", HIGHLIGHT))
            .flatten()
    }
    pub fn set_highlight_color(&self, core: &mut TplCore, v: Option<i64>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| enum_set(rpr, "w:highlight", "w:val", HIGHLIGHT, v))
    }

    /// Underline as a WD_UNDERLINE int.
    pub fn underline(&self, core: &mut TplCore) -> Option<i64> {
        self.read_rpr(core, |rpr| enum_get(rpr, "w:u", "w:val", UNDERLINE))
            .flatten()
    }
    pub fn set_underline(&self, core: &mut TplCore, v: Option<i64>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| enum_set(rpr, "w:u", "w:val", UNDERLINE, v))
    }

    /// RGB as a 6-digit hex string (e.g. "FF0000"), None when not set.
    pub fn color_rgb(&self, core: &mut TplCore) -> Option<String> {
        self.read_rpr(core, |rpr| attr_get(rpr, "w:color", "w:val"))
            .flatten()
    }
    /// hex: 6-digit uppercase hex string; None clears.
    pub fn set_color_rgb(&self, core: &mut TplCore, hex: Option<String>) -> Result<(), String> {
        self.edit_rpr(core, |rpr| attr_set(rpr, "w:color", "w:val", hex.as_deref()))
    }
}

// ---------------------------------------------------------------- ParagraphFormat handle

/// Pure-Rust python-docx ParagraphFormat (w:pPr properties).
pub struct ParagraphFormat {
    pub target: PfTarget,
}

impl ParagraphFormat {
    pub fn edit_ppr<R>(
        &self,
        core: &mut TplCore,
        f: impl FnOnce(&mut Element) -> R,
    ) -> Result<R, String> {
        match &self.target {
            PfTarget::Para { index } => para_edit(core, *index, |el| f(ensure_ppr(el))),
            PfTarget::Style { style_id } => style_edit(core, style_id, |el| {
                if el.find("w:pPr").is_none() {
                    el.children.push(Node::Elem(Element::new("w:pPr")));
                }
                f(el.find_mut("w:pPr").unwrap())
            }),
        }
    }

    pub fn read_ppr<R>(&self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        match &self.target {
            PfTarget::Para { index } => {
                para_read(core, *index, |el| el.find("w:pPr").map(|e| f(e))).flatten()
            }
            PfTarget::Style { style_id } => {
                style_read(core, style_id, |el| el.find("w:pPr").map(|e| f(e))).flatten()
            }
        }
    }
}

/// tri-state bool properties on ParagraphFormat.
macro_rules! pf_tri {
    ($($name:ident / $setter:ident => $tag:literal),* $(,)?) => {
        impl ParagraphFormat {
            $(
            pub fn $name(&self, core: &mut TplCore) -> Option<bool> {
                self.read_ppr(core, |ppr| tri_get(ppr, $tag)).flatten()
            }
            pub fn $setter(&self, core: &mut TplCore, v: Option<bool>) -> Result<(), String> {
                self.edit_ppr(core, |ppr| tri_set(ppr, $tag, v))
            }
            )*
        }
    };
}

pf_tri! {
    keep_together / set_keep_together => "w:keepLines",
    keep_with_next / set_keep_with_next => "w:keepNext",
    page_break_before / set_page_break_before => "w:pageBreakBefore",
}

impl ParagraphFormat {
    /// Alignment as a WD_ALIGN_PARAGRAPH int.
    pub fn alignment(&self, core: &mut TplCore) -> Option<i64> {
        self.read_ppr(core, |ppr| enum_get(ppr, "w:jc", "w:val", ALIGN))
            .flatten()
    }
    pub fn set_alignment(&self, core: &mut TplCore, v: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| enum_set(ppr, "w:jc", "w:val", ALIGN, v))
    }

    pub fn left_indent(&self, core: &mut TplCore) -> Option<Length> {
        self.read_ppr(core, |ppr| len_get(ppr, "w:ind", "w:left"))
            .flatten()
    }
    /// emu: EMU (None removes the attribute).
    pub fn set_left_indent(&self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| len_set(ppr, "w:ind", "w:left", emu))
    }
    pub fn right_indent(&self, core: &mut TplCore) -> Option<Length> {
        self.read_ppr(core, |ppr| len_get(ppr, "w:ind", "w:right"))
            .flatten()
    }
    pub fn set_right_indent(&self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| len_set(ppr, "w:ind", "w:right", emu))
    }

    /// First-line indent; negative values become a hanging indent.
    pub fn first_line_indent(&self, core: &mut TplCore) -> Option<Length> {
        self.read_ppr(core, |ppr| {
            if let Some(l) = len_get(ppr, "w:ind", "w:firstLine") {
                return Some(l);
            }
            len_get(ppr, "w:ind", "w:hanging").map(|l| Length { emu: -l.emu })
        })
        .flatten()
    }
    pub fn set_first_line_indent(&self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| {
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

    pub fn space_before(&self, core: &mut TplCore) -> Option<Length> {
        self.read_ppr(core, |ppr| len_get(ppr, "w:spacing", "w:before"))
            .flatten()
    }
    pub fn set_space_before(&self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| len_set(ppr, "w:spacing", "w:before", emu))
    }
    pub fn space_after(&self, core: &mut TplCore) -> Option<Length> {
        self.read_ppr(core, |ppr| len_get(ppr, "w:spacing", "w:after"))
            .flatten()
    }
    pub fn set_space_after(&self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| len_set(ppr, "w:spacing", "w:after", emu))
    }

    /// Line spacing: a float multiple when the rule is auto, otherwise a
    /// Length (exact / at-least).
    pub fn line_spacing(&self, core: &mut TplCore) -> Option<LineSpacing> {
        let (line, rule) = self.read_ppr(core, |ppr| {
            let sp = ppr.find("w:spacing")?;
            let line = sp.get_attr("w:line")?.parse::<i64>().ok()?;
            let rule = sp.get_attr("w:lineRule").unwrap_or("auto").to_string();
            Some((line, rule))
        })??;
        Some(if rule == "auto" {
            LineSpacing::Multiple(line as f64 / 240.0)
        } else {
            LineSpacing::Exact(Length::from_twips(line))
        })
    }
    pub fn set_line_spacing(&self, core: &mut TplCore, v: Option<LineSpacing>) -> Result<(), String> {
        match v {
            None => self.edit_ppr(core, |ppr| {
                if let Some(sp) = ppr.find_mut("w:spacing") {
                    remove_attr(sp, "w:line");
                    remove_attr(sp, "w:lineRule");
                }
            }),
            Some(LineSpacing::Multiple(f)) => {
                let line = (f * 240.0).round() as i64;
                self.edit_ppr(core, |ppr| {
                    let sp = ensure_child(ppr, "w:spacing");
                    sp.set_attr("w:line", &line.to_string());
                    sp.set_attr("w:lineRule", "auto");
                })
            }
            Some(LineSpacing::Exact(l)) => self.edit_ppr(core, |ppr| {
                let sp = ensure_child(ppr, "w:spacing");
                sp.set_attr("w:line", &l.twips().to_string());
                // keep an existing atLeast rule, else exact
                if sp.get_attr("w:lineRule") != Some("atLeast") {
                    sp.set_attr("w:lineRule", "exact");
                }
            }),
        }
    }

    /// Line spacing rule as a WD_LINE_SPACING int.
    pub fn line_spacing_rule(&self, core: &mut TplCore) -> Option<i64> {
        self.read_ppr(core, |ppr| {
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
    pub fn set_line_spacing_rule(&self, core: &mut TplCore, v: Option<i64>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| {
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
    pub fn widow_control(&self, core: &mut TplCore) -> Option<bool> {
        Some(
            self.read_ppr(core, |ppr| tri_get(ppr, "w:widowControl"))
                .flatten()
                .unwrap_or(true),
        )
    }
    pub fn set_widow_control(&self, core: &mut TplCore, v: Option<bool>) -> Result<(), String> {
        self.edit_ppr(core, |ppr| tri_set(ppr, "w:widowControl", v))
    }
}

// ---------------------------------------------------------------- TabStops handle

/// Pure-Rust python-docx TabStops (w:pPr/w:tabs).
pub struct TabStops {
    pub target: PfTarget,
}

impl TabStops {
    fn pf(&self) -> ParagraphFormat {
        ParagraphFormat {
            target: self.target.clone(),
        }
    }

    pub fn len(&self, core: &mut TplCore) -> usize {
        self.pf()
            .read_ppr(core, |ppr| {
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

    pub fn is_empty(&self, core: &mut TplCore) -> bool {
        self.len(core) == 0
    }

    /// Add a tab stop; returns the index of the new stop. position: EMU;
    /// alignment/leader: WD_TAB_ALIGNMENT / WD_TAB_LEADER ints (must be valid).
    pub fn add_tab_stop(
        &self,
        core: &mut TplCore,
        pos_emu: i64,
        align: i64,
        leader: i64,
    ) -> Result<usize, String> {
        self.pf().edit_ppr(core, |ppr| {
            let tabs = ensure_child(ppr, "w:tabs");
            let n = tabs
                .children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tab"))
                .count();
            let mut tab = Element::new("w:tab");
            tab.set_attr("w:val", xml_of(TAB_ALIGN, align).unwrap());
            tab.set_attr("w:leader", xml_of(TAB_LEADER, leader).unwrap());
            tab.set_attr("w:pos", &(pos_emu / 635).to_string());
            tabs.children.push(Node::Elem(tab));
            n
        })
    }

    pub fn clear_all(&self, core: &mut TplCore) -> Result<(), String> {
        self.pf().edit_ppr(core, |ppr| remove_child(ppr, "w:tabs"))
    }
}

/// A single tab stop (live proxy, read-only).
pub struct TabStop {
    pub target: PfTarget,
    pub index: usize,
}

impl TabStop {
    fn read_tab<R>(&self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let pf = ParagraphFormat {
            target: self.target.clone(),
        };
        let index = self.index;
        pf.read_ppr(core, |ppr| {
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

    pub fn position(&self, core: &mut TplCore) -> Option<Length> {
        self.read_tab(core, |t| {
            t.get_attr("w:pos")
                .and_then(|s| s.parse::<i64>().ok())
                .map(Length::from_twips)
        })
        .flatten()
    }
    pub fn alignment(&self, core: &mut TplCore) -> Option<i64> {
        self.read_tab(core, |t| t.get_attr("w:val").and_then(|s| int_of(TAB_ALIGN, s)))
            .flatten()
    }
    pub fn leader(&self, core: &mut TplCore) -> Option<i64> {
        self.read_tab(core, |t| t.get_attr("w:leader").and_then(|s| int_of(TAB_LEADER, s)))
            .flatten()
    }
}

// ---------------------------------------------------------------- Paragraph handle

/// Item of a paragraph's inner content (python-docx iter_inner_content):
/// a run or a hyperlink, by index among its siblings of the same kind.
#[derive(Debug, Clone, Copy)]
pub enum ParaItem {
    Run(usize),
    Hyperlink(usize),
}

/// Item of a run's inner content (python-docx iter_inner_content).
#[derive(Debug, Clone)]
pub enum RunItem {
    /// contiguous text-ish range (w:t/w:tab/w:br/w:cr/w:noBreakHyphen)
    Text(String),
    /// a w:drawing, addressed by its full element path from the document
    /// root ([body, para, run, child] element-child indices)
    Drawing(Vec<usize>),
    RenderedPageBreak,
}

/// Pure-Rust paragraph handle (body-level w:p by index).
#[derive(Debug, Clone, Copy)]
pub struct Paragraph {
    pub index: usize,
}

impl Paragraph {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        para_read(core, self.index, f)
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        para_edit(core, self.index, f)
    }

    pub fn text(self, core: &mut TplCore) -> String {
        self.read(core, element_text).unwrap_or_default()
    }

    /// python-docx: clear content, add a single run.
    pub fn set_text(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |p| {
            p.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:r" || e.name == "w:hyperlink"));
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v.to_string()));
            r.children.push(Node::Elem(t));
            p.children.push(Node::Elem(r));
        })
    }

    pub fn style(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |p| {
            p.find("w:pPr")
                .and_then(|ppr| ppr.find("w:pStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }

    /// Accepts a style name or id (resolved via the styles part).
    pub fn set_style(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let sid = crate::subdocbuilder::resolve_style_id(core, v);
        self.edit(core, |p| {
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

    pub fn run_count(self, core: &mut TplCore) -> usize {
        self.read(core, |p| {
            p.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:r"))
                .count()
        })
        .unwrap_or(0)
    }

    pub fn add_run(self, core: &mut TplCore, text: &str) -> Result<Run, String> {
        let index = self.index;
        let n = self.edit(core, |p| {
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
        Ok(Run { para: index, index: n })
    }

    pub fn paragraph_format(self) -> ParagraphFormat {
        ParagraphFormat {
            target: PfTarget::Para { index: self.index },
        }
    }

    /// Alignment shortcut (WD_ALIGN_PARAGRAPH int).
    pub fn alignment(self, core: &mut TplCore) -> Option<i64> {
        self.paragraph_format().alignment(core)
    }
    pub fn set_alignment(self, core: &mut TplCore, v: Option<i64>) -> Result<(), String> {
        self.paragraph_format().set_alignment(core, v)
    }

    /// Remove all content, keeping the paragraph properties (python-docx clear).
    pub fn clear(self, core: &mut TplCore) -> Result<(), String> {
        self.edit(core, |p| {
            p.children
                .retain(|c| matches!(c, Node::Elem(e) if e.name == "w:pPr"));
        })
    }

    /// Insert a new paragraph before this one (python-docx
    /// insert_paragraph_before); returns the new paragraph (which takes over
    /// this paragraph's index). style: style name or id.
    pub fn insert_paragraph_before(
        self,
        core: &mut TplCore,
        text: Option<&str>,
        style: Option<&str>,
    ) -> Result<Paragraph, String> {
        let sid = style.map(|s| crate::subdocbuilder::resolve_style_id(core, s));
        let index = self.index;
        let mut found = false;
        mutate_document(core, |body| {
            let mut seen = 0usize;
            let mut pos = None;
            for (i, c) in body.children.iter().enumerate() {
                if matches!(c, Node::Elem(e) if e.name == "w:p") {
                    if seen == index {
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
        })?;
        if !found {
            return Err("paragraph not found".to_string());
        }
        Ok(Paragraph { index })
    }

    pub fn hyperlink_count(self, core: &mut TplCore) -> usize {
        self.read(core, |p| {
            p.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:hyperlink"))
                .count()
        })
        .unwrap_or(0)
    }

    /// True when a rendered page break occurs in this paragraph.
    pub fn contains_page_break(self, core: &mut TplCore) -> bool {
        self.read(core, |p| {
            let mut out = Vec::new();
            p.iter_descendants("w:lastRenderedPageBreak", &mut out);
            !out.is_empty()
        })
        .unwrap_or(false)
    }

    pub fn rendered_page_break_count(self, core: &mut TplCore) -> usize {
        self.read(core, |p| {
            let mut out = Vec::new();
            p.iter_descendants("w:lastRenderedPageBreak", &mut out);
            out.len()
        })
        .unwrap_or(0)
    }

    /// Runs and hyperlinks of this paragraph in document order.
    pub fn iter_inner_content(self, core: &mut TplCore) -> Vec<ParaItem> {
        self.read(core, |p| {
            let mut ri = 0usize;
            let mut hi = 0usize;
            let mut out = Vec::new();
            for c in &p.children {
                match c {
                    Node::Elem(e) if e.name == "w:r" => {
                        out.push(ParaItem::Run(ri));
                        ri += 1;
                    }
                    Node::Elem(e) if e.name == "w:hyperlink" => {
                        out.push(ParaItem::Hyperlink(hi));
                        hi += 1;
                    }
                    _ => {}
                }
            }
            out
        })
        .unwrap_or_default()
    }

    pub fn field_count(self, core: &mut TplCore) -> usize {
        self.read(core, |p| field_spans(p).len()).unwrap_or(0)
    }

    /// Append a complex field (begin/instrText/separate/cached/end) to this
    /// paragraph; returns a handle to the new field.
    pub fn add_field(self, core: &mut TplCore, instr: &str, cached: &str) -> Result<Field, String> {
        let instr = instr.trim().to_string();
        let cached = cached.to_string();
        let para = self.index;
        let index = self.edit(core, |p| {
            let index = field_spans(p).len();
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
        Ok(Field { para, index })
    }
}

// ---------------------------------------------------------------- Run handle

/// Pure-Rust run handle (w:r by paragraph/run index).
#[derive(Debug, Clone, Copy)]
pub struct Run {
    pub para: usize,
    pub index: usize,
}

impl Run {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        run_read(core, self.para, self.index, f)
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        run_edit(core, self.para, self.index, f)
    }

    pub fn text(self, core: &mut TplCore) -> String {
        self.read(core, element_text).unwrap_or_default()
    }

    pub fn set_text(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |r| {
            r.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:t" || e.name == "w:tab" || e.name == "w:br" || e.name == "w:cr"));
            run_append_text(r, v);
        })
    }

    /// Two-state on/off flag (legacy run facade): missing -> None; set false
    /// removes the element.
    pub fn flag(self, core: &mut TplCore, tag: &str) -> Option<bool> {
        self.read(core, |r| read_flag(r, tag)).flatten()
    }
    pub fn set_flag(self, core: &mut TplCore, tag: &str, v: bool) -> Result<(), String> {
        self.edit(core, |r| set_flag(r, tag, v))
    }

    pub fn bold(self, core: &mut TplCore) -> Option<bool> {
        self.flag(core, "w:b")
    }
    pub fn set_bold(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:b", v)
    }
    pub fn italic(self, core: &mut TplCore) -> Option<bool> {
        self.flag(core, "w:i")
    }
    pub fn set_italic(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:i", v)
    }
    pub fn strike(self, core: &mut TplCore) -> Option<bool> {
        self.flag(core, "w:strike")
    }
    pub fn set_strike(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:strike", v)
    }

    pub fn underline(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:u"))
                .and_then(|u| u.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    /// Some(v) writes w:u@w:val; None removes the element.
    pub fn set_underline(self, core: &mut TplCore, v: Option<&str>) -> Result<(), String> {
        match v {
            Some(s) => self.edit(core, |r| set_val_tag(r, "w:u", "w:val", s)),
            None => self.edit(core, |r| set_flag(r, "w:u", false)),
        }
    }

    pub fn style(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:rStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    /// Accepts a style name or id (resolved via the styles part).
    pub fn set_style(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let sid = crate::subdocbuilder::resolve_style_id(core, v);
        self.edit(core, |r| set_val_tag(r, "w:rStyle", "w:val", &sid))
    }

    pub fn font_name(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:rFonts"))
                .and_then(|e| e.get_attr("w:ascii").map(|s| s.to_string()))
        })
        .flatten()
    }
    pub fn set_font_name(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |r| {
            let rpr = ensure_rpr(r);
            if let Some(el) = rpr.find_mut("w:rFonts") {
                el.set_attr("w:ascii", v);
                el.set_attr("w:hAnsi", v);
                el.set_attr("w:cs", v);
            } else {
                let mut el = Element::new("w:rFonts");
                el.set_attr("w:ascii", v);
                el.set_attr("w:hAnsi", v);
                el.set_attr("w:cs", v);
                rpr.children.push(Node::Elem(el));
            }
        })
    }

    /// Font size in half-points (w:sz; set also writes w:szCs).
    pub fn size(self, core: &mut TplCore) -> Option<u32> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:sz"))
                .and_then(|e| e.get_attr("w:val").and_then(|v| v.parse().ok()))
        })
        .flatten()
    }
    pub fn set_size(self, core: &mut TplCore, v: u32) -> Result<(), String> {
        let s = v.to_string();
        self.edit(core, |r| {
            set_val_tag(r, "w:sz", "w:val", &s);
            set_val_tag(r, "w:szCs", "w:val", &s);
        })
    }

    pub fn color(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:color"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    pub fn set_color(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let c = v.strip_prefix('#').unwrap_or(v).to_string();
        self.edit(core, |r| set_val_tag(r, "w:color", "w:val", &c))
    }

    /// Highlight via w:shd@w:fill (legacy run facade).
    pub fn highlight(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:shd"))
                .and_then(|e| e.get_attr("w:fill").map(|s| s.to_string()))
        })
        .flatten()
    }
    pub fn set_highlight(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let c = v.strip_prefix('#').unwrap_or(v).to_string();
        self.edit(core, |r| set_val_tag(r, "w:shd", "w:fill", &c))
    }

    pub fn subscript(self, core: &mut TplCore) -> Option<bool> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:vertAlign"))
                .map(|e| e.get_attr("w:val") == Some("subscript"))
        })
        .flatten()
    }
    pub fn set_subscript(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        if v {
            self.edit(core, |r| set_val_tag(r, "w:vertAlign", "w:val", "subscript"))
        } else {
            self.edit(core, |r| set_flag(r, "w:vertAlign", false))
        }
    }
    pub fn superscript(self, core: &mut TplCore) -> Option<bool> {
        self.read(core, |r| {
            r.find("w:rPr")
                .and_then(|rpr| rpr.find("w:vertAlign"))
                .map(|e| e.get_attr("w:val") == Some("superscript"))
        })
        .flatten()
    }
    pub fn set_superscript(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        if v {
            self.edit(core, |r| set_val_tag(r, "w:vertAlign", "w:val", "superscript"))
        } else {
            self.edit(core, |r| set_flag(r, "w:vertAlign", false))
        }
    }

    /// Full python-docx Font facade over this run's w:rPr.
    pub fn font(self) -> Font {
        Font {
            target: FontTarget::Run {
                para: self.para,
                index: self.index,
            },
        }
    }

    /// Add a break; break_type is a WD_BREAK int (6=line, 7=page, 8=column,
    /// 9/10/11=textWrapping clear left/right/all).
    pub fn add_break(self, core: &mut TplCore, break_type: i64) -> Result<(), String> {
        self.edit(core, |r| {
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
    pub fn add_tab(self, core: &mut TplCore) -> Result<(), String> {
        self.edit(core, |r| r.children.push(Node::Elem(Element::new("w:tab"))))
    }

    /// Append text (w:t), preserving leading/trailing whitespace.
    pub fn add_text(self, core: &mut TplCore, text: &str) -> Result<(), String> {
        self.edit(core, |r| run_append_text(r, text))
    }

    /// Remove all content, keeping run properties (python-docx clear).
    pub fn clear(self, core: &mut TplCore) -> Result<(), String> {
        self.edit(core, |r| {
            r.children
                .retain(|c| matches!(c, Node::Elem(e) if e.name == "w:rPr"));
        })
    }

    /// True when a rendered page break (w:lastRenderedPageBreak) occurs in
    /// this run (hard breaks are not counted, python-docx semantics).
    pub fn contains_page_break(self, core: &mut TplCore) -> bool {
        self.read(core, |r| {
            let mut out = Vec::new();
            r.iter_descendants("w:lastRenderedPageBreak", &mut out);
            !out.is_empty()
        })
        .unwrap_or(false)
    }

    /// Content items of this run in order (python-docx iter_inner_content).
    pub fn iter_inner_content(self, core: &mut TplCore) -> Vec<RunItem> {
        let path_and_items: Option<(Vec<usize>, Vec<RunItem>)> = read_body(core, |body| {
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
            let mut items: Vec<RunItem> = Vec::new();
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
                                items.push(RunItem::Text(std::mem::take(&mut cur)));
                            }
                            items.push(RunItem::Drawing(vec![i]));
                        }
                        "w:lastRenderedPageBreak" => {
                            if !cur.is_empty() {
                                items.push(RunItem::Text(std::mem::take(&mut cur)));
                            }
                            items.push(RunItem::RenderedPageBreak);
                        }
                        _ => {}
                    }
                }
            }
            if !cur.is_empty() {
                items.push(RunItem::Text(cur));
            }
            Some((vec![p_pos, r_pos], items))
        })
        .flatten();
        let Some((tail, items)) = path_and_items else {
            return Vec::new();
        };
        // element index of w:body within w:document (almost always 0)
        let body_idx = core
            .document_dom()
            .ok()
            .and_then(|dom| {
                dom.root
                    .children
                    .iter()
                    .filter_map(|c| match c {
                        Node::Elem(e) => Some(e),
                        _ => None,
                    })
                    .position(|e| e.name == "w:body")
            })
            .unwrap_or(0);
        let mut base = vec![body_idx];
        base.extend(tail);
        items
            .into_iter()
            .map(|it| match it {
                RunItem::Drawing(mut p) => {
                    let mut full = base.clone();
                    full.append(&mut p);
                    RunItem::Drawing(full)
                }
                other => other,
            })
            .collect()
    }

    /// Append a picture to this run (python-docx run.add_picture).
    pub fn add_picture(
        self,
        core: &mut TplCore,
        blob: &[u8],
        filename: Option<&str>,
        width: Option<i64>,
        height: Option<i64>,
    ) -> Result<(), String> {
        core.init_docx(false)?;
        let drawing = crate::inline_image::drawing_xml(
            core,
            DOCUMENT_PART,
            blob,
            filename,
            width,
            height,
            None,
            None,
            None,
        )?;
        self.edit(core, |r| {
            if let Ok(frag) = crate::subdoc::parse_body_fragment(&drawing) {
                r.children.extend(frag.root.children);
            }
        })
    }

    /// Mark the range from this run to `last` as belonging to the comment
    /// `comment_id` (python-docx run.mark_comment_range).
    pub fn mark_comment_range(self, core: &mut TplCore, last: Run, comment_id: i64) -> Result<(), String> {
        let id = comment_id.to_string();
        mutate_document(core, |body| {
            // end marker first so positions for the start marker are stable
            for (para_idx, run_idx, is_start) in [
                (last.para, last.index, false),
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
    }
}

// ---------------------------------------------------------------- Hyperlink handle

/// Pure-Rust hyperlink handle (w:hyperlink inside a body paragraph).
#[derive(Debug, Clone, Copy)]
pub struct Hyperlink {
    pub para: usize,
    /// index among the paragraph's w:hyperlink children
    pub index: usize,
}

impl Hyperlink {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let index = self.index;
        para_read(core, self.para, |p| {
            nth_direct_ref(p, "w:hyperlink", index).map(|h| f(h))
        })
        .flatten()
    }

    /// Visible text of the hyperlink.
    pub fn text(self, core: &mut TplCore) -> String {
        self.read(core, element_text).unwrap_or_default()
    }

    /// The URL the hyperlink points to ("" for internal jumps).
    pub fn address(self, core: &mut TplCore) -> String {
        let rid = self
            .read(core, |h| h.get_attr("r:id").map(|s| s.to_string()))
            .flatten();
        let Some(rid) = rid else { return String::new() };
        core.init_docx(false).ok();
        core.package
            .as_ref()
            .map(|pkg| {
                pkg.rels(DOCUMENT_PART)
                    .rels
                    .iter()
                    .find(|r| r.id == rid)
                    .map(|r| r.target.clone())
                    .unwrap_or_default()
            })
            .unwrap_or_default()
    }

    /// Fragment reference (w:anchor), e.g. a bookmark name.
    pub fn fragment(self, core: &mut TplCore) -> String {
        self.read(core, |h| h.get_attr("w:anchor").map(|s| s.to_string()))
            .flatten()
            .unwrap_or_default()
    }

    /// True when the hyperlink text is broken across pages
    /// (w:lastRenderedPageBreak present).
    pub fn contains_page_break(self, core: &mut TplCore) -> bool {
        self.read(core, |h| {
            let mut out = Vec::new();
            h.iter_descendants("w:lastRenderedPageBreak", &mut out);
            !out.is_empty()
        })
        .unwrap_or(false)
    }
}

// ---------------------------------------------------------------- fields

/// Span of a field within a paragraph's children.
#[derive(Debug)]
pub enum FieldSpan {
    /// w:fldSimple child position
    Simple(usize),
    /// (begin run, separate run, end run) child positions
    Complex(usize, Option<usize>, usize),
}

/// Locate all fields of a paragraph, in document order.
pub fn field_spans(p: &Element) -> Vec<FieldSpan> {
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

/// Pure-Rust field handle (w:fldSimple / complex field in a body paragraph).
#[derive(Debug, Clone, Copy)]
pub struct Field {
    pub para: usize,
    /// index among the paragraph's fields, in document order
    pub index: usize,
}

impl Field {
    fn with_span<R>(self, core: &mut TplCore, f: impl FnOnce(&Element, &FieldSpan) -> R) -> Option<R> {
        let index = self.index;
        para_read(core, self.para, move |p| {
            field_spans(p).get(index).map(|s| f(p, s))
        })
        .flatten()
    }

    fn with_span_mut<R>(
        self,
        core: &mut TplCore,
        f: impl FnOnce(&mut Element, &FieldSpan) -> R,
    ) -> Result<R, String> {
        let index = self.index;
        let mut out = None;
        para_edit(core, self.para, |p| {
            let spans = field_spans(p);
            if let Some(s) = spans.get(index) {
                // field spans were computed on the unmodified tree; f applies
                // them in a single mutation
                out = Some(f(p, s));
            }
        })?;
        out.ok_or_else(|| "field not found".to_string())
    }

    /// "simple" (w:fldSimple) or "complex" (fldChar-based).
    pub fn kind(self, core: &mut TplCore) -> Option<String> {
        self.with_span(core, |_, s| match s {
            FieldSpan::Simple(_) => "simple".to_string(),
            FieldSpan::Complex(..) => "complex".to_string(),
        })
    }

    /// The field instruction (e.g. `PAGE \* MERGEFORMAT`), trimmed.
    pub fn instr(self, core: &mut TplCore) -> Option<String> {
        self.with_span(core, |p, s| match s {
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

    pub fn set_instr(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.with_span_mut(core, |p, s| match s {
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
    pub fn text(self, core: &mut TplCore) -> Option<String> {
        self.with_span(core, |p, s| match s {
            FieldSpan::Simple(i) => match &p.children[*i] {
                Node::Elem(e) => element_text(e),
                _ => String::new(),
            },
            FieldSpan::Complex(_, sep, e) => {
                let mut text = String::new();
                if let Some(sep) = sep {
                    for c in &p.children[*sep + 1..*e] {
                        if let Node::Elem(r) = c {
                            text.push_str(&element_text(r));
                        }
                    }
                }
                text
            }
        })
    }

    pub fn set_text(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.with_span_mut(core, |p, s| {
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v.to_string()));
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

// ---------------------------------------------------------------- Table handle

/// Pure-Rust table handle (body-level w:tbl by index).
#[derive(Debug, Clone, Copy)]
pub struct Table {
    pub index: usize,
}

impl Table {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        tbl_read(core, self.index, f)
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        tbl_edit(core, self.index, f)
    }

    pub fn row_count(self, core: &mut TplCore) -> usize {
        self.read(core, |t| {
            t.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tr"))
                .count()
        })
        .unwrap_or(0)
    }

    pub fn add_row(self, core: &mut TplCore) -> Result<TableRow, String> {
        let index = self.index;
        let row = self.edit(core, |t| {
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
            n
        })?;
        Ok(TableRow { index, row })
    }

    /// The table style id.
    pub fn style(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |t| {
            t.find("w:tblPr")
                .and_then(|p| p.find("w:tblStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }

    /// Accepts a style name or id (resolved via the styles part).
    pub fn set_style(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let sid = crate::subdocbuilder::resolve_style_id(core, v);
        self.edit(core, |t| {
            // w:tblPr must be the first child of w:tbl
            let tblpr = ensure_tblpr(t);
            // w:tblStyle must be the first child of w:tblPr
            if tblpr.find("w:tblStyle").is_none() {
                tblpr.children.insert(0, Node::Elem(Element::new("w:tblStyle")));
            }
            tblpr.find_mut("w:tblStyle").unwrap().set_attr("w:val", &sid);
        })
    }

    /// Table alignment as a WD_TABLE_ALIGNMENT int.
    pub fn alignment(self, core: &mut TplCore) -> Option<i64> {
        self.read(core, |t| {
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
    /// val: xml vocabulary ("left"/"center"/"right"); None removes w:jc.
    pub fn set_alignment(self, core: &mut TplCore, val: Option<&str>) -> Result<(), String> {
        self.edit(core, |t| {
            let tblpr = ensure_tblpr(t);
            match val {
                Some(x) => {
                    let jc = ensure_child(tblpr, "w:jc");
                    jc.set_attr("w:val", x);
                }
                None => tblpr
                    .children
                    .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:jc")),
            }
        })
    }

    /// Autofit (tblLayout type=autofit vs fixed; missing -> True).
    pub fn autofit(self, core: &mut TplCore) -> bool {
        self.read(core, |t| {
            t.find("w:tblPr")
                .and_then(|p| p.find("w:tblLayout"))
                .and_then(|e| e.get_attr("w:type"))
                .map(|ty| ty != "fixed")
                .unwrap_or(true)
        })
        .unwrap_or(true)
    }
    pub fn set_autofit(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.edit(core, |t| {
            let tblpr = ensure_tblpr(t);
            let l = ensure_child(tblpr, "w:tblLayout");
            l.set_attr("w:type", if v { "autofit" } else { "fixed" });
        })
    }

    /// Append a column of the given width in EMU (gridCol + one cell per row).
    pub fn add_column(self, core: &mut TplCore, emu: i64) -> Result<Column, String> {
        let index = self.index;
        let twips = (emu / 635).to_string();
        let col = self.edit(core, |t| {
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
        Ok(Column { index, col })
    }

    pub fn column_count(self, core: &mut TplCore) -> usize {
        self.read(core, |t| {
            t.find("w:tblGrid")
                .map(|g| {
                    g.children
                        .iter()
                        .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:gridCol"))
                        .count()
                })
                .unwrap_or(0)
        })
        .unwrap_or(0)
    }

    /// Table direction: 0=ltr, 1=rtl (w:bidiVisual).
    pub fn table_direction(self, core: &mut TplCore) -> i64 {
        self.read(core, |t| {
            t.find("w:tblPr")
                .map(|p| p.find("w:bidiVisual").is_some() as i64)
        })
        .flatten()
        .unwrap_or(0)
    }
    pub fn set_table_direction(self, core: &mut TplCore, v: i64) -> Result<(), String> {
        self.edit(core, |t| {
            let tblpr = ensure_tblpr(t);
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

    /// Number of cells in row `i` (for row_cells).
    pub fn row_cell_count(self, core: &mut TplCore, i: usize) -> usize {
        self.read(core, |t| {
            nth_direct_ref(t, "w:tr", i)
                .map(|r| {
                    r.children
                        .iter()
                        .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                        .count()
                })
                .unwrap_or(0)
        })
        .unwrap_or(0)
    }
}

/// Pure-Rust table column handle (w:tblGrid/w:gridCol proxy).
#[derive(Debug, Clone, Copy)]
pub struct Column {
    pub index: usize,
    pub col: usize,
}

impl Column {
    pub fn width(self, core: &mut TplCore) -> Option<Length> {
        let col = self.col;
        tbl_read(core, self.index, |t| {
            t.find("w:tblGrid")
                .and_then(|g| nth_direct_ref(g, "w:gridCol", col))
                .and_then(|gc| gc.get_attr("w:w"))
                .and_then(|s| s.parse::<i64>().ok())
                .map(Length::from_twips)
        })
        .flatten()
    }

    /// emu: EMU; None removes the width attribute (creating grid columns as
    /// needed, python-docx semantics).
    pub fn set_width(self, core: &mut TplCore, emu: Option<i64>) -> Result<(), String> {
        let col = self.col;
        tbl_edit(core, self.index, |t| {
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
            if let Some(gc) = nth_direct(grid, "w:gridCol", col) {
                match emu {
                    Some(e) => gc.set_attr("w:w", &(e / 635).to_string()),
                    None => remove_attr(gc, "w:w"),
                }
            }
        })
    }
}

// ---------------------------------------------------------------- TableRow handle

/// Pure-Rust table row handle.
#[derive(Debug, Clone, Copy)]
pub struct TableRow {
    pub index: usize,
    pub row: usize,
}

impl TableRow {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        row_read(core, self.index, self.row, f)
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        row_edit(core, self.index, self.row, f)
    }

    pub fn cell_count(self, core: &mut TplCore) -> usize {
        self.read(core, |r| {
            r.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                .count()
        })
        .unwrap_or(0)
    }

    /// Row height (w:trPr/w:trHeight w:val).
    pub fn height(self, core: &mut TplCore) -> Option<Length> {
        self.read(core, |r| {
            r.find("w:trPr")
                .and_then(|p| p.find("w:trHeight"))
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
        })
        .flatten()
        .map(Length::from_twips)
    }
    /// None removes the w:val attribute (element kept).
    pub fn set_height(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.edit(core, |r| {
            let trpr = ensure_trpr(r);
            let h = ensure_child(trpr, "w:trHeight");
            match v {
                Some(l) => h.set_attr("w:val", &l.twips().to_string()),
                None => remove_attr(h, "w:val"),
            }
        })
    }

    /// Row height rule as a WD_ROW_HEIGHT_RULE int (0=auto, 1=atLeast, 2=exact).
    pub fn height_rule(self, core: &mut TplCore) -> Option<i64> {
        self.read(core, |r| {
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
    /// val: xml vocabulary; None removes the w:hRule attribute (element kept).
    pub fn set_height_rule(self, core: &mut TplCore, val: Option<&str>) -> Result<(), String> {
        self.edit(core, |r| {
            let trpr = ensure_trpr(r);
            let h = ensure_child(trpr, "w:trHeight");
            match val {
                Some(x) => h.set_attr("w:hRule", x),
                None => remove_attr(h, "w:hRule"),
            }
        })
    }

    fn grid_cols(self, core: &mut TplCore, tag: &str) -> i64 {
        self.read(core, |r| {
            r.find("w:trPr")
                .and_then(|p| p.find(tag))
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
                .unwrap_or(0)
        })
        .unwrap_or(0)
    }
    /// Grid columns before this row (trPr/gridBefore; default 0).
    pub fn grid_cols_before(self, core: &mut TplCore) -> i64 {
        self.grid_cols(core, "w:gridBefore")
    }
    /// Grid columns after this row (trPr/gridAfter; default 0).
    pub fn grid_cols_after(self, core: &mut TplCore) -> i64 {
        self.grid_cols(core, "w:gridAfter")
    }
}

// ---------------------------------------------------------------- Cell handle

/// Pure-Rust table cell handle.
#[derive(Debug, Clone, Copy)]
pub struct Cell {
    pub index: usize,
    pub row: usize,
    pub col: usize,
}

impl Cell {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        cell_read(core, self.index, self.row, self.col, f)
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        cell_edit(core, self.index, self.row, self.col, f)
    }

    pub fn text(self, core: &mut TplCore) -> String {
        self.read(core, element_text).unwrap_or_default()
    }

    pub fn set_text(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |c| {
            c.children
                .retain(|ch| !matches!(ch, Node::Elem(e) if e.name == "w:p"));
            let mut p = Element::new("w:p");
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v.to_string()));
            r.children.push(Node::Elem(t));
            p.children.push(Node::Elem(r));
            c.children.push(Node::Elem(p));
        })
    }

    pub fn paragraph_count(self, core: &mut TplCore) -> usize {
        self.read(core, |c| {
            c.children
                .iter()
                .filter(|ch| matches!(ch, Node::Elem(e) if e.name == "w:p"))
                .count()
        })
        .unwrap_or(0)
    }

    /// Append a paragraph to this cell (python-docx cell.add_paragraph);
    /// style: style name or id.
    pub fn add_paragraph(
        self,
        core: &mut TplCore,
        text: &str,
        style: Option<&str>,
    ) -> Result<CellParagraph, String> {
        let sid = style.map(|s| crate::subdocbuilder::resolve_style_id(core, s));
        let text = text.to_string();
        let (index, row, col) = (self.index, self.row, self.col);
        let para = self.edit(core, |c| {
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
        Ok(CellParagraph {
            index,
            row,
            col,
            para,
        })
    }

    /// Vertical alignment as a WD_CELL_VERTICAL_ALIGNMENT int.
    pub fn vertical_alignment(self, core: &mut TplCore) -> Option<i64> {
        self.read(core, |c| {
            c.find("w:tcPr")
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
    }
    /// val: xml vocabulary; None removes w:vAlign.
    pub fn set_vertical_alignment(self, core: &mut TplCore, val: Option<&str>) -> Result<(), String> {
        self.edit(core, |c| {
            let tcpr = tcpr_mut(c);
            match val {
                Some(x) => {
                    let va = ensure_child(tcpr, "w:vAlign");
                    va.set_attr("w:val", x);
                }
                None => tcpr
                    .children
                    .retain(|ch| !matches!(ch, Node::Elem(e) if e.name == "w:vAlign")),
            }
        })
    }

    /// Cell width (w:tcPr/w:tcW, dxa only).
    pub fn width(self, core: &mut TplCore) -> Option<Length> {
        self.read(core, |c| {
            c.find("w:tcPr")
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
        .map(Length::from_twips)
    }
    /// Some forces type=dxa; None removes w:tcW.
    pub fn set_width(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.edit(core, |c| {
            let tcpr = tcpr_mut(c);
            match v {
                Some(l) => {
                    let tcw = ensure_child(tcpr, "w:tcW");
                    tcw.set_attr("w:type", "dxa");
                    tcw.set_attr("w:w", &l.twips().to_string());
                }
                None => tcpr
                    .children
                    .retain(|ch| !matches!(ch, Node::Elem(e) if e.name == "w:tcW")),
            }
        })
    }

    /// Grid columns spanned by this cell (w:gridSpan; default 1).
    pub fn grid_span(self, core: &mut TplCore) -> i64 {
        self.read(core, |c| {
            c.find("w:tcPr")
                .and_then(|p| p.find("w:gridSpan"))
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
                .unwrap_or(1)
        })
        .unwrap_or(1)
    }

    pub fn table_count(self, core: &mut TplCore) -> usize {
        self.read(core, |c| {
            c.children
                .iter()
                .filter(|ch| matches!(ch, Node::Elem(e) if e.name == "w:tbl"))
                .count()
        })
        .unwrap_or(0)
    }

    /// Append a rows x cols table to this cell (python-docx cell.add_table).
    pub fn add_table(self, core: &mut TplCore, rows: usize, cols: usize) -> Result<CellTable, String> {
        let (index, row, col) = (self.index, self.row, self.col);
        let tindex = self.edit(core, |c| {
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
        Ok(CellTable {
            index,
            row,
            col,
            tindex,
        })
    }

    /// Paragraphs and tables of this cell in document order.
    pub fn iter_inner_content(self, core: &mut TplCore) -> Vec<BlockItem> {
        self.read(core, |c| {
            let mut pi = 0usize;
            let mut ti = 0usize;
            let mut out = Vec::new();
            for ch in &c.children {
                match ch {
                    Node::Elem(e) if e.name == "w:p" => {
                        out.push(BlockItem::Paragraph(pi));
                        pi += 1;
                    }
                    Node::Elem(e) if e.name == "w:tbl" => {
                        out.push(BlockItem::Table(ti));
                        ti += 1;
                    }
                    _ => {}
                }
            }
            out
        })
        .unwrap_or_default()
    }

    /// Merge this cell with `other` into one cell spanning the rectangular
    /// region between them (python-docx cell.merge). Returns the merged cell.
    pub fn merge(self, core: &mut TplCore, other: Cell) -> Result<Cell, String> {
        if other.index != self.index {
            return Err("cannot merge cells of different tables".to_string());
        }
        let (r1, r2) = (self.row.min(other.row), self.row.max(other.row));
        let (c1, c2) = (self.col.min(other.col), self.col.max(other.col));
        if r1 != r2 || c1 != c2 {
            let mut found = false;
            mutate_document(core, |body| {
                if let Some(t) = nth_direct(body, "w:tbl", self.index) {
                    merge_region(t, r1, r2, c1, c2);
                    found = true;
                }
            })?;
            if !found {
                return Err("table not found".to_string());
            }
        }
        Ok(Cell {
            index: self.index,
            row: r1,
            col: c1,
        })
    }
}

// ---------------------------------------------------------------- cell paragraphs

/// Pure-Rust handle for a paragraph inside a table cell.
#[derive(Debug, Clone, Copy)]
pub struct CellParagraph {
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub para: usize,
}

impl CellParagraph {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let para = self.para;
        cell_read(core, self.index, self.row, self.col, |tc| {
            nth_direct_ref(tc, "w:p", para).map(|p| f(p))
        })
        .flatten()
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        let para = self.para;
        cell_edit(core, self.index, self.row, self.col, |tc| {
            nth_direct(tc, "w:p", para)
                .map(|p| f(p))
                .ok_or_else(|| "paragraph not found".to_string())
        })?
    }

    pub fn text(self, core: &mut TplCore) -> String {
        self.read(core, element_text).unwrap_or_default()
    }

    pub fn set_text(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |p| {
            p.children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:r" || e.name == "w:hyperlink"));
            let mut r = Element::new("w:r");
            let mut t = Element::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.children.push(Node::Text(v.to_string()));
            r.children.push(Node::Elem(t));
            p.children.push(Node::Elem(r));
        })
    }

    pub fn style(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |p| {
            p.find("w:pPr")
                .and_then(|ppr| ppr.find("w:pStyle"))
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }

    /// Accepts a style name or id (resolved via the styles part).
    pub fn set_style(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let sid = crate::subdocbuilder::resolve_style_id(core, v);
        self.edit(core, |p| {
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

// ---------------------------------------------------------------- nested tables

/// Pure-Rust handle for a table nested inside a table cell.
#[derive(Debug, Clone, Copy)]
pub struct CellTable {
    pub index: usize,
    pub row: usize,
    pub col: usize,
    /// index among the cell's direct w:tbl children
    pub tindex: usize,
}

impl CellTable {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        let tindex = self.tindex;
        cell_read(core, self.index, self.row, self.col, |tc| {
            nth_direct_ref(tc, "w:tbl", tindex).map(|t| f(t))
        })
        .flatten()
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        let tindex = self.tindex;
        cell_edit(core, self.index, self.row, self.col, |tc| {
            nth_direct(tc, "w:tbl", tindex)
                .map(|t| f(t))
                .ok_or_else(|| "table not found".to_string())
        })?
    }

    pub fn row_count(self, core: &mut TplCore) -> usize {
        self.read(core, |t| {
            t.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tr"))
                .count()
        })
        .unwrap_or(0)
    }

    pub fn cell(self, nrow: usize, ncol: usize) -> NestedCell {
        NestedCell {
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
            nrow,
            ncol,
        }
    }
}

/// Pure-Rust handle for a row of a nested table.
#[derive(Debug, Clone, Copy)]
pub struct NestedRow {
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub tindex: usize,
    pub nrow: usize,
}

impl NestedRow {
    pub fn cell_count(self, core: &mut TplCore) -> usize {
        let nrow = self.nrow;
        CellTable {
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
        }
        .read(core, |tbl| {
            nth_direct_ref(tbl, "w:tr", nrow).map(|tr| {
                tr.children
                    .iter()
                    .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:tc"))
                    .count()
            })
        })
        .flatten()
        .unwrap_or(0)
    }
}

/// Pure-Rust handle for a cell of a nested table.
#[derive(Debug, Clone, Copy)]
pub struct NestedCell {
    pub index: usize,
    pub row: usize,
    pub col: usize,
    pub tindex: usize,
    pub nrow: usize,
    pub ncol: usize,
}

impl NestedCell {
    fn table(self) -> CellTable {
        CellTable {
            index: self.index,
            row: self.row,
            col: self.col,
            tindex: self.tindex,
        }
    }

    pub fn text(self, core: &mut TplCore) -> String {
        let (nrow, ncol) = (self.nrow, self.ncol);
        self.table()
            .read(core, |tbl| {
                nth_direct_ref(tbl, "w:tr", nrow)
                    .and_then(|tr| nth_direct_ref(tr, "w:tc", ncol))
                    .map(element_text)
            })
            .flatten()
            .unwrap_or_default()
    }

    pub fn set_text(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let (nrow, ncol) = (self.nrow, self.ncol);
        self.table().edit(core, |tbl| {
            if let Some(tr) = nth_direct(tbl, "w:tr", nrow) {
                if let Some(tc) = nth_direct(tr, "w:tc", ncol) {
                    tc.children
                        .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:p"));
                    let mut p = Element::new("w:p");
                    let mut r = Element::new("w:r");
                    let mut wt = Element::new("w:t");
                    wt.set_attr("xml:space", "preserve");
                    wt.children.push(Node::Text(v.to_string()));
                    r.children.push(Node::Elem(wt));
                    p.children.push(Node::Elem(r));
                    tc.children.push(Node::Elem(p));
                }
            }
        })
    }
}

// ---------------------------------------------------------------- Section handle

/// Pure-Rust section handle (w:sectPr by document order).
#[derive(Debug, Clone, Copy)]
pub struct Section {
    pub index: usize,
}

/// pgMar defaults used when creating the element (python-docx behavior).
const PGMAR_DEFAULTS: &[(&str, &str)] = &[
    ("w:left", "1800"),
    ("w:right", "1800"),
    ("w:top", "1440"),
    ("w:bottom", "1440"),
];
const PGSZ_DEFAULTS: &[(&str, &str)] = &[("w:w", "12240"), ("w:h", "15840")];

impl Section {
    pub fn read<R>(self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        sect_read(core, self.index, f)
    }

    pub fn edit<R>(self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        sect_edit(core, self.index, f)
    }

    fn get_dim(self, core: &mut TplCore, tag: &str, attr: &str) -> Option<Length> {
        self.read(core, |sp| get_twips(sp, tag, attr))
            .flatten()
            .map(Length::from_twips)
    }
    fn set_dim(
        self,
        core: &mut TplCore,
        tag: &str,
        attr: &str,
        v: Option<Length>,
        defaults: &[(&str, &str)],
    ) -> Result<(), String> {
        let twips = v.map(|l| l.twips());
        self.edit(core, |sp| set_twips(sp, tag, attr, twips, defaults))
    }

    pub fn page_width(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgSz", "w:w")
    }
    pub fn set_page_width(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgSz", "w:w", v, PGSZ_DEFAULTS)
    }
    pub fn page_height(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgSz", "w:h")
    }
    pub fn set_page_height(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgSz", "w:h", v, PGSZ_DEFAULTS)
    }
    pub fn left_margin(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:left")
    }
    pub fn set_left_margin(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:left", v, PGMAR_DEFAULTS)
    }
    pub fn right_margin(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:right")
    }
    pub fn set_right_margin(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:right", v, PGMAR_DEFAULTS)
    }
    pub fn top_margin(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:top")
    }
    pub fn set_top_margin(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:top", v, PGMAR_DEFAULTS)
    }
    pub fn bottom_margin(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:bottom")
    }
    pub fn set_bottom_margin(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:bottom", v, PGMAR_DEFAULTS)
    }
    pub fn header_distance(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:header")
    }
    pub fn set_header_distance(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:header", v, PGMAR_DEFAULTS)
    }
    pub fn footer_distance(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:footer")
    }
    pub fn set_footer_distance(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:footer", v, PGMAR_DEFAULTS)
    }
    pub fn gutter(self, core: &mut TplCore) -> Option<Length> {
        self.get_dim(core, "w:pgMar", "w:gutter")
    }
    pub fn set_gutter(self, core: &mut TplCore, v: Option<Length>) -> Result<(), String> {
        self.set_dim(core, "w:pgMar", "w:gutter", v, PGMAR_DEFAULTS)
    }

    pub fn orientation(self, core: &mut TplCore) -> Option<String> {
        self.read(core, |sp| {
            sp.find("w:pgSz")
                .and_then(|e| e.get_attr("w:orient").map(|s| s.to_string()))
        })
        .flatten()
    }
    /// Swapping orientation also swaps page dimensions (python-docx).
    pub fn set_orientation(self, core: &mut TplCore, v: &str) -> Result<(), String> {
        let orient = if v.to_lowercase().starts_with("land") { "landscape" } else { "portrait" };
        self.edit(core, |sp| {
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

    /// Section start type as a WD_SECTION_START int; missing w:type reads as
    /// 2 (next page).
    pub fn start_type(self, core: &mut TplCore) -> i64 {
        self.read(core, |sp| {
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
    /// val: xml vocabulary; None or "nextPage" drops the element (default).
    pub fn set_start_type(self, core: &mut TplCore, val: Option<&str>) -> Result<(), String> {
        self.edit(core, |sp| match val {
            None | Some("nextPage") => sp
                .children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:type")),
            Some(x) => {
                let t = ensure_child(sp, "w:type");
                t.set_attr("w:val", x);
            }
        })
    }

    pub fn different_first_page_header_footer(self, core: &mut TplCore) -> bool {
        self.read(core, |sp| sp.find("w:titlePg").is_some())
            .unwrap_or(false)
    }
    pub fn set_different_first_page_header_footer(self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.edit(core, |sp| {
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

    /// Paragraphs and tables of this section in document order. Section
    /// boundaries are the paragraphs carrying a paragraph-level sectPr; the
    /// last section ends at the body-level sectPr.
    pub fn iter_inner_content(self, core: &mut TplCore) -> Vec<BlockItem> {
        read_body(core, |body| {
            // section ranges: items per section
            let mut sections: Vec<Vec<BlockItem>> = vec![Vec::new()];
            let mut pi = 0usize;
            let mut ti = 0usize;
            for c in &body.children {
                let Node::Elem(e) = c else { continue };
                match e.name.as_str() {
                    "w:p" => {
                        sections.last_mut().unwrap().push(BlockItem::Paragraph(pi));
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
                        sections.last_mut().unwrap().push(BlockItem::Table(ti));
                        ti += 1;
                    }
                    _ => {}
                }
            }
            sections.get(self.index).cloned().unwrap_or_default()
        })
        .unwrap_or_default()
    }
}

/// All body-level paragraphs, in document order.
pub fn paragraphs(core: &mut TplCore) -> Vec<Paragraph> {
    (0..count_in_body(core, "w:p"))
        .map(|index| Paragraph { index })
        .collect()
}

/// All body-level tables, in document order.
pub fn tables(core: &mut TplCore) -> Vec<Table> {
    (0..count_in_body(core, "w:tbl"))
        .map(|index| Table { index })
        .collect()
}

/// All sections (one per w:sectPr, in document order).
pub fn sections(core: &mut TplCore) -> Vec<Section> {
    (0..section_count(core))
        .map(|index| Section { index })
        .collect()
}

/// Number of sections (w:sectPr count in the document).
pub fn section_count(core: &mut TplCore) -> usize {
    core.document_dom()
        .map(|dom| {
            let mut v: Vec<&Element> = Vec::new();
            dom.root.iter_descendants("w:sectPr", &mut v);
            v.len()
        })
        .unwrap_or(0)
}

// ---------------------------------------------------------------- headers/footers

/// Pure-Rust header/footer handle of a section.
/// kind: "header" | "footer" | "even_header" | "even_footer" |
/// "first_header" | "first_footer".
#[derive(Debug, Clone)]
pub struct HdrFtr {
    pub section: usize,
    pub kind: String,
}

/// Split a hdrftr kind into (w:type value, base "header"|"footer").
pub fn split_kind(kind: &str) -> (&'static str, &'static str) {
    match kind {
        "footer" => ("default", "footer"),
        "even_header" => ("even", "header"),
        "even_footer" => ("even", "footer"),
        "first_header" => ("first", "header"),
        "first_footer" => ("first", "footer"),
        _ => ("default", "header"),
    }
}

fn rel_type_for(kind: &str) -> &'static str {
    let (_, base) = split_kind(kind);
    if base == "header" {
        crate::package::rel_type::HEADER
    } else {
        crate::package::rel_type::FOOTER
    }
}

/// find the header/footer part path linked to a section (by headerReference
/// order within its sectPr), if any; returns (rid, part_path)
pub fn find_hdrftr_part(
    core: &mut TplCore,
    section_idx: usize,
    kind: &str,
) -> Option<(String, String)> {
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

fn remove_hdrftr_reference(core: &mut TplCore, section_idx: usize, kind: &str) -> Result<(), String> {
    mutate_document(core, |body| {
        let mut sects: Vec<&mut Element> = Vec::new();
        collect_sectprs_mut(body, &mut sects);
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
        collect_sectprs_mut(body, &mut sects);
        if let Some(sect) = sects.get_mut(section_idx) {
            let mut el = Element::new(tag);
            el.set_attr("w:type", wtype);
            el.set_attr("r:id", &rid);
            sect.children.insert(0, Node::Elem(el));
        }
    })
}

impl HdrFtr {
    pub fn is_linked_to_previous(&self, core: &mut TplCore) -> bool {
        find_hdrftr_part(core, self.section, &self.kind).is_none()
    }

    pub fn set_is_linked_to_previous(&self, core: &mut TplCore, v: bool) -> Result<(), String> {
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
    }

    /// Texts of the paragraphs in this header/footer part.
    pub fn paragraphs(&self, core: &mut TplCore) -> Vec<String> {
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
    }

    pub fn add_paragraph(&self, core: &mut TplCore, text: &str) -> Result<(), String> {
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
    }
}

// ---------------------------------------------------------------- Style handle

/// Pure-Rust style handle (w:style by styleId).
#[derive(Debug, Clone)]
pub struct Style {
    pub style_id: String,
}

impl Style {
    pub fn read<R>(&self, core: &mut TplCore, f: impl FnOnce(&Element) -> R) -> Option<R> {
        style_read(core, &self.style_id, f)
    }

    pub fn edit<R>(&self, core: &mut TplCore, f: impl FnOnce(&mut Element) -> R) -> Result<R, String> {
        style_edit(core, &self.style_id, f)
    }

    pub fn name(&self, core: &mut TplCore) -> Option<String> {
        self.read(core, |st| {
            st.find("w:name")
                .and_then(|n| n.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    pub fn set_name(&self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |st| {
            if let Some(n) = st.find_mut("w:name") {
                n.set_attr("w:val", v);
            } else {
                let mut n = Element::new("w:name");
                n.set_attr("w:val", v);
                st.children.insert(0, Node::Elem(n));
            }
        })
    }

    pub fn style_type(&self, core: &mut TplCore) -> Option<String> {
        self.read(core, |st| st.get_attr("w:type").map(|s| s.to_string()))
            .flatten()
    }

    pub fn base_style(&self, core: &mut TplCore) -> Option<String> {
        self.read(core, |st| {
            st.find("w:basedOn")
                .and_then(|b| b.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    pub fn set_base_style(&self, core: &mut TplCore, v: &str) -> Result<(), String> {
        self.edit(core, |st| {
            if let Some(b) = st.find_mut("w:basedOn") {
                b.set_attr("w:val", v);
            } else {
                let mut b = Element::new("w:basedOn");
                b.set_attr("w:val", v);
                st.children.push(Node::Elem(b));
            }
        })
    }

    fn flag(&self, core: &mut TplCore, tag: &str) -> bool {
        self.read(core, |st| st.find(tag).is_some()).unwrap_or(false)
    }
    fn set_flag(&self, core: &mut TplCore, tag: &str, v: bool) -> Result<(), String> {
        self.edit(core, |st| style_flag(st, tag, v))
    }

    /// Hidden in the UI until used (w:semiHidden).
    pub fn hidden(&self, core: &mut TplCore) -> bool {
        self.flag(core, "w:semiHidden")
    }
    pub fn set_hidden(&self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:semiHidden", v)
    }
    /// Locked against editing (w:locked).
    pub fn locked(&self, core: &mut TplCore) -> bool {
        self.flag(core, "w:locked")
    }
    pub fn set_locked(&self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:locked", v)
    }
    /// Shown in the quick style gallery (w:qFormat).
    pub fn quick_style(&self, core: &mut TplCore) -> bool {
        self.flag(core, "w:qFormat")
    }
    pub fn set_quick_style(&self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:qFormat", v)
    }
    /// Re-hide when the style is no longer used (w:unhideWhenUsed).
    pub fn unhide_when_used(&self, core: &mut TplCore) -> bool {
        self.flag(core, "w:unhideWhenUsed")
    }
    pub fn set_unhide_when_used(&self, core: &mut TplCore, v: bool) -> Result<(), String> {
        self.set_flag(core, "w:unhideWhenUsed", v)
    }

    /// UI priority (w:uiPriority w:val); None removes it.
    pub fn priority(&self, core: &mut TplCore) -> Option<i64> {
        self.read(core, |st| {
            st.find("w:uiPriority")
                .and_then(|e| e.get_attr("w:val"))
                .and_then(|s| s.parse::<i64>().ok())
        })
        .flatten()
    }
    pub fn set_priority(&self, core: &mut TplCore, v: Option<i64>) -> Result<(), String> {
        self.edit(core, |st| match v {
            Some(n) => {
                let e = ensure_child(st, "w:uiPriority");
                e.set_attr("w:val", &n.to_string());
            }
            None => st
                .children
                .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:uiPriority")),
        })
    }

    /// Builtin styles lack the w:customStyle attribute (read-only).
    pub fn builtin(&self, core: &mut TplCore) -> bool {
        self.read(core, |st| {
            !matches!(st.get_attr("w:customStyle"), Some("1") | Some("true") | Some("on"))
        })
        .unwrap_or(true)
    }

    /// Style applied to the next paragraph (w:next; paragraph styles).
    pub fn next_paragraph_style(&self, core: &mut TplCore) -> Option<String> {
        self.read(core, |st| {
            st.find("w:next")
                .and_then(|e| e.get_attr("w:val").map(|s| s.to_string()))
        })
        .flatten()
    }
    /// Some: style name or id (resolved); None removes w:next.
    pub fn set_next_paragraph_style(&self, core: &mut TplCore, v: Option<&str>) -> Result<(), String> {
        match v {
            None => self.edit(core, |st| {
                st.children
                    .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:next"));
            }),
            Some(name) => {
                let sid = crate::subdocbuilder::resolve_style_id(core, name);
                self.edit(core, |st| {
                    let e = ensure_child(st, "w:next");
                    e.set_attr("w:val", &sid);
                })
            }
        }
    }

    pub fn delete(&self, core: &mut TplCore) -> Result<(), String> {
        with_styles(core, |root| {
            root.children.retain(|c| {
                !(matches!(c, Node::Elem(e) if e.name == "w:style" && e.get_attr("w:styleId") == Some(self.style_id.as_str())))
            });
        })
    }

    /// Full python-docx Font facade over the style's w:rPr.
    pub fn font(&self) -> Font {
        Font {
            target: FontTarget::Style {
                style_id: self.style_id.clone(),
            },
        }
    }

    /// Paragraph formatting of the style.
    pub fn paragraph_format(&self) -> ParagraphFormat {
        ParagraphFormat {
            target: PfTarget::Style {
                style_id: self.style_id.clone(),
            },
        }
    }
}

/// All style ids in document order.
pub fn style_ids(core: &mut TplCore) -> Vec<String> {
    core.part_dom("word/styles.xml")
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
}

/// Create a new style; returns its generated style id.
/// type_str: "paragraph" | "character" | "table" | "numbering".
pub fn add_style(core: &mut TplCore, name: &str, type_str: &str, builtin: bool) -> Result<String, String> {
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
}

// ---------------------------------------------------------------- settings

/// Read a boolean flag element from word/settings.xml (missing -> false).
pub fn settings_flag(core: &mut TplCore, tag: &str) -> bool {
    core.part_dom("word/settings.xml")
        .map(|dom| dom.root.find(tag).is_some())
        .unwrap_or(false)
}

/// Write a boolean flag element to word/settings.xml (creating the part).
pub fn set_settings_flag(core: &mut TplCore, tag: &str, v: bool) -> Result<(), String> {
    with_settings(core, |root| {
        let exists = root.find(tag).is_some();
        if v && !exists {
            root.children.push(Node::Elem(Element::new(tag)));
        } else if !v && exists {
            root.children.retain(|c| !matches!(c, Node::Elem(e) if e.name == tag));
        }
    })
}

// ---------------------------------------------------------------- core properties

/// python-docx core property attribute -> xml tag.
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
    core.part_dom("docProps/core.xml")
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

// ---------------------------------------------------------------- inline shapes

/// (width EMU, height EMU, kind) of each inline shape (wp:inline).
pub fn inline_shapes(core: &mut TplCore) -> Vec<(i64, i64, String)> {
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
                out.push((cx, cy, "picture".to_string()));
            }
            out
        })
        .unwrap_or_default()
}

// ---------------------------------------------------------------- comments

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
    let dom = core.part_dom("word/comments.xml")?;
    let mut max_id: i64 = -1;
    let mut comments: Vec<&Element> = Vec::new();
    dom.root.iter_descendants("w:comment", &mut comments);
    for c in comments {
        if let Some(Ok(n)) = c.get_attr("w:id").map(|v| v.parse::<i64>()) {
            max_id = max_id.max(n);
        }
    }
    Ok(max_id)
}

/// Append a comment entry to the comments part, returns its id.
pub fn append_comment(
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

    let frag = Document::parse(&comment).map_err(|e| format!("bad comment xml: {}", e))?;
    let dom = core.part_dom("word/comments.xml")?;
    dom.root.children.push(Node::Elem(frag.root));
    core.mark_part_dirty("word/comments.xml");
    Ok(id)
}

/// All comment ids in document order.
pub fn comment_ids(core: &mut TplCore) -> Vec<i64> {
    core.part_dom("word/comments.xml")
        .map(|dom| {
            let mut out = Vec::new();
            let mut comments: Vec<&Element> = Vec::new();
            dom.root.iter_descendants("w:comment", &mut comments);
            for c in comments {
                if let Some(id) = c.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()) {
                    out.push(id);
                }
            }
            out
        })
        .unwrap_or_default()
}

/// Read the comment element with the given id.
pub fn comment_read<R>(core: &mut TplCore, comment_id: i64, f: impl FnOnce(&Element) -> R) -> Option<R> {
    let dom = core.part_dom("word/comments.xml").ok()?;
    let mut comments: Vec<&Element> = Vec::new();
    dom.root.iter_descendants("w:comment", &mut comments);
    comments
        .into_iter()
        .find(|c| c.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()) == Some(comment_id))
        .map(|c| f(c))
}

/// Anchor a comment to the runs between `first` and `last` (para, index)
/// pairs (Document.add_comment).
pub fn anchor_comment(
    core: &mut TplCore,
    first: (usize, usize),
    last: (usize, usize),
    comment_id: i64,
) -> Result<(), String> {
    mutate_document(core, |body| {
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
}

// ---------------------------------------------------------------- add_* core

/// Append a paragraph to the document body (python-docx add_paragraph);
/// style: style name or id. Returns a handle to the new paragraph.
pub fn add_paragraph(core: &mut TplCore, text: &str, style: Option<&str>) -> Result<Paragraph, String> {
    let sid = style.map(|s| crate::subdocbuilder::resolve_style_id(core, s));
    let mut p = String::from("<w:p>");
    if let Some(sid) = &sid {
        p.push_str(&format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", sid));
    }
    if !text.is_empty() {
        p.push_str(&crate::richtext::richtext_run(text, &crate::richtext::TextProps::default()));
    }
    p.push_str("</w:p>");
    append_to_body(core, &p)?;
    Ok(Paragraph {
        index: count_in_body(core, "w:p") - 1,
    })
}

/// Append a heading paragraph (python-docx add_heading).
pub fn add_heading(core: &mut TplCore, text: &str, level: u32) -> Result<Paragraph, String> {
    add_paragraph(core, text, Some(&format!("Heading {}", level.max(1))))
}

/// Append a picture paragraph (python-docx add_picture).
pub fn add_picture(
    core: &mut TplCore,
    blob: &[u8],
    filename: Option<&str>,
    width: Option<i64>,
    height: Option<i64>,
) -> Result<(), String> {
    core.init_docx(false)?;
    let drawing = crate::inline_image::drawing_xml(
        core,
        DOCUMENT_PART,
        blob,
        filename,
        width,
        height,
        None,
        None,
        None,
    )?;
    append_to_body(core, &format!("<w:p><w:r>{}</w:r></w:p>", drawing))
}

/// Append a table (python-docx add_table); returns a handle to it.
pub fn add_table(core: &mut TplCore, rows: usize, cols: usize) -> Result<Table, String> {
    let usable = crate::subdocbuilder::master_usable_width_twips(core);
    let xml = crate::subdocbuilder::table_xml(&vec![vec![String::new(); cols]; rows], usable);
    append_to_body(core, &xml)?;
    Ok(Table {
        index: count_in_body(core, "w:tbl") - 1,
    })
}

/// Append a page break paragraph (python-docx add_page_break).
pub fn add_page_break(core: &mut TplCore) -> Result<(), String> {
    append_to_body(core, "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>")
}

/// python-docx add_section: close the current section with a paragraph-level
/// sectPr copy and set the body sectPr's start type (xml vocabulary).
pub fn add_section(core: &mut TplCore, type_str: &str) -> Result<Section, String> {
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
        // find the current body-level sectPr and set its type
        let mut sects: Vec<&mut Element> = Vec::new();
        collect_sectprs_mut(body, &mut sects);
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
    Ok(Section {
        index: section_count(core) - 1,
    })
}


#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write as _;

    /// Build a minimal docx with the given body xml.
    fn make_docx(body: &str) -> Vec<u8> {
        let ct = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
        let doc = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}<w:sectPr/></w:body></w:document>",
            body
        );
        let mut cursor = std::io::Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut cursor);
            let opts = zip::write::SimpleFileOptions::default();
            for (name, data) in [
                ("[Content_Types].xml", &ct[..]),
                ("_rels/.rels", &rels[..]),
                (DOCUMENT_PART, doc.as_bytes()),
            ] {
                w.start_file(name, opts).unwrap();
                w.write_all(data).unwrap();
            }
            w.finish().unwrap();
        }
        cursor.into_inner()
    }

    fn tp(text: &str) -> String {
        format!("<w:p><w:r><w:t>{}</w:t></w:r></w:p>", text)
    }

    fn core_of(body: &str) -> TplCore {
        let mut core = TplCore::new(make_docx(body));
        core.init_docx(false).unwrap();
        core
    }

    #[test]
    fn test_paragraph_and_run_handles() {
        let mut core = core_of(&(tp("hello") + &tp("world")));
        let paras = paragraphs(&mut core);
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text(&mut core), "hello");
        paras[1].set_text(&mut core, "rust").unwrap();
        assert_eq!(paras[1].text(&mut core), "rust");
        // runs
        assert_eq!(paras[0].run_count(&mut core), 1);
        let run0 = Run { para: 0, index: 0 };
        run0.set_text(&mut core, "hi").unwrap();
        assert_eq!(run0.text(&mut core), "hi");
        // run font tri-state
        let font = run0.font();
        assert_eq!(font.bold(&mut core), None);
        font.set_bold(&mut core, Some(true)).unwrap();
        assert_eq!(font.bold(&mut core), Some(true));
    }

    #[test]
    fn test_paragraph_format_and_alignment() {
        let mut core = core_of(&tp("x"));
        let p = paragraphs(&mut core)[0];
        assert_eq!(p.alignment(&mut core), None);
        p.set_alignment(&mut core, Some(1)).unwrap(); // center
        assert_eq!(p.alignment(&mut core), Some(1));
        let pf = p.paragraph_format();
        pf.set_space_after(&mut core, Some(12 * 12700)).unwrap(); // Pt(12) in emu
        assert_eq!(pf.space_after(&mut core).unwrap().emu, 12 * 12700);
    }

    #[test]
    fn test_table_handles() {
        let body = "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>\
                    <w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        let mut core = core_of(body);
        let tbls = tables(&mut core);
        assert_eq!(tbls.len(), 1);
        let cell = Cell { index: 0, row: 0, col: 0 };
        assert_eq!(cell.text(&mut core), "a");
        cell.set_text(&mut core, "b").unwrap();
        assert_eq!(cell.text(&mut core), "b");
        assert_eq!(tbls[0].row_count(&mut core), 1);
    }

    #[test]
    fn test_add_paragraph_and_sections() {
        let mut core = core_of(&tp("first"));
        let p = add_paragraph(&mut core, "second", None).unwrap();
        assert_eq!(p.text(&mut core), "second");
        assert_eq!(paragraphs(&mut core).len(), 2);
        let secs = sections(&mut core);
        assert_eq!(secs.len(), 1);
        assert_eq!(secs[0].start_type(&mut core), 2); // missing w:type -> nextPage
    }

    #[test]
    fn test_pure_rust_save_roundtrip() {
        let mut core = core_of(&tp("x"));
        add_paragraph(&mut core, "added", None).unwrap();
        let bytes = core.save_bytes().unwrap();
        let pkg = crate::package::Package::from_bytes(&bytes).unwrap();
        let xml = pkg.get_string(DOCUMENT_PART).unwrap();
        assert!(xml.contains(">x<") && xml.contains(">added<"));
    }
}
