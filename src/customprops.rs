//! Dissolving DOCPROPERTY fields that reference custom document properties
//! (docProps/custom.xml) into literal values — the Composer's
//! `remove_property_fields` option (docxcompose
//! `CustomProperties.dissolve_fields` parity): fields naming a property that
//! exists in custom.xml are replaced by the property's value; fields naming
//! unknown properties (or all fields when there is no custom.xml) are left
//! untouched.

use crate::package::Package;
use crate::template::DOCUMENT_PART;
use crate::xmldom::{Document, Element, Node};

const CUSTOM_PROPS_PART: &str = "docProps/custom.xml";

/// Replace DOCPROPERTY fields referencing custom document properties by
/// their literal values in `word/document.xml` and all
/// `word/header*.xml`/`word/footer*.xml` parts of the given docx. Returns
/// the (possibly unchanged) docx bytes.
pub fn dissolve_custom_property_fields(docx_bytes: &[u8]) -> Result<Vec<u8>, String> {
    let mut pkg = Package::from_bytes(docx_bytes)?;
    let Some(custom_xml) = pkg.get_string(CUSTOM_PROPS_PART) else {
        return Ok(docx_bytes.to_vec());
    };
    let props = custom_properties(&custom_xml);
    if props.is_empty() {
        return Ok(docx_bytes.to_vec());
    }
    let mut parts: Vec<String> = Vec::new();
    if pkg.contains(DOCUMENT_PART) {
        parts.push(DOCUMENT_PART.to_string());
    }
    for e in &pkg.entries {
        if (e.name.starts_with("word/header") || e.name.starts_with("word/footer"))
            && e.name.ends_with(".xml")
        {
            parts.push(e.name.clone());
        }
    }
    for part in parts {
        let Some(xml) = pkg.get_string(&part) else {
            continue;
        };
        if !xml.contains("DOCPROPERTY") {
            continue;
        }
        let enc = pkg.encoding_of(&part);
        if let Some(new_xml) = dissolve_in_part(&xml, &props) {
            pkg.set(&part, crate::package::encode_part_owned(new_xml, &enc));
        }
    }
    pkg.to_bytes()
}

/// Parse docProps/custom.xml into (name, display value) pairs. Values live in
/// a typed child element (`vt:lpwstr`, `vt:i4`, ...); booleans render as
/// "Y"/"N" (docxcompose FieldBase._format_value).
fn custom_properties(xml: &str) -> Vec<(String, String)> {
    let Ok(dom) = Document::parse(xml) else {
        return Vec::new();
    };
    let mut els: Vec<&Element> = Vec::new();
    dom.root.iter_descendants("property", &mut els);
    let mut props = Vec::new();
    for p in els {
        let Some(name) = p.get_attr("name") else {
            continue;
        };
        let Some(value_el) = p.children.iter().find_map(|c| match c {
            Node::Elem(e) => Some(e),
            Node::Text(_) => None,
        }) else {
            continue;
        };
        let text = value_el.text_content();
        let value = if value_el.name == "vt:bool" {
            if text.trim().eq_ignore_ascii_case("true") {
                "Y".to_string()
            } else {
                "N".to_string()
            }
        } else {
            text
        };
        props.push((name.to_string(), value));
    }
    props
}

fn prop_value<'a>(props: &'a [(String, String)], name: &str) -> Option<&'a str> {
    props
        .iter()
        .find(|(n, _)| n.eq_ignore_ascii_case(name))
        .map(|(_, v)| v.as_str())
}

/// Parse the property name out of a DOCPROPERTY instruction
/// (` DOCPROPERTY "name" \* MERGEFORMAT ` and quoting variants). More lenient
/// than docxcompose's regex (which requires a trailing `\* MERGEFORMAT`).
fn docproperty_name(instr: &str) -> Option<String> {
    let idx = instr.find("DOCPROPERTY")?;
    let rest = instr[idx + "DOCPROPERTY".len()..].trim_start();
    let name = if let Some(q) = rest.strip_prefix('"') {
        let end = q.find('"')?;
        &q[..end]
    } else {
        rest.split_whitespace().next()?
    };
    if name.is_empty() {
        None
    } else {
        Some(name.to_string())
    }
}

/// Dissolve the known DOCPROPERTY fields of one part; None when nothing
/// changed (or the part does not parse).
fn dissolve_in_part(xml: &str, props: &[(String, String)]) -> Option<String> {
    let mut dom = Document::parse(xml).ok()?;
    let mut changed = false;
    dissolve_in_element(&mut dom.root, props, &mut changed);
    if changed {
        Some(dom.serialize())
    } else {
        None
    }
}

fn dissolve_in_element(el: &mut Element, props: &[(String, String)], changed: &mut bool) {
    // simple fields: <w:fldSimple w:instr=" DOCPROPERTY ...">runs</w:fldSimple>
    for i in 0..el.children.len() {
        let replacement = match &el.children[i] {
            Node::Elem(e) if e.name == "w:fldSimple" => e
                .get_attr("w:instr")
                .filter(|instr| instr.contains("DOCPROPERTY"))
                .and_then(docproperty_name)
                .and_then(|name| prop_value(props, &name))
                .map(|value| simple_field_replacement(e, value)),
            _ => None,
        };
        if let Some(node) = replacement {
            el.children[i] = node;
            *changed = true;
        }
    }
    // complex fields: begin -> instrText -> separate -> value runs -> end
    dissolve_complex_fields(el, props, changed);
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            dissolve_in_element(e, props, changed);
        }
    }
}

/// docxcompose SimpleField.replace_field_with_value: the field wrapper is
/// dropped, its first run kept; here the run's text is also set to the
/// current property value (a stale cached value is not kept).
fn simple_field_replacement(fld: &Element, value: &str) -> Node {
    let mut run = match fld.children.iter().find_map(|c| match c {
        Node::Elem(e) if e.name == "w:r" => Some(e.clone()),
        _ => None,
    }) {
        Some(r) => r,
        None => Element::new("w:r"),
    };
    set_run_text(&mut run, value);
    Node::Elem(run)
}

/// Set the run's first w:t to `value`, clear any further w:t (a stale cached
/// value split over several text nodes must not leak through), or append a
/// w:t when the run has none.
fn set_run_text(run: &mut Element, value: &str) {
    let mut done = false;
    set_first_wt_and_clear_rest(run, value, &mut done);
    if !done {
        let mut t = Element::new("w:t");
        t.set_attr("xml:space", "preserve");
        t.children.push(Node::Text(value.to_string()));
        run.children.push(Node::Elem(t));
    }
}

fn set_first_wt_and_clear_rest(el: &mut Element, value: &str, done: &mut bool) {
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            if e.name == "w:t" {
                if *done {
                    e.children.clear();
                } else {
                    e.children = vec![Node::Text(value.to_string())];
                    e.set_attr("xml:space", "preserve");
                    *done = true;
                }
            } else {
                set_first_wt_and_clear_rest(e, value, done);
            }
        }
    }
}

fn run_fld_char_type(run: &Element) -> Option<String> {
    let mut out: Vec<&Element> = Vec::new();
    run.iter_descendants("w:fldChar", &mut out);
    out.first()
        .and_then(|e| e.get_attr("w:fldCharType"))
        .map(|s| s.to_string())
}

fn run_instr_text(run: &Element) -> String {
    let mut out: Vec<&Element> = Vec::new();
    run.iter_descendants("w:instrText", &mut out);
    out.iter().map(|e| e.text_content()).collect()
}

fn has_wt(el: &Element) -> bool {
    let mut out: Vec<&Element> = Vec::new();
    el.iter_descendants("w:t", &mut out);
    !out.is_empty()
}

/// Clone a field run without its fldChar (keeps the run properties as a
/// template for the value run, like docxcompose's ComplexField.update).
fn value_run_template(run: &Element) -> Element {
    let mut r = run.clone();
    r.children
        .retain(|c| !matches!(c, Node::Elem(e) if e.name == "w:fldChar" || e.name == "w:instrText"));
    r
}

/// Dissolve complex DOCPROPERTY fields among the direct children of `el`
/// (fields never span paragraphs in practice: begin/instr/separate/end runs
/// are siblings inside one w:p).
fn dissolve_complex_fields(el: &mut Element, props: &[(String, String)], changed: &mut bool) {
    let mut scan_from = 0usize;
    while scan_from < el.children.len() {
        // next field-begin run at/after scan_from
        let mut begin = None;
        for i in scan_from..el.children.len() {
            if let Node::Elem(r) = &el.children[i] {
                if r.name == "w:r" && run_fld_char_type(r).as_deref() == Some("begin") {
                    begin = Some(i);
                    break;
                }
            }
        }
        let Some(b) = begin else { break };
        // scan the field: instruction runs up to separate, value runs, end
        let mut instr = String::new();
        let mut separate = None;
        let mut end = None;
        let mut j = b + 1;
        while j < el.children.len() {
            if let Node::Elem(r) = &el.children[j] {
                if r.name == "w:r" {
                    match run_fld_char_type(r).as_deref() {
                        Some("end") => {
                            end = Some(j);
                            break;
                        }
                        Some("separate") if separate.is_none() => separate = Some(j),
                        // nested field begin: the outer field is malformed
                        Some("begin") => break,
                        _ => {}
                    }
                    if separate.is_none() {
                        instr.push_str(&run_instr_text(r));
                    }
                }
            }
            j += 1;
        }
        let Some(e) = end else { break }; // unterminated field: leave alone
        let value = docproperty_name(&instr).and_then(|n| prop_value(props, &n).map(|v| v.to_string()));
        match value {
            Some(value) => {
                apply_complex_dissolve(el, b, separate, e, &value);
                *changed = true;
                scan_from = b + 1;
            }
            // not a known property field: skip past it
            None => scan_from = e + 1,
        }
    }
}

/// Remove the field machinery of one complex field, keeping a single literal
/// value (docxcompose ComplexField.replace_field_with_value + update).
fn apply_complex_dissolve(
    el: &mut Element,
    begin: usize,
    separate: Option<usize>,
    end: usize,
    value: &str,
) {
    let template = match &el.children[begin] {
        Node::Elem(r) => value_run_template(r),
        _ => Element::new("w:r"),
    };
    match separate {
        Some(s) => {
            let n_value = end - (s + 1);
            // drop the end run, then everything from begin through separate
            el.children.remove(end);
            for _ in begin..=s {
                el.children.remove(begin);
            }
            // the former value runs now start at `begin`
            if n_value == 0 {
                let mut r = template;
                set_run_text(&mut r, value);
                el.children.insert(begin, Node::Elem(r));
            } else {
                if let Node::Elem(r) = &mut el.children[begin] {
                    set_run_text(r, value);
                }
                // additional value runs carrying text would duplicate the value
                let mut remaining = n_value;
                let mut k = 1;
                while k < remaining {
                    let text_run = matches!(&el.children[begin + k], Node::Elem(r) if has_wt(r));
                    if text_run {
                        el.children.remove(begin + k);
                        remaining -= 1;
                    } else {
                        k += 1;
                    }
                }
            }
        }
        None => {
            // no separate: the field has no cached value; replace the whole
            // field by a single value run
            for _ in begin..=end {
                el.children.remove(begin);
            }
            let mut r = template;
            set_run_text(&mut r, value);
            el.children.insert(begin, Node::Elem(r));
        }
    }
}
