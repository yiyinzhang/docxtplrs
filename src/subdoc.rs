//! Subdoc support: merge a sub-document's body into the master document,
//! remapping relationships (images, hyperlinks, recursive parts), styles
//! (with conflict renaming), numbering and footnotes.

use crate::image::ImageInfo;
use crate::package::{rel_type, resolve_target, Package};
use crate::patch::sub as psub;
use crate::template::{TplCore, DOCUMENT_PART};
use crate::xmldom::{Document, Element, Node};
use std::collections::{HashMap, HashSet};

const STYLES_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const FOOTNOTES_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";

/// Produce the block-level xml for a subdocument to be inserted into the master.
pub fn subdoc_xml(tpl: &mut TplCore, subdoc_bytes: &[u8]) -> Result<String, String> {
    let sub_pkg = Package::from_bytes(subdoc_bytes)?;
    let doc_xml = sub_pkg
        .get_string(DOCUMENT_PART)
        .ok_or_else(|| "subdoc has no word/document.xml".to_string())?;

    // extract body children, drop trailing sectPr
    let dom = Document::parse(&doc_xml)?;
    let body = dom
        .root
        .find("w:body")
        .ok_or_else(|| "subdoc has no w:body".to_string())?;
    let mut body_children: Vec<Node> = Vec::new();
    for c in &body.children {
        if let Node::Elem(e) = c {
            if e.name == "w:sectPr" {
                continue;
            }
        }
        body_children.push(c.clone());
    }
    let mut body_xml = String::new();
    for c in &body_children {
        match c {
            Node::Elem(e) => e.serialize(&mut body_xml),
            Node::Text(t) => body_xml.push_str(t),
        }
    }

    // dissolve custom property fields (keep the values)
    body_xml = dissolve_property_fields(body_xml);

    let mut rid_map: HashMap<String, String> = HashMap::new();

    // remap relationships of the sub document part (recursively for parts)
    let sub_rels = sub_pkg.rels(DOCUMENT_PART);
    for rel in &sub_rels.rels {
        let new_rid = if rel.rel_type == rel_type::IMAGE {
            let target_part = resolve_target(DOCUMENT_PART, &rel.target);
            let Some(blob) = sub_pkg.get(&target_part) else {
                continue;
            };
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            let (partname, _) = add_image_part(pkg, blob, &target_part);
            let target = crate::package::relative_target(DOCUMENT_PART, &partname);
            pkg.add_rel(DOCUMENT_PART, rel_type::IMAGE, &target, false)
        } else if rel.rel_type == rel_type::HYPERLINK {
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            pkg.add_rel(DOCUMENT_PART, rel_type::HYPERLINK, &rel.target, true)
        } else if rel.rel_type == rel_type::STYLES
            || rel.rel_type == rel_type::NUMBERING
            || rel.rel_type == rel_type::FOOTNOTES
        {
            // handled separately below
            continue;
        } else if !rel.is_external {
            // recursively copy the target part and its own relationships
            let target_part = resolve_target(DOCUMENT_PART, &rel.target);
            if !sub_pkg.contains(&target_part) {
                continue;
            }
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            copy_part_recursive(pkg, &sub_pkg, &target_part, 0);
            let target = crate::package::relative_target(DOCUMENT_PART, &target_part);
            pkg.add_rel(DOCUMENT_PART, &rel.rel_type, &target, false)
        } else {
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            pkg.add_rel(DOCUMENT_PART, &rel.rel_type, &rel.target, true)
        };
        rid_map.insert(rel.id.clone(), new_rid);
    }

    // apply rId remapping in the body xml (single pass over the xml,
    // independent of how many ids were remapped)
    body_xml = remap_rids(&body_xml, &rid_map);

    // renumber bookmarks so ids don't collide with the master document
    body_xml = renumber_bookmarks(tpl, body_xml);

    // styles merge (only referenced styles; conflicts are renamed)
    let style_renames = merge_styles(tpl, &sub_pkg, &mut body_xml)?;

    // numbering merge (applies numId + style renames)
    body_xml = merge_numbering(tpl, &sub_pkg, body_xml, &style_renames)?;

    // restart numbering of the first list in the subdoc (docxcompose default)
    body_xml = restart_first_numbering(tpl, body_xml)?;

    // footnotes merge
    body_xml = merge_footnotes(tpl, &sub_pkg, body_xml)?;

    Ok(body_xml)
}

/// Apply an rId remap with a single regex pass over the xml (previously one
/// full scan per remapped id). Attribute values not present in the map are
/// left untouched, which also preserves the old `old != new` skip.
fn remap_rids(xml: &str, map: &HashMap<String, String>) -> String {
    if map.is_empty() {
        return xml.to_string();
    }
    psub(
        r#"((?:r|o):(?:embed|id|link|pict|dm|lo|qs|cs|relid|rel)=")([^"]+)""#,
        |m| {
            let old = m.get(2).unwrap().as_str();
            match map.get(old) {
                Some(new) if new != old => format!("{}{}\"", m.get(1).unwrap().as_str(), new),
                _ => m.get(0).unwrap().as_str().to_string(),
            }
        },
        xml,
    )
}

/// Wrap a body fragment in a root element for DOM parsing.
fn parse_body_fragment(body_xml: &str) -> Result<Document, String> {
    Document::parse(&format!(
        "<w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{}</w:body>",
        body_xml
    ))
}

fn serialize_body_fragment(dom: &Document) -> String {
    let mut out = String::new();
    for c in &dom.root.children {
        match c {
            Node::Elem(e) => e.serialize(&mut out),
            Node::Text(t) => out.push_str(t),
        }
    }
    out
}

/// Restart the numbering of the first list of the subdoc (docxcompose's
/// restart_first_numbering, enabled by default): the first numbered
/// paragraph (per style) gets a new <w:num> with startOverride=1 so the
/// inserted list restarts instead of continuing the master's sequence.
fn restart_first_numbering(tpl: &mut TplCore, body_xml: String) -> Result<String, String> {
    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    if !pkg.contains("word/numbering.xml") {
        return Ok(body_xml);
    }
    let mut dom = match parse_body_fragment(&body_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };

    let styles_xml = pkg.get_string("word/styles.xml").unwrap_or_default();
    let styles_dom = Document::parse(&styles_xml).ok();

    // style lookup table, built once (was: re-collected and deep-cloned
    // from the styles part for every paragraph)
    let mut style_elems: Vec<&Element> = Vec::new();
    if let Some(d) = styles_dom.as_ref() {
        collect_style_elems(&d.root, &mut style_elems);
    }
    let mut style_by_id: HashMap<String, &Element> = HashMap::new();
    for e in style_elems {
        if let Some(id) = e.get_attr("w:styleId") {
            style_by_id.entry(id.to_string()).or_insert(e);
        }
    }

    let mut restarted: HashSet<String> = HashSet::new();
    let mut new_nums: Vec<Element> = Vec::new();
    // paragraphs to retarget: (paragraph index in body, new numId)
    let mut retarget: Vec<(usize, i64)> = Vec::new();

    // next free numId in master numbering (after the merge)
    let numbering_xml = pkg.get_string("word/numbering.xml").unwrap_or_default();
    let numbering_dom = match Document::parse(&numbering_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };
    let mut max_num: i64 = 0;
    let mut max_abs: i64 = -1;
    find_max_numbering(&numbering_dom.root, &mut max_abs, &mut max_num);
    let mut next_num_id = max_num + 1;

    // num / abstractNum lookup tables, built once (first match wins, like
    // the previous per-paragraph linear find)
    let mut num_by_id: HashMap<i64, &Element> = HashMap::new();
    let mut abs_by_id: HashMap<i64, &Element> = HashMap::new();
    {
        let mut nums: Vec<&Element> = Vec::new();
        numbering_dom.root.iter_descendants("w:num", &mut nums);
        for e in nums {
            if let Some(id) = e.get_attr("w:numId").and_then(|v| v.parse::<i64>().ok()) {
                num_by_id.entry(id).or_insert(e);
            }
        }
        let mut abss: Vec<&Element> = Vec::new();
        numbering_dom
            .root
            .iter_descendants("w:abstractNum", &mut abss);
        for e in abss {
            if let Some(id) = e
                .get_attr("w:abstractNumId")
                .and_then(|v| v.parse::<i64>().ok())
            {
                abs_by_id.entry(id).or_insert(e);
            }
        }
    }

    let para_count = dom
        .root
        .children
        .iter()
        .filter(|c| matches!(c, Node::Elem(e) if e.name == "w:p"))
        .count();
    let mut para_idx = 0usize;
    for i in 0..para_count {
        let _ = i;
        // find the para_idx-th w:p
        let Some(p) = nth_direct_para(&dom.root, para_idx) else {
            break;
        };
        para_idx += 1;
        let Some(ppr) = p.find("w:pPr") else { continue };
        let Some(style_id) = ppr
            .find("w:pStyle")
            .and_then(|e| e.get_attr("w:val"))
            .map(|s| s.to_string())
        else {
            continue;
        };
        if restarted.contains(&style_id) {
            continue;
        }
        let Some(style_el) = style_by_id.get(style_id.as_str()) else {
            continue;
        };
        if style_el.find("w:outlineLvl").is_some() {
            continue; // headings are not restarted
        }
        // numId from the paragraph, else from the style
        let num_id = ppr
            .find("w:numPr")
            .and_then(|n| n.find("w:numId"))
            .and_then(|e| e.get_attr("w:val").and_then(|v| v.parse::<i64>().ok()))
            .or_else(|| {
                style_el
                    .find("w:numId")
                    .and_then(|e| e.get_attr("w:val").and_then(|v| v.parse::<i64>().ok()))
            });
        let Some(num_id) = num_id else { continue };

        // locate the num + abstractNum in master numbering
        let Some(num_el) = num_by_id.get(&num_id) else {
            continue;
        };
        let Some(anum_id) = num_el
            .find("w:abstractNumId")
            .and_then(|e| e.get_attr("w:val").and_then(|v| v.parse::<i64>().ok()))
        else {
            continue;
        };
        let Some(anum_el) = abs_by_id.get(&anum_id) else {
            continue;
        };
        // do not restart bullets
        let is_bullet = anum_el.find_all("w:lvl").iter().any(|lvl| {
            lvl.get_attr("w:ilvl") == Some("0")
                && lvl
                    .find("w:numFmt")
                    .and_then(|f| f.get_attr("w:val"))
                    == Some("bullet")
        });
        if is_bullet {
            continue;
        }

        // create the new num with startOverride=1
        let mut new_num = (*num_el).clone();
        new_num.set_attr("w:numId", &next_num_id.to_string());
        let mut lvl_override = Element::new("w:lvlOverride");
        lvl_override.set_attr("w:ilvl", "0");
        let mut start_override = Element::new("w:startOverride");
        start_override.set_attr("w:val", "1");
        lvl_override.children.push(Node::Elem(start_override));
        new_num.children.push(Node::Elem(lvl_override));
        new_nums.push(new_num);
        retarget.push((para_idx - 1, next_num_id));
        next_num_id += 1;
        restarted.insert(style_id);
    }

    if new_nums.is_empty() {
        return Ok(body_xml);
    }

    // apply retargets
    for (pidx, new_id) in retarget {
        if let Some(p) = nth_direct_para_mut(&mut dom.root, pidx) {
            if let Some(ppr) = p.find_mut("w:pPr") {
                if let Some(numpr) = ppr.find_mut("w:numPr") {
                    if let Some(numid) = numpr.find_mut("w:numId") {
                        numid.set_attr("w:val", &new_id.to_string());
                    }
                } else {
                    // insert a numPr after pStyle
                    let mut numpr = Element::new("w:numPr");
                    let mut ilvl = Element::new("w:ilvl");
                    ilvl.set_attr("w:val", "0");
                    let mut numid = Element::new("w:numId");
                    numid.set_attr("w:val", &new_id.to_string());
                    numpr.children.push(Node::Elem(ilvl));
                    numpr.children.push(Node::Elem(numid));
                    let pos = ppr
                        .children
                        .iter()
                        .position(|c| matches!(c, Node::Elem(e) if e.name == "w:pStyle"))
                        .map(|i| i + 1)
                        .unwrap_or(0);
                    ppr.children.insert(pos, Node::Elem(numpr));
                }
            }
        }
    }

    // append new nums to master numbering (after existing w:num elements)
    let mut numbering_dom = numbering_dom;
    let mut insert_pos = numbering_dom.root.children.len();
    for (i, c) in numbering_dom.root.children.iter().enumerate() {
        if matches!(c, Node::Elem(e) if e.name == "w:num") {
            insert_pos = i + 1;
        }
    }
    for n in new_nums {
        numbering_dom
            .root
            .children
            .insert(insert_pos, Node::Elem(n));
        insert_pos += 1;
    }
    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    pkg.set("word/numbering.xml", numbering_dom.serialize().into_bytes());

    Ok(serialize_body_fragment(&dom))
}

fn nth_direct_para(body: &Element, n: usize) -> Option<&Element> {
    body.children
        .iter()
        .filter_map(|c| match c {
            Node::Elem(e) if e.name == "w:p" => Some(e),
            _ => None,
        })
        .nth(n)
}

fn nth_direct_para_mut(body: &mut Element, n: usize) -> Option<&mut Element> {
    body.children
        .iter_mut()
        .filter_map(|c| match c {
            Node::Elem(e) if e.name == "w:p" => Some(e),
            _ => None,
        })
        .nth(n)
}

fn add_image_part(pkg: &mut Package, blob: &[u8], target_part: &str) -> (String, String) {
    let (ext, content_type) = match ImageInfo::parse(blob) {
        Ok(info) => (info.default_ext.to_string(), info.content_type.to_string()),
        Err(_) => {
            let ext = target_part.rsplit('.').next().unwrap_or("png").to_string();
            (ext.clone(), guess_content_type(&ext).to_string())
        }
    };
    let partname = pkg.get_or_add_image(blob, &ext, &content_type);
    (partname, content_type)
}

/// Recursively copy a part and its related parts into the master package,
/// remapping relationship ids inside the copied XML.
fn copy_part_recursive(master: &mut Package, sub: &Package, part: &str, depth: usize) {
    if depth > 8 || master.contains(part) {
        return;
    }
    let Some(blob) = sub.get(part) else {
        return;
    };

    // remap the part's own relationships first
    let sub_rels = sub.rels(part);
    let mut xml = String::from_utf8_lossy(blob).to_string();
    let is_xml = part.ends_with(".xml") || part.ends_with(".rels");
    if !sub_rels.rels.is_empty() && is_xml {
        let mut rid_map: HashMap<String, String> = HashMap::new();
        for rel in &sub_rels.rels {
            let new_rid = if rel.rel_type == rel_type::IMAGE {
                let target = resolve_target(part, &rel.target);
                let Some(img) = sub.get(&target) else { continue };
                let (partname, _) = add_image_part(master, img, &target);
                let rel_target = crate::package::relative_target(part, &partname);
                master.add_rel(part, rel_type::IMAGE, &rel_target, false)
            } else if rel.is_external {
                master.add_rel(part, &rel.rel_type, &rel.target, true)
            } else {
                let target = resolve_target(part, &rel.target);
                if sub.contains(&target) {
                    copy_part_recursive(master, sub, &target, depth + 1);
                }
                let rel_target = crate::package::relative_target(part, &target);
                master.add_rel(part, &rel.rel_type, &rel_target, false)
            };
            rid_map.insert(rel.id.clone(), new_rid);
        }
        xml = remap_rids(&xml, &rid_map);
    }

    master.set(part, xml.into_bytes());
    if let Some(ct) = sub_content_type_override(sub, part) {
        master.ensure_content_type_override(part, &ct);
    }
    // copy the rels file itself if the part had no rels to remap but a rels file exists
    if sub_rels.rels.is_empty() {
        let rels_path = crate::package::rels_path_for(part);
        if let Some(b) = sub.get(&rels_path) {
            if !master.contains(&rels_path) {
                master.set(&rels_path, b.to_vec());
            }
        }
    }
}

/// Dissolve DOCPROPERTY fields, keeping the displayed values.
fn dissolve_property_fields(mut xml: String) -> String {
    // simple fields: <w:fldSimple w:instr=" DOCPROPERTY name \* ...">content</w:fldSimple>
    xml = psub(
        r#"(?s)<w:fldSimple [^>]*w:instr="[^"]*DOCPROPERTY[^"]*"[^>]*>(.*?)</w:fldSimple>"#,
        |m| m.get(1).unwrap().as_str().to_string(),
        &xml,
    );
    // complex fields: begin .. separate .. end sequences containing DOCPROPERTY
    xml = psub(
        r#"(?s)(?:<w:r[ >].*?<w:fldChar w:fldCharType="begin"/>.*?</w:r>)(?:(?!<w:fldChar w:fldCharType="separate"/>).)*?DOCPROPERTY.*?<w:r[ >].*?<w:fldChar w:fldCharType="separate"/>.*?</w:r>(.*?)<w:r[ >].*?<w:fldChar w:fldCharType="end"/>.*?</w:r>"#,
        |m| m.get(1).unwrap().as_str().to_string(),
        &xml,
    );
    xml
}

/// Renumber w:bookmarkStart/w:bookmarkEnd ids past the master's max id.
fn renumber_bookmarks(tpl: &mut TplCore, mut body_xml: String) -> String {
    let _ = tpl.flush_doc();
    let master_xml = tpl
        .package
        .as_ref()
        .and_then(|p| p.get_string(DOCUMENT_PART))
        .unwrap_or_default();
    let id_re = crate::patch::re(r#"<w:bookmarkStart[^>]* w:id="(\d+)""#);
    let mut max_id: i64 = 0;
    for cap in id_re.captures_iter(&master_xml).flatten() {
        if let Ok(n) = cap[1].parse::<i64>() {
            max_id = max_id.max(n);
        }
    }
    if !id_re.find(&body_xml).ok().flatten().is_some() {
        return body_xml;
    }
    let offset = max_id + 1;
    body_xml = psub(
        r#"(<w:bookmark(?:Start|End)[^>]* w:id=")(\d+)(")"#,
        move |m: &fancy_regex::Captures| {
            let n: i64 = m.get(2).unwrap().as_str().parse().unwrap_or(0);
            format!(
                "{}{}{}",
                m.get(1).unwrap().as_str(),
                n + offset,
                m.get(3).unwrap().as_str()
            )
        },
        &body_xml,
    );
    body_xml
}

fn guess_content_type(ext: &str) -> &'static str {
    match ext.to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        _ => "application/octet-stream",
    }
}

fn sub_content_type_override(pkg: &Package, part_name: &str) -> Option<String> {
    let xml = pkg.get_string("[Content_Types].xml")?;
    let pat = format!("PartName=\"/{}\"", part_name);
    let pos = xml.find(&pat)?;
    let after = &xml[pos..];
    let ct_start = after.find("ContentType=\"")? + "ContentType=\"".len();
    let rest = &after[ct_start..];
    let ct_end = rest.find('"')?;
    Some(rest[..ct_end].to_string())
}

// ---------------- styles ----------------

/// Merge styles referenced by the subdoc body. Conflicting style definitions
/// are renamed (like docxcompose) and references updated.
/// Returns the map of renamed style ids (old -> new).
fn merge_styles(
    tpl: &mut TplCore,
    sub: &Package,
    body_xml: &mut String,
) -> Result<HashMap<String, String>, String> {
    let mut renames: HashMap<String, String> = HashMap::new();
    let Some(sub_styles_xml) = sub.get_string("word/styles.xml") else {
        return Ok(renames);
    };
    let sub_dom = match Document::parse(&sub_styles_xml) {
        Ok(d) => d,
        Err(_) => return Ok(renames),
    };

    // collect style ids referenced from the subdoc body
    let mut referenced: HashSet<String> = HashSet::new();
    let style_ref_re = crate::patch::re(r#"w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)""#);
    for cap in style_ref_re.captures_iter(body_xml).flatten() {
        referenced.insert(cap[1].to_string());
    }

    let mut sub_styles: Vec<&Element> = Vec::new();
    collect_style_elems(&sub_dom.root, &mut sub_styles);
    let style_by_id: HashMap<String, &Element> = sub_styles
        .iter()
        .filter_map(|e| e.get_attr("w:styleId").map(|id| (id.to_string(), *e)))
        .collect();

    // expand referenced set with basedOn / link / next chains
    let mut queue: Vec<String> = referenced.iter().cloned().collect();
    while let Some(id) = queue.pop() {
        if let Some(el) = style_by_id.get(&id) {
            for tag in ["w:basedOn", "w:link", "w:next"] {
                if let Some(dep) = el.find(tag) {
                    if let Some(v) = dep.get_attr("w:val") {
                        if referenced.insert(v.to_string()) {
                            queue.push(v.to_string());
                        }
                    }
                }
            }
        }
    }

    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    if !pkg.contains("word/styles.xml") {
        // master has no styles part: copy sub's styles entirely
        pkg.set("word/styles.xml", sub_styles_xml.clone().into_bytes());
        pkg.ensure_content_type_override("word/styles.xml", STYLES_CT);
        if pkg.rels(DOCUMENT_PART).by_type(rel_type::STYLES).next().is_none() {
            pkg.add_rel(DOCUMENT_PART, rel_type::STYLES, "styles.xml", false);
        }
        return Ok(renames);
    }

    let master_xml = pkg.get_string("word/styles.xml").unwrap_or_default();
    let mut master_dom = match Document::parse(&master_xml) {
        Ok(d) => d,
        Err(_) => return Ok(renames),
    };

    let mut master_styles: Vec<&Element> = Vec::new();
    collect_style_elems(&master_dom.root, &mut master_styles);
    let master_by_id: HashMap<String, &Element> = master_styles
        .iter()
        .filter_map(|e| e.get_attr("w:styleId").map(|id| (id.to_string(), *e)))
        .collect();

    // decide renames for conflicting definitions
    let existing_ids: HashSet<String> = master_by_id.keys().cloned().collect();
    for st in &sub_styles {
        let Some(id) = st.get_attr("w:styleId").map(|s| s.to_string()) else {
            continue;
        };
        if !referenced.contains(&id) {
            continue;
        }
        if let Some(&master_st) = master_by_id.get(&id) {
            if !elements_equivalent(st, master_st) {
                // conflict: rename like docxcompose (id_1, id_2, ...)
                let mut n = 1;
                let mut new_id = format!("{}_{}", id, n);
                while existing_ids.contains(&new_id) || style_by_id.contains_key(&new_id) {
                    n += 1;
                    new_id = format!("{}_{}", id, n);
                }
                renames.insert(id, new_id);
            }
        }
    }

    // apply renames to body xml and to merged styles' references (single
    // pass; previously one full scan per renamed style)
    if !renames.is_empty() {
        *body_xml = psub(
            r#"(w:(?:pStyle|rStyle|tblStyle|basedOn|link|next) w:val=")([^"]+)(")"#,
            |m| {
                let old = m.get(2).unwrap().as_str();
                match renames.get(old) {
                    Some(new) => format!(
                        "{}{}{}",
                        m.get(1).unwrap().as_str(),
                        new,
                        m.get(3).unwrap().as_str()
                    ),
                    None => m.get(0).unwrap().as_str().to_string(),
                }
            },
            body_xml,
        );
    }

    // append referenced styles that the master does not have (with renames applied)
    let mut added = 0usize;
    for st in sub_styles {
        let Some(id) = st.get_attr("w:styleId").map(|s| s.to_string()) else {
            continue;
        };
        if !referenced.contains(&id) {
            continue;
        }
        if existing_ids.contains(&id) && !renames.contains_key(&id) {
            continue; // identical definition already present
        }
        let mut st = (*st).clone();
        if let Some(new_id) = renames.get(&id) {
            st.set_attr("w:styleId", new_id);
            // remap internal references
            for tag in ["w:basedOn", "w:link", "w:next"] {
                if let Some(dep) = st.find_mut(tag) {
                    if let Some(v) = dep.get_attr("w:val") {
                        if let Some(nv) = renames.get(v) {
                            let nv = nv.clone();
                            dep.set_attr("w:val", &nv);
                        }
                    }
                }
            }
        }
        master_dom.root.children.push(Node::Elem(st));
        added += 1;
    }
    if added > 0 {
        let xml = master_dom.serialize();
        pkg.set("word/styles.xml", xml.into_bytes());
    }
    Ok(renames)
}

/// Cheap equivalence check for style definitions (attribute-order-insensitive
/// enough for practical purposes: compare sorted child serialization).
fn elements_equivalent(a: &Element, b: &Element) -> bool {
    let mut sa = String::new();
    let mut sb = String::new();
    a.serialize(&mut sa);
    b.serialize(&mut sb);
    sa == sb
}

fn collect_style_elems<'a>(el: &'a Element, out: &mut Vec<&'a Element>) {
    for c in &el.children {
        if let Node::Elem(e) = c {
            if e.name == "w:style" {
                out.push(e);
            }
        }
    }
}

// ---------------- numbering ----------------

/// Merge numbering definitions; returns body xml with remapped numIds.
fn merge_numbering(
    tpl: &mut TplCore,
    sub: &Package,
    mut body_xml: String,
    style_renames: &HashMap<String, String>,
) -> Result<String, String> {
    let Some(sub_num_xml) = sub.get_string("word/numbering.xml") else {
        return Ok(body_xml);
    };
    let sub_dom = match Document::parse(&sub_num_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };

    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    if !pkg.contains("word/numbering.xml") {
        pkg.set("word/numbering.xml", sub_num_xml.clone().into_bytes());
        pkg.ensure_content_type_override("word/numbering.xml", NUMBERING_CT);
        if pkg.rels(DOCUMENT_PART).by_type(rel_type::NUMBERING).next().is_none() {
            pkg.add_rel(DOCUMENT_PART, rel_type::NUMBERING, "numbering.xml", false);
        }
        // numbering picture bullets reference images via numbering.xml.rels
        merge_numbering_pics(pkg, sub);
        return Ok(body_xml);
    }

    // numbering picture bullets
    let pic_rid_map = merge_numbering_pics(pkg, sub);

    let master_xml = pkg.get_string("word/numbering.xml").unwrap_or_default();
    let mut master_dom = match Document::parse(&master_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };

    // find max abstractNumId and numId in master
    let mut max_abstract: i64 = -1;
    let mut max_num: i64 = 0;
    find_max_numbering(&master_dom.root, &mut max_abstract, &mut max_num);
    let abstract_offset = max_abstract + 1;
    let num_offset = max_num + 1;

    // remap sub numbering: build numId map and adjusted elements
    let mut num_map: HashMap<i64, i64> = HashMap::new();
    let mut new_elems: Vec<Node> = Vec::new();
    for c in &sub_dom.root.children {
        let Node::Elem(e) = c else { continue };
        let mut e = e.clone();
        if e.name == "w:abstractNum" {
            if let Some(id) = e.get_attr("w:abstractNumId").and_then(|v| v.parse::<i64>().ok()) {
                e.set_attr("w:abstractNumId", &(id + abstract_offset).to_string());
            }
            apply_numbering_elem_remaps(&mut e, style_renames, &pic_rid_map);
            new_elems.push(Node::Elem(e));
        } else if e.name == "w:num" {
            let old_id = e.get_attr("w:numId").and_then(|v| v.parse::<i64>().ok());
            if let Some(id) = old_id {
                let new_id = id + num_offset;
                num_map.insert(id, new_id);
                e.set_attr("w:numId", &new_id.to_string());
            }
            if let Some(an) = e.find_mut("w:abstractNumId") {
                if let Some(v) = an.get_attr("w:val").and_then(|v| v.parse::<i64>().ok()) {
                    an.set_attr("w:val", &(v + abstract_offset).to_string());
                }
            }
            apply_numbering_elem_remaps(&mut e, style_renames, &pic_rid_map);
            new_elems.push(Node::Elem(e));
        }
    }
    master_dom.root.children.extend(new_elems);
    let xml = master_dom.serialize();
    pkg.set("word/numbering.xml", xml.into_bytes());

    // remap numId references in the sub body xml (single pass)
    if !num_map.is_empty() {
        body_xml = psub(
            r#"(<w:numId w:val=")(\d+)(")"#,
            |m| {
                let old: i64 = m.get(2).unwrap().as_str().parse().unwrap_or(-1);
                match num_map.get(&old) {
                    Some(new) => format!(
                        "{}{}{}",
                        m.get(1).unwrap().as_str(),
                        new,
                        m.get(3).unwrap().as_str()
                    ),
                    None => m.get(0).unwrap().as_str().to_string(),
                }
            },
            &body_xml,
        );
    }

    Ok(body_xml)
}

/// apply style renames (lvl pStyle) and numPicBullet rid remaps inside a
/// numbering element subtree
fn apply_numbering_elem_remaps(
    el: &mut Element,
    style_renames: &HashMap<String, String>,
    pic_rid_map: &HashMap<String, String>,
) {
    if el.name == "w:pStyle" {
        if let Some(v) = el.get_attr("w:val") {
            if let Some(nv) = style_renames.get(v) {
                let nv = nv.clone();
                el.set_attr("w:val", &nv);
            }
        }
    }
    if el.name == "w:numPicBullet" {
        if let Some(v) = el.get_attr("r:id") {
            if let Some(nv) = pic_rid_map.get(v) {
                let nv = nv.clone();
                el.set_attr("r:id", &nv);
            }
        }
    }
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            apply_numbering_elem_remaps(e, style_renames, pic_rid_map);
        }
    }
}

/// Copy numbering picture bullet images and remap their rIds.
/// Returns old rid -> new rid map (within the numbering part's rels).
fn merge_numbering_pics(master: &mut Package, sub: &Package) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let sub_rels = sub.rels("word/numbering.xml");
    for rel in &sub_rels.rels {
        if rel.rel_type != rel_type::IMAGE || rel.is_external {
            continue;
        }
        let target = resolve_target("word/numbering.xml", &rel.target);
        let Some(blob) = sub.get(&target) else { continue };
        let (partname, _) = add_image_part(master, blob, &target);
        let rel_target = crate::package::relative_target("word/numbering.xml", &partname);
        let new_rid = master.add_rel("word/numbering.xml", rel_type::IMAGE, &rel_target, false);
        map.insert(rel.id.clone(), new_rid);
    }
    map
}

fn find_max_numbering(el: &Element, max_abstract: &mut i64, max_num: &mut i64) {
    if el.name == "w:abstractNum" {
        if let Some(v) = el.get_attr("w:abstractNumId").and_then(|v| v.parse::<i64>().ok()) {
            *max_abstract = (*max_abstract).max(v);
        }
    } else if el.name == "w:num" {
        if let Some(v) = el.get_attr("w:numId").and_then(|v| v.parse::<i64>().ok()) {
            *max_num = (*max_num).max(v);
        }
    }
    for c in &el.children {
        if let Node::Elem(e) = c {
            find_max_numbering(e, max_abstract, max_num);
        }
    }
}

// ---------------- footnotes ----------------

/// Merge footnote definitions from the subdoc; returns body xml with
/// remapped w:footnoteReference ids.
fn merge_footnotes(
    tpl: &mut TplCore,
    sub: &Package,
    mut body_xml: String,
) -> Result<String, String> {
    let Some(mut sub_fn_xml) = sub.get_string("word/footnotes.xml") else {
        return Ok(body_xml);
    };
    // remap image relationships inside footnotes (rare but supported by
    // docxcompose's add_referenced_parts)
    {
        let sub_fn_rels = sub.rels("word/footnotes.xml");
        let mut rid_map: HashMap<String, String> = HashMap::new();
        for rel in &sub_fn_rels.rels {
            let new_rid = if rel.rel_type == rel_type::IMAGE {
                let target = resolve_target("word/footnotes.xml", &rel.target);
                let Some(blob) = sub.get(&target) else { continue };
                let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
                let (partname, _) = add_image_part(pkg, blob, &target);
                let rel_target = crate::package::relative_target("word/footnotes.xml", &partname);
                pkg.add_rel("word/footnotes.xml", rel_type::IMAGE, &rel_target, false)
            } else if rel.is_external {
                let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
                pkg.add_rel("word/footnotes.xml", &rel.rel_type, &rel.target, true)
            } else {
                continue;
            };
            rid_map.insert(rel.id.clone(), new_rid);
        }
        sub_fn_xml = remap_rids(&sub_fn_xml, &rid_map);
    }
    // quick check: are there real footnotes (id > 1)?
    let sub_dom = match Document::parse(&sub_fn_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };
    let mut sub_notes: Vec<&Element> = Vec::new();
    collect_footnotes(&sub_dom.root, &mut sub_notes);
    let has_real = sub_notes.iter().any(|e| {
        e.get_attr("w:id")
            .and_then(|v| v.parse::<i64>().ok())
            .map(|id| id > 1)
            .unwrap_or(false)
    });
    if !has_real {
        return Ok(body_xml);
    }

    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    if !pkg.contains("word/footnotes.xml") {
        pkg.set("word/footnotes.xml", sub_fn_xml.clone().into_bytes());
        pkg.ensure_content_type_override("word/footnotes.xml", FOOTNOTES_CT);
        if pkg.rels(DOCUMENT_PART).by_type(rel_type::FOOTNOTES).next().is_none() {
            pkg.add_rel(DOCUMENT_PART, rel_type::FOOTNOTES, "footnotes.xml", false);
        }
        return Ok(body_xml);
    }

    let master_xml = pkg.get_string("word/footnotes.xml").unwrap_or_default();
    let mut master_dom = match Document::parse(&master_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };
    let mut master_notes: Vec<&Element> = Vec::new();
    collect_footnotes(&master_dom.root, &mut master_notes);
    let max_id = master_notes
        .iter()
        .filter_map(|e| e.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()))
        .max()
        .unwrap_or(1);
    let offset = max_id + 1;

    // id remap for real notes
    let mut id_map: HashMap<i64, i64> = HashMap::new();
    for note in &sub_notes {
        if let Some(id) = note.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()) {
            if id > 1 {
                let mut n = (*note).clone();
                let new_id = id + offset;
                n.set_attr("w:id", &new_id.to_string());
                id_map.insert(id, new_id);
                master_dom.root.children.push(Node::Elem(n));
            }
        }
    }
    let xml = master_dom.serialize();
    pkg.set("word/footnotes.xml", xml.into_bytes());

    if !id_map.is_empty() {
        body_xml = psub(
            r#"(<w:footnoteReference w:id=")(\d+)(")"#,
            |m| {
                let old: i64 = m.get(2).unwrap().as_str().parse().unwrap_or(-1);
                match id_map.get(&old) {
                    Some(new) => format!(
                        "{}{}{}",
                        m.get(1).unwrap().as_str(),
                        new,
                        m.get(3).unwrap().as_str()
                    ),
                    None => m.get(0).unwrap().as_str().to_string(),
                }
            },
            &body_xml,
        );
    }
    Ok(body_xml)
}

fn collect_footnotes<'a>(el: &'a Element, out: &mut Vec<&'a Element>) {
    if el.name == "w:footnote" {
        out.push(el);
        return;
    }
    for c in &el.children {
        if let Node::Elem(e) = c {
            collect_footnotes(e, out);
        }
    }
}
