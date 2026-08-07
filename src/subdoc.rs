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
const ENDNOTE_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

/// Produce the block-level xml for a subdocument to be inserted into the
/// master. With `keep_sections` the subdoc's non-empty body-level `w:sectPr`
/// is preserved: the subdoc content becomes its own section (page
/// size/orientation/margins and header/footer references kept).
pub fn subdoc_xml_opts(
    tpl: &mut TplCore,
    subdoc_bytes: &[u8],
    keep_sections: bool,
) -> Result<String, String> {
    Ok(subdoc_xml_info(tpl, subdoc_bytes, keep_sections)?.0)
}

/// Full subdoc merge; returns the body fragment xml and whether a section
/// break was preserved (callers may use it to skip their own page break).
pub(crate) fn subdoc_xml_info(
    tpl: &mut TplCore,
    subdoc_bytes: &[u8],
    keep_sections: bool,
) -> Result<(String, bool), String> {
    let sub_pkg = Package::from_bytes(subdoc_bytes)?;
    let doc_xml = sub_pkg
        .get_string(DOCUMENT_PART)
        .ok_or_else(|| "subdoc has no word/document.xml".to_string())?;

    // extract body children; stash a non-empty body-level sectPr when
    // keep_sections is set, drop it otherwise (docxcompose parity)
    let dom = Document::parse(&doc_xml)?;
    let body = dom
        .root
        .find("w:body")
        .ok_or_else(|| "subdoc has no w:body".to_string())?;
    let mut body_children: Vec<Node> = Vec::new();
    let mut sectpr_xml: Option<String> = None;
    for c in &body.children {
        if let Node::Elem(e) = c {
            if e.name == "w:sectPr" {
                if keep_sections
                    && e.children.iter().any(|ch| matches!(ch, Node::Elem(_)))
                {
                    let mut s = String::new();
                    e.serialize(&mut s);
                    sectpr_xml = Some(s); // last one wins
                }
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

    // the subdoc content becomes its own section: append its sectPr as a
    // paragraph-level sectPr (before remap_rids so header/footer references
    // participate in the rId remap)
    if let Some(sp) = &sectpr_xml {
        body_xml.push_str("<w:p><w:pPr>");
        body_xml.push_str(sp);
        body_xml.push_str("</w:pPr></w:p>");
    }

    // dissolve custom property fields (keep the values)
    body_xml = dissolve_property_fields(body_xml);

    let mut rid_map: HashMap<String, String> = HashMap::new();
    // header/footer parts copied for the preserved sections (master part
    // names); their styles/numbering references are merged together with
    // the body's below
    let mut hf_parts: Vec<String> = Vec::new();

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
            || rel.rel_type == rel_type::COMMENTS
            || rel.rel_type == rel_type::ENDNOTE
        {
            // handled separately below
            continue;
        } else if keep_sections
            && (rel.rel_type == rel_type::HEADER || rel.rel_type == rel_type::FOOTER)
        {
            // referenced from the preserved sectPr: copy with a fresh part
            // name on collision so the master's own header/footer parts
            // stay untouched
            let target_part = resolve_target(DOCUMENT_PART, &rel.target);
            if !sub_pkg.contains(&target_part) {
                continue;
            }
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            let new_part = copy_part_renamed(pkg, &sub_pkg, &target_part);
            let target = crate::package::relative_target(DOCUMENT_PART, &new_part);
            hf_parts.push(new_part);
            pkg.add_rel(DOCUMENT_PART, &rel.rel_type, &target, false)
        } else if !rel.is_external {
            // recursively copy the target part and its own relationships
            let target_part = resolve_target(DOCUMENT_PART, &rel.target);
            if !sub_pkg.contains(&target_part) {
                continue;
            }
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            copy_part_recursive(pkg, &sub_pkg, &target_part, &target_part, 0);
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

    // resume-section break: clone the master's current body sectPr so the
    // content before the subdoc keeps the master's page setup. Inserted
    // AFTER remap_rids: its r:ids are the master's own and must not be
    // touched by the subdoc's rid_map.
    if sectpr_xml.is_some() {
        if let Some(master_sp) = master_resume_sectpr(tpl) {
            body_xml = format!("<w:p><w:pPr>{}</w:pPr></w:p>{}", master_sp, body_xml);
        }
    }

    // renumber bookmarks so ids don't collide with the master document
    body_xml = renumber_bookmarks(tpl, body_xml);

    // prepare the footnotes/comments contents early: their image/external
    // rids are remapped here; the style/numbering references inside them are
    // merged jointly with the body below, so a style/numId referenced from
    // the body and from a footnote/comment maps to the same new id
    let mut fn_content = prepare_note_part_content(tpl, &sub_pkg, "word/footnotes.xml")?;
    let mut en_content = prepare_note_part_content(tpl, &sub_pkg, "word/endnotes.xml")?;
    let mut c_content = if body_xml.contains("w:commentReference") {
        prepare_note_part_content(tpl, &sub_pkg, "word/comments.xml")?
    } else {
        None
    };

    // styles + numbering merge. The copied header/footer parts and the
    // footnotes/comments contents may reference the subdoc's
    // styles/numbering too: merge body and those contents as ONE string so a
    // style/numId referenced from several places maps to the same new id,
    // then split back and write the hf parts out.
    let style_renames;
    let extras = hf_parts.len()
        + fn_content.is_some() as usize
        + en_content.is_some() as usize
        + c_content.is_some() as usize;
    if extras == 0 {
        style_renames = merge_styles(tpl, &sub_pkg, &mut body_xml)?;
        body_xml = merge_numbering(tpl, &sub_pkg, body_xml, &style_renames)?;
    } else {
        let hf_contents: Vec<String> = {
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            hf_parts
                .iter()
                .map(|p| pkg.get_string(p).unwrap_or_default())
                .collect()
        };
        let mut joined = body_xml;
        for c in &hf_contents {
            joined.push_str(HF_PART_SEP);
            joined.push_str(c);
        }
        if let Some(c) = &fn_content {
            joined.push_str(HF_PART_SEP);
            joined.push_str(c);
        }
        if let Some(c) = &en_content {
            joined.push_str(HF_PART_SEP);
            joined.push_str(c);
        }
        if let Some(c) = &c_content {
            joined.push_str(HF_PART_SEP);
            joined.push_str(c);
        }
        style_renames = merge_styles(tpl, &sub_pkg, &mut joined)?;
        joined = merge_numbering(tpl, &sub_pkg, joined, &style_renames)?;
        let mut segs = joined.split(HF_PART_SEP);
        body_xml = segs.next().unwrap_or_default().to_string();
        {
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            for (part, seg) in hf_parts.iter().zip(segs.by_ref()) {
                let enc = pkg.encoding_of(part);
                pkg.set(
                    part,
                    crate::package::encode_part_owned(seg.to_string(), &enc),
                );
            }
        }
        if fn_content.is_some() {
            fn_content = Some(segs.next().unwrap_or_default().to_string());
        }
        if en_content.is_some() {
            en_content = Some(segs.next().unwrap_or_default().to_string());
        }
        if c_content.is_some() {
            c_content = Some(segs.next().unwrap_or_default().to_string());
        }
        for part in &hf_parts {
            tpl.invalidate_part(part);
        }
    }

    // restart numbering of the first list in the subdoc (docxcompose default)
    body_xml = restart_first_numbering(tpl, body_xml)?;

    // footnotes merge
    body_xml = merge_footnotes(tpl, body_xml, fn_content)?;

    // endnotes merge
    body_xml = merge_endnotes(tpl, body_xml, en_content)?;

    // comments merge (+ w15 commentsExtended / w16cid commentsIds)
    let (body_xml, comments_merged) = merge_comments(tpl, body_xml, c_content)?;
    if comments_merged {
        merge_comments_extended(tpl, &sub_pkg);
    }

    Ok((body_xml, sectpr_xml.is_some()))
}

/// Separator between the body, the copied header/footer contents and the
/// footnotes/comments contents during the joint styles/numbering merge
/// (never present in real xml).
const HF_PART_SEP: &str = "\u{1}DTPLHF\u{1}";

/// Serialize the master document's body-level sectPr for the resume-section
/// break; None when it is missing/empty, or when the last block before it
/// already ends a section (paragraph-level sectPr — e.g. from a previous
/// keep_sections merge), where a resume break would create an empty section
/// (a blank page).
fn master_resume_sectpr(tpl: &mut TplCore) -> Option<String> {
    let pkg = tpl.package.as_ref()?;
    let xml = pkg.get_string(DOCUMENT_PART)?;
    let dom = Document::parse(&xml).ok()?;
    let body = dom.root.find("w:body")?;
    let mut sectpr: Option<&Element> = None;
    let mut last_block: Option<&Element> = None;
    for c in &body.children {
        if let Node::Elem(e) = c {
            if e.name == "w:sectPr" {
                sectpr = Some(e);
            } else {
                last_block = Some(e);
            }
        }
    }
    if let Some(p) = last_block {
        if p.name == "w:p"
            && p.find("w:pPr")
                .and_then(|ppr| ppr.find("w:sectPr"))
                .is_some()
        {
            return None;
        }
    }
    let e = sectpr?;
    if !e.children.iter().any(|ch| matches!(ch, Node::Elem(_))) {
        return None;
    }
    let mut s = String::new();
    e.serialize(&mut s);
    Some(s)
}

/// Copy a part into the master, allocating a fresh part name on collision
/// (word/header1.xml -> word/header2.xml). Returns the master part name.
fn copy_part_renamed(master: &mut Package, sub: &Package, part: &str) -> String {
    if !master.contains(part) {
        copy_part_recursive(master, sub, part, part, 0);
        return part.to_string();
    }
    let (dir, file) = match part.rfind('/') {
        Some(i) => (&part[..i + 1], &part[i + 1..]),
        None => ("", part),
    };
    let (stem, ext) = match file.find('.') {
        Some(i) => (&file[..i], &file[i..]),
        None => (file, ""),
    };
    let base = stem.trim_end_matches(|c: char| c.is_ascii_digit());
    for n in 1..1000 {
        let cand = format!("{}{}{}{}", dir, base, n, ext);
        if !master.contains(&cand) {
            copy_part_recursive(master, sub, part, &cand, 0);
            return cand;
        }
    }
    part.to_string() // unreachable in practice
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
pub(crate) fn parse_body_fragment(body_xml: &str) -> Result<Document, String> {
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
/// Recursively copy `src` (in the sub package) to `dst` (in the master),
/// remapping its own relationships. Nested parts keep their names (src == dst
/// in the recursion); only the top-level part may be renamed by
/// [`copy_part_renamed`].
fn copy_part_recursive(master: &mut Package, sub: &Package, src: &str, dst: &str, depth: usize) {
    if depth > 8 || master.contains(dst) {
        return;
    }
    let Some(blob) = sub.get(src) else {
        return;
    };

    // remap the part's own relationships first
    let sub_rels = sub.rels(src);
    let is_xml = src.ends_with(".xml") || src.ends_with(".rels");
    if is_xml {
        let mut xml = String::from_utf8_lossy(blob).into_owned();
        if !sub_rels.rels.is_empty() {
            let mut rid_map: HashMap<String, String> = HashMap::new();
            for rel in &sub_rels.rels {
                let new_rid = if rel.rel_type == rel_type::IMAGE {
                    let target = resolve_target(src, &rel.target);
                    let Some(img) = sub.get(&target) else { continue };
                    let (partname, _) = add_image_part(master, img, &target);
                    let rel_target = crate::package::relative_target(dst, &partname);
                    master.add_rel(dst, rel_type::IMAGE, &rel_target, false)
                } else if rel.is_external {
                    master.add_rel(dst, &rel.rel_type, &rel.target, true)
                } else {
                    let target = resolve_target(src, &rel.target);
                    if sub.contains(&target) {
                        copy_part_recursive(master, sub, &target, &target, depth + 1);
                    }
                    let rel_target = crate::package::relative_target(dst, &target);
                    master.add_rel(dst, &rel.rel_type, &rel_target, false)
                };
                rid_map.insert(rel.id.clone(), new_rid);
            }
            xml = remap_rids(&xml, &rid_map);
        }
        master.set(dst, xml.into_bytes());
    } else {
        // binary parts (embeddings, activeX, …): copy verbatim — a lossy
        // UTF-8 round-trip would replace invalid bytes with U+FFFD and
        // silently corrupt the part
        master.set(dst, blob.to_vec());
    }
    if let Some(ct) = sub_content_type_override(sub, src) {
        master.ensure_content_type_override(dst, &ct);
    }
    // copy the rels file itself if the part had no rels to remap but a rels file exists
    if sub_rels.rels.is_empty() {
        let src_rels = crate::package::rels_path_for(src);
        if let Some(b) = sub.get(&src_rels) {
            let dst_rels = crate::package::rels_path_for(dst);
            if !master.contains(&dst_rels) {
                master.set(&dst_rels, b.to_vec());
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
    let id_re = crate::patch::re(r#"<w:bookmarkStart[^>]* w:id="(\d+)""#);
    // cheap reject before touching the master at all (most subdoc bodies
    // have no bookmarks)
    if !id_re.find(&body_xml).ok().flatten().is_some() {
        return body_xml;
    }
    // the master max id is scanned once, then advanced locally as each
    // subdoc's bookmarks are renumbered (they land in the master in order)
    let base = match tpl.bookmark_next_id {
        Some(b) => b,
        None => {
            let _ = tpl.flush_doc();
            let master_xml = tpl
                .package
                .as_ref()
                .and_then(|p| p.get_string(DOCUMENT_PART))
                .unwrap_or_default();
            let mut max_id: i64 = 0;
            for cap in id_re.captures_iter(&master_xml).flatten() {
                if let Ok(n) = cap[1].parse::<i64>() {
                    max_id = max_id.max(n);
                }
            }
            max_id + 1
        }
    };
    let mut max_assigned = base;
    body_xml = psub(
        r#"(<w:bookmark(?:Start|End)[^>]* w:id=")(\d+)(")"#,
        |m: &fancy_regex::Captures| {
            let n: i64 = m.get(2).unwrap().as_str().parse().unwrap_or(0);
            let new = n + base;
            max_assigned = max_assigned.max(new);
            format!(
                "{}{}{}",
                m.get(1).unwrap().as_str(),
                new,
                m.get(3).unwrap().as_str()
            )
        },
        &body_xml,
    );
    tpl.bookmark_next_id = Some(max_assigned + 1);
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
        tpl.invalidate_part("word/styles.xml");
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
        tpl.invalidate_part("word/styles.xml");
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

/// Produce a subdoc note part's content (footnotes/comments) with its image
/// and external relationships remapped into the master package. Style and
/// numbering references inside the content are remapped later, jointly with
/// the body (see subdoc_xml_info).
fn prepare_note_part_content(
    tpl: &mut TplCore,
    sub: &Package,
    part: &str,
) -> Result<Option<String>, String> {
    let Some(mut xml) = sub.get_string(part) else {
        return Ok(None);
    };
    // remap image relationships inside the part (rare but supported by
    // docxcompose's add_referenced_parts)
    let sub_rels = sub.rels(part);
    let mut rid_map: HashMap<String, String> = HashMap::new();
    for rel in &sub_rels.rels {
        let new_rid = if rel.rel_type == rel_type::IMAGE {
            let target = resolve_target(part, &rel.target);
            let Some(blob) = sub.get(&target) else { continue };
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            let (partname, _) = add_image_part(pkg, blob, &target);
            let rel_target = crate::package::relative_target(part, &partname);
            pkg.add_rel(part, rel_type::IMAGE, &rel_target, false)
        } else if rel.is_external {
            let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
            pkg.add_rel(part, &rel.rel_type, &rel.target, true)
        } else {
            continue;
        };
        rid_map.insert(rel.id.clone(), new_rid);
    }
    xml = remap_rids(&xml, &rid_map);
    Ok(Some(xml))
}

/// Merge footnote definitions from the subdoc; returns body xml with
/// remapped w:footnoteReference ids. `fn_content` is the subdoc's
/// footnotes.xml with image rids and style/numbering references already
/// remapped (None when the subdoc has no footnotes part).
fn merge_footnotes(
    tpl: &mut TplCore,
    mut body_xml: String,
    fn_content: Option<String>,
) -> Result<String, String> {
    let Some(sub_fn_xml) = fn_content else {
        return Ok(body_xml);
    };
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

// ---------------- endnotes ----------------

/// Merge endnote definitions from the subdoc; returns body xml with
/// remapped w:endnoteReference ids (same algorithm as merge_footnotes).
fn merge_endnotes(
    tpl: &mut TplCore,
    mut body_xml: String,
    en_content: Option<String>,
) -> Result<String, String> {
    let Some(sub_en_xml) = en_content else {
        return Ok(body_xml);
    };
    // quick check: are there real endnotes (id > 1)?
    let sub_dom = match Document::parse(&sub_en_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };
    let mut sub_notes: Vec<&Element> = Vec::new();
    collect_endnotes(&sub_dom.root, &mut sub_notes);
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
    if !pkg.contains("word/endnotes.xml") {
        pkg.set("word/endnotes.xml", sub_en_xml.clone().into_bytes());
        pkg.ensure_content_type_override("word/endnotes.xml", ENDNOTE_CT);
        if pkg.rels(DOCUMENT_PART).by_type(rel_type::ENDNOTE).next().is_none() {
            pkg.add_rel(DOCUMENT_PART, rel_type::ENDNOTE, "endnotes.xml", false);
        }
        return Ok(body_xml);
    }

    let master_xml = pkg.get_string("word/endnotes.xml").unwrap_or_default();
    let mut master_dom = match Document::parse(&master_xml) {
        Ok(d) => d,
        Err(_) => return Ok(body_xml),
    };
    let mut master_notes: Vec<&Element> = Vec::new();
    collect_endnotes(&master_dom.root, &mut master_notes);
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
    pkg.set("word/endnotes.xml", xml.into_bytes());

    if !id_map.is_empty() {
        body_xml = psub(
            r#"(<w:endnoteReference w:id=")(\d+)(")"#,
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

fn collect_endnotes<'a>(el: &'a Element, out: &mut Vec<&'a Element>) {
    if el.name == "w:endnote" {
        out.push(el);
        return;
    }
    for c in &el.children {
        if let Node::Elem(e) = c {
            collect_endnotes(e, out);
        }
    }
}

// ---------------- comments ----------------

const COMMENTS_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

/// Merge comment definitions from the subdoc; returns body xml with remapped
/// w:commentRangeStart / w:commentRangeEnd / w:commentReference ids, plus a
/// flag telling whether subdoc comments were merged (gates the
/// commentsExtended/commentsIds merge). `c_content` is the subdoc's
/// comments.xml with image rids and style/numbering references already
/// remapped (None when the subdoc has no comments part or the body has no
/// comment references).
/// (docxcompose does not merge comments at all; this goes beyond parity.)
fn merge_comments(
    tpl: &mut TplCore,
    mut body_xml: String,
    c_content: Option<String>,
) -> Result<(String, bool), String> {
    if !body_xml.contains("w:commentReference") {
        return Ok((body_xml, false));
    }
    let Some(sub_c_xml) = c_content else {
        return Ok((body_xml, false));
    };
    let sub_dom = match Document::parse(&sub_c_xml) {
        Ok(d) => d,
        Err(_) => return Ok((body_xml, false)),
    };
    let mut sub_comments: Vec<&Element> = Vec::new();
    collect_comments(&sub_dom.root, &mut sub_comments);
    if sub_comments.is_empty() {
        return Ok((body_xml, false));
    }

    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    if !pkg.contains("word/comments.xml") {
        pkg.set("word/comments.xml", sub_c_xml.clone().into_bytes());
        pkg.ensure_content_type_override("word/comments.xml", COMMENTS_CT);
        if pkg.rels(DOCUMENT_PART).by_type(rel_type::COMMENTS).next().is_none() {
            pkg.add_rel(DOCUMENT_PART, rel_type::COMMENTS, "comments.xml", false);
        }
        return Ok((body_xml, true));
    }

    let master_xml = pkg.get_string("word/comments.xml").unwrap_or_default();
    let mut master_dom = match Document::parse(&master_xml) {
        Ok(d) => d,
        Err(_) => return Ok((body_xml, false)),
    };
    let mut master_comments: Vec<&Element> = Vec::new();
    collect_comments(&master_dom.root, &mut master_comments);
    let max_id = master_comments
        .iter()
        .filter_map(|e| e.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()))
        .max()
        .unwrap_or(-1);
    let offset = max_id + 1;

    let mut id_map: HashMap<i64, i64> = HashMap::new();
    for c in &sub_comments {
        if let Some(id) = c.get_attr("w:id").and_then(|v| v.parse::<i64>().ok()) {
            let mut n = (*c).clone();
            let new_id = id + offset;
            n.set_attr("w:id", &new_id.to_string());
            id_map.insert(id, new_id);
            master_dom.root.children.push(Node::Elem(n));
        }
    }
    let xml = master_dom.serialize();
    pkg.set("word/comments.xml", xml.into_bytes());

    if !id_map.is_empty() {
        body_xml = psub(
            r#"(<w:comment(?:RangeStart|RangeEnd|Reference) w:id=")(\d+)(")"#,
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
    Ok((body_xml, true))
}

/// Merge the w15 commentsExtended / w16cid commentsIds parts (comment
/// threading/resolved state) from the subdoc. paraId/durableId values are
/// random hex and are kept as-is (cross-document collision ~2^-32).
fn merge_comments_extended(tpl: &mut TplCore, sub: &Package) {
    const PARTS: [(&str, &str, &str, &str); 2] = [
        (
            "word/commentsExtended.xml",
            "application/vnd.ms-word.commentsExtended+xml",
            "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
            "w15:commentEx",
        ),
        (
            "word/commentsIds.xml",
            "application/vnd.ms-word.commentsIds+xml",
            "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
            "w16cid:commentId",
        ),
    ];
    for (part, ct, rt, item_tag) in PARTS {
        let Some(sub_xml) = sub.get_string(part) else {
            continue;
        };
        let pkg = match tpl.package.as_mut() {
            Some(p) => p,
            None => return,
        };
        if !pkg.contains(part) {
            pkg.set(part, sub_xml.into_bytes());
            pkg.ensure_content_type_override(part, ct);
            if pkg.rels(DOCUMENT_PART).by_type(rt).next().is_none() {
                let target = crate::package::relative_target(DOCUMENT_PART, part);
                pkg.add_rel(DOCUMENT_PART, rt, &target, false);
            }
            continue;
        }
        // append the subdoc's items to the master's part
        let master_xml = pkg.get_string(part).unwrap_or_default();
        let Ok(mut master_dom) = Document::parse(&master_xml) else {
            continue;
        };
        let Ok(sub_dom) = Document::parse(&sub_xml) else {
            continue;
        };
        let mut items: Vec<&Element> = Vec::new();
        sub_dom.root.iter_descendants(item_tag, &mut items);
        if items.is_empty() {
            continue;
        }
        for it in items {
            master_dom.root.children.push(Node::Elem(it.clone()));
        }
        pkg.set(part, master_dom.serialize().into_bytes());
    }
}

fn collect_comments<'a>(el: &'a Element, out: &mut Vec<&'a Element>) {
    if el.name == "w:comment" {
        out.push(el);
        return;
    }
    for c in &el.children {
        if let Node::Elem(e) = c {
            collect_comments(e, out);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut cursor = std::io::Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut cursor);
            for (name, data) in entries {
                w.start_file(name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                w.write_all(data).unwrap();
            }
            w.finish().unwrap();
        }
        cursor.into_inner()
    }

    /// Binary parts (embeddings, activeX, …) must be copied verbatim: the old
    /// lossy UTF-8 round-trip replaced invalid bytes with U+FFFD and silently
    /// corrupted them.
    #[test]
    fn test_copy_part_recursive_binary_copied_verbatim() {
        // invalid UTF-8 (0xFF/0xFE/0x80…) with no valid decoding
        let blob: &[u8] = &[0xD0, 0xCF, 0x11, 0xE0, 0xFF, 0xFE, 0x80, 0x00, 0xF5, 0xC3];
        let sub = Package::from_bytes(&build_zip(&[(
            "word/embeddings/oleObject1.bin",
            blob,
        )]))
        .unwrap();
        let mut master = Package::from_bytes(&build_zip(&[])).unwrap();

        copy_part_recursive(&mut master, &sub, "word/embeddings/oleObject1.bin", "word/embeddings/oleObject1.bin", 0);

        assert_eq!(
            master.get("word/embeddings/oleObject1.bin"),
            Some(blob.as_ref())
        );
    }

    /// Comments from a subdoc are merged with offset ids and the body
    /// references (range start/end + reference) are remapped accordingly.
    #[test]
    fn test_merge_comments_offsets_ids() {
        const W: &str = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        let master_doc = format!("<w:document {W}><w:body><w:p/></w:body></w:document>");
        let master_comments = format!(
            "<w:comments {W}><w:comment w:id=\"0\"><w:p><w:r><w:t>master note</w:t></w:r></w:p></w:comment></w:comments>"
        );
        let master_zip = build_zip(&[
            ("word/document.xml", master_doc.as_bytes()),
            ("word/comments.xml", master_comments.as_bytes()),
        ]);
        let mut core = TplCore::new(master_zip);
        core.init_docx(false).unwrap();

        let sub_comments = format!(
            "<w:comments {W}><w:comment w:id=\"0\"><w:p><w:r><w:t>sub note</w:t></w:r></w:p></w:comment></w:comments>"
        );
        let sub = Package::from_bytes(&build_zip(&[(
            "word/comments.xml",
            sub_comments.as_bytes(),
        )]))
        .unwrap();

        let body = "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>x</w:t></w:r>\
                    <w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>"
            .to_string();
        let (out, merged) =
            merge_comments(&mut core, body, sub.get_string("word/comments.xml")).unwrap();
        assert!(merged);

        assert!(out.contains("<w:commentRangeStart w:id=\"1\"/>"));
        assert!(out.contains("<w:commentRangeEnd w:id=\"1\"/>"));
        assert!(out.contains("<w:commentReference w:id=\"1\"/>"));
        let merged = core
            .package
            .as_ref()
            .unwrap()
            .get_string("word/comments.xml")
            .unwrap();
        assert!(merged.contains("master note"));
        assert!(merged.contains("sub note"));
        assert!(merged.contains("w:id=\"1\""));
    }

    /// The resume break is skipped when the master's last block already ends
    /// a section (otherwise each consecutive keep_sections append would add
    /// an empty section = a blank page).
    #[test]
    fn test_master_resume_sectpr_skips_after_section_break() {
        const W: &str = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        let sectpr = "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>";
        // plain trailing paragraph -> resume break offered
        let doc1 = format!("<w:document {W}><w:body><w:p/>{sectpr}</w:body></w:document>");
        let mut core1 = TplCore::new(build_zip(&[("word/document.xml", doc1.as_bytes())]));
        core1.init_docx(false).unwrap();
        assert!(master_resume_sectpr(&mut core1).is_some());
        // trailing paragraph-level sectPr -> skipped
        let doc2 = format!(
            "<w:document {W}><w:body><w:p><w:pPr>{sectpr}</w:pPr></w:p>{sectpr}</w:body></w:document>"
        );
        let mut core2 = TplCore::new(build_zip(&[("word/document.xml", doc2.as_bytes())]));
        core2.init_docx(false).unwrap();
        assert!(master_resume_sectpr(&mut core2).is_none());
        // empty body sectPr -> nothing worth resuming
        let doc3 = format!("<w:document {W}><w:body><w:p/><w:sectPr/></w:body></w:document>");
        let mut core3 = TplCore::new(build_zip(&[("word/document.xml", doc3.as_bytes())]));
        core3.init_docx(false).unwrap();
        assert!(master_resume_sectpr(&mut core3).is_none());
    }

    /// commentsExtended items are appended to the master's part (or the part
    /// is copied whole when the master lacks it).
    #[test]
    fn test_merge_comments_extended_append_and_copy() {
        const W15: &str = "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"";
        let sub_ex = format!(
            "<w15:commentsEx {W15}><w15:commentEx w15:paraId=\"AAAA0001\" w15:done=\"1\"/></w15:commentsEx>"
        );
        let sub = Package::from_bytes(&build_zip(&[(
            "word/commentsExtended.xml",
            sub_ex.as_bytes(),
        )]))
        .unwrap();
        let master_doc = b"<w:document/><w:body/>".to_vec();

        // master lacks the part -> copied whole
        let mut core = TplCore::new(build_zip(&[("word/document.xml", &master_doc)]));
        core.init_docx(false).unwrap();
        merge_comments_extended(&mut core, &sub);
        let pkg = core.package.as_ref().unwrap();
        assert!(
            pkg.get_string("word/commentsExtended.xml")
                .unwrap()
                .contains("AAAA0001")
        );

        // master has the part -> subdoc items appended
        let master_ex = format!(
            "<w15:commentsEx {W15}><w15:commentEx w15:paraId=\"BBBB0002\" w15:done=\"0\"/></w15:commentsEx>"
        );
        let mut core2 = TplCore::new(build_zip(&[
            ("word/document.xml", &master_doc),
            ("word/commentsExtended.xml", master_ex.as_bytes()),
        ]));
        core2.init_docx(false).unwrap();
        merge_comments_extended(&mut core2, &sub);
        let merged = core2
            .package
            .as_ref()
            .unwrap()
            .get_string("word/commentsExtended.xml")
            .unwrap();
        assert!(merged.contains("BBBB0002") && merged.contains("AAAA0001"));
    }
}
