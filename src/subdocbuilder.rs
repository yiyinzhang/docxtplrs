//! Bound Subdoc support: build sub-document content programmatically
//! (add_paragraph / add_heading / add_picture / add_table), serialized
//! lazily at render time.

use crate::richtext::{richtext_run, TextProps};
use crate::template::TplCore;

#[derive(Debug, Clone, Default)]
pub struct SubRun {
    pub text: String,
    pub props: TextProps,
}

#[derive(Debug, Clone)]
pub enum Block {
    Paragraph {
        style: Option<String>,
        runs: Vec<SubRun>,
    },
    Picture {
        blob: Vec<u8>,
        filename: Option<String>,
        width: Option<i64>,
        height: Option<i64>,
    },
    Table {
        rows: Vec<Vec<String>>,
    },
}

/// Serialize blocks to body-level xml for insertion into the master document.
pub fn serialize_blocks(tpl: &mut TplCore, part: &str, blocks: &[Block]) -> Result<String, String> {
    let mut out = String::new();
    for b in blocks {
        match b {
            Block::Paragraph { style, runs } => {
                out.push_str("<w:p>");
                if let Some(st) = style {
                    let sid = resolve_style_id(tpl, st);
                    out.push_str(&format!(
                        "<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>",
                        sid
                    ));
                }
                for r in runs {
                    let mut props = r.props.clone();
                    if let Some(st) = &props.style {
                        props.style = Some(resolve_style_id(tpl, st));
                    }
                    out.push_str(&richtext_run(&r.text, &props));
                }
                out.push_str("</w:p>");
            }
            Block::Picture {
                blob,
                filename,
                width,
                height,
            } => {
                let drawing = crate::inline_image::drawing_xml(
                    tpl,
                    part,
                    blob,
                    filename.as_deref(),
                    *width,
                    *height,
                    None,
                    None,
                    None,
                )?;
                out.push_str(&format!("<w:p><w:r>{}</w:r></w:p>", drawing));
            }
            Block::Table { rows } => {
                // python-docx distributes the usable page width evenly
                out.push_str(&table_xml(rows, usable_width_twips(tpl)));
            }
        }
    }
    Ok(out)
}

/// Serialize a table block like python-docx's add_table.
pub fn table_xml(rows: &[Vec<String>], usable_twips: i64) -> String {
    let ncols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    let col_w = usable_twips / ncols.max(1) as i64;
    let mut out = String::from("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>");
    out.push_str(
        "<w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>\
         <w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>\
         <w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>\
         <w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>\
         <w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>\
         <w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders></w:tblPr>",
    );
    out.push_str("<w:tblGrid>");
    for _ in 0..ncols {
        out.push_str(&format!("<w:gridCol w:w=\"{}\"/>", col_w));
    }
    out.push_str("</w:tblGrid>");
    for row in rows {
        out.push_str("<w:tr>");
        for cell in row {
            out.push_str("<w:tc><w:tcPr><w:tcW w:w=\"0\" w:type=\"auto\"/></w:tcPr>");
            out.push_str("<w:p><w:r><w:t xml:space=\"preserve\">");
            out.push_str(&crate::richtext::html_escape(cell));
            out.push_str("</w:t></w:r></w:p></w:tc>");
        }
        out.push_str("</w:tr>");
    }
    out.push_str("</w:tbl>");
    out
}

/// Usable page width in twips of the *master* document's last section
/// (python-docx add_table on a Document uses its own section width).
pub fn master_usable_width_twips(tpl: &mut TplCore) -> i64 {
    let _ = tpl.flush_doc();
    let Some(pkg) = tpl.package.as_ref() else {
        return 8640;
    };
    let Some(xml) = pkg.get_string(crate::template::DOCUMENT_PART) else {
        return 8640;
    };
    let Ok(dom) = crate::xmldom::Document::parse(&xml) else {
        return 8640;
    };
    let mut sectprs: Vec<&crate::xmldom::Element> = Vec::new();
    dom.root.iter_descendants("w:sectPr", &mut sectprs);
    let Some(sp) = sectprs.last() else {
        return 8640;
    };
    let w = sp
        .find("w:pgSz")
        .and_then(|e| e.get_attr("w:w"))
        .and_then(|v| v.parse::<i64>().ok());
    let (left, right) = match sp.find("w:pgMar") {
        Some(m) => (
            m.get_attr("w:left").and_then(|v| v.parse::<i64>().ok()).unwrap_or(1800),
            m.get_attr("w:right").and_then(|v| v.parse::<i64>().ok()).unwrap_or(1800),
        ),
        None => (1800, 1800),
    };
    match w {
        Some(w) => (w - left - right).max(1),
        None => 8640,
    }
}

/// Usable page width in twips of python-docx's *default* template
/// (Letter, 1" margins: 12240 - 2*1800 = 8640). docxtpl's bound Subdoc
/// builds content on top of the python-docx default document, so its
/// tables are sized against that template, not the master.
fn usable_width_twips(_tpl: &mut TplCore) -> i64 {
    8640
}

/// Resolve a style name or id to a style id, consulting the master styles part.
pub fn resolve_style_id(tpl: &mut TplCore, style: &str) -> String {
    let Ok(dom) = tpl.part_dom("word/styles.xml") else {
        return style.to_string();
    };
    // exact styleId match first, then by w:name (case-insensitive,
    // python-docx uses name lookup); single walk over the cached DOM
    let want = style.to_lowercase();
    let mut by_name: Option<String> = None;
    let mut stack: Vec<&crate::xmldom::Element> = vec![&dom.root];
    while let Some(el) = stack.pop() {
        // push in reverse so traversal follows document order
        for c in el.children.iter().rev() {
            if let crate::xmldom::Node::Elem(e) = c {
                if e.name == "w:style" {
                    if e.get_attr("w:styleId") == Some(style) {
                        return style.to_string();
                    }
                    if by_name.is_none() {
                        if let Some(v) = e
                            .find("w:name")
                            .and_then(|n| n.get_attr("w:val"))
                        {
                            if v.to_lowercase() == want {
                                by_name = e.get_attr("w:styleId").map(|s| s.to_string());
                            }
                        }
                    }
                }
                stack.push(e);
            }
        }
    }
    by_name.unwrap_or_else(|| style.to_string())
}
