//! InlineImage support: build the drawing XML for an image, register the
//! image part and relationships in the current rendering part.

use crate::image::ImageInfo;
use crate::package::{escape_xml_attr, rel_type, relative_target};
use crate::template::TplCore;

const NS_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// next available positive integer id in the part XML (like python-docx next_id)
///
/// Equivalent to scanning with the regex `\bid="(\d+)"`, but implemented as a
/// plain linear string scan: the backtracking regex VM is far too slow on
/// multi-megabyte part XML (it is evaluated once per InlineImage).
fn next_shape_id(part_xml: &str) -> u32 {
    let bytes = part_xml.as_bytes();
    let mut max = 0u32;
    let mut start = 0usize;
    while let Some(rel) = part_xml[start..].find("id=\"") {
        let i = start + rel;
        start = i + 4;
        // emulate the \b before `id`: previous char must be a non-word char
        if i > 0 {
            let prev = bytes[i - 1];
            if prev.is_ascii_alphanumeric() || prev == b'_' {
                continue;
            }
        }
        let ds = i + 4;
        let mut de = ds;
        while de < bytes.len() && bytes[de].is_ascii_digit() {
            de += 1;
        }
        // the regex requires a closing quote right after the digits
        if de > ds && de < bytes.len() && bytes[de] == b'"' {
            if let Ok(n) = part_xml[ds..de].parse::<u32>() {
                max = max.max(n);
            }
        }
    }
    max + 1
}

/// Generate the xml string that InlineImage renders to inside a template,
/// registering the image into `part` of the package.
pub fn inline_image_xml(
    tpl: &mut TplCore,
    part: &str,
    blob: &[u8],
    filename: Option<&str>,
    width: Option<i64>,
    height: Option<i64>,
    anchor: Option<&str>,
    title: Option<&str>,
    descr: Option<&str>,
) -> Result<String, String> {
    let drawing = drawing_xml(tpl, part, blob, filename, width, height, anchor, title, descr)?;
    Ok(format!(
        "</w:t></w:r><w:r>{}</w:r><w:r><w:t xml:space=\"preserve\">",
        drawing
    ))
}

/// Generate the `<w:drawing>...</w:drawing>` xml for an inline image,
/// registering the image part and relationships in `part`.
pub fn drawing_xml(
    tpl: &mut TplCore,
    part: &str,
    blob: &[u8],
    filename: Option<&str>,
    width: Option<i64>,
    height: Option<i64>,
    anchor: Option<&str>,
    title: Option<&str>,
    descr: Option<&str>,
) -> Result<String, String> {
    let info = ImageInfo::parse(blob)?;
    let ext = filename
        .and_then(|f| f.rsplit('.').next().map(|s| s.to_string()))
        .filter(|e| !e.is_empty() && e.len() <= 5 && e.chars().all(|c| c.is_ascii_alphanumeric()))
        .unwrap_or_else(|| info.default_ext.to_string());
    let display_name = filename
        .map(|f| f.to_string())
        .unwrap_or_else(|| format!("image.{}", info.default_ext));

    let (cx, cy) = info.scaled_dimensions(width, height);

    // per-part monotonic shape ids: seed once from the part's current max
    // id, then bump in memory (avoids rescanning the whole part per image)
    let shape_id = match tpl.next_shape_ids.get_mut(part) {
        Some(n) => {
            let id = *n;
            *n += 1;
            id
        }
        None => {
            let pkg = tpl.package.as_ref().ok_or("package not loaded")?;
            let part_xml = pkg.get_string(part).unwrap_or_default();
            let id = next_shape_id(&part_xml);
            tpl.next_shape_ids.insert(part.to_string(), id + 1);
            id
        }
    };

    let pkg = tpl.package.as_mut().ok_or("package not loaded")?;
    let partname = pkg.get_or_add_image(blob, &ext, info.content_type);
    let target = relative_target(part, &partname);
    let rid = pkg.add_rel(part, rel_type::IMAGE, &target, false);

    let hlink = if let Some(url) = anchor {
        let hrid = pkg.add_rel(part, rel_type::HYPERLINK, url, true);
        format!("<a:hlinkClick r:id=\"{}\"/>", hrid)
    } else {
        String::new()
    };

    // title/descr attributes on wp:docPr and pic:cNvPr
    let mut extra_attrs = String::new();
    if let Some(t) = title {
        extra_attrs += &format!(" title=\"{}\"", escape_xml_attr(t));
    }
    if let Some(d) = descr {
        extra_attrs += &format!(" descr=\"{}\"", escape_xml_attr(d));
    }

    let docpr_extra = hlink.clone();
    let cnvpr_extra = hlink;

    let pic = format!(
        "<wp:inline xmlns:wp=\"{NS_WP}\" xmlns:a=\"{NS_A}\" xmlns:pic=\"{NS_PIC}\" xmlns:r=\"{NS_R}\">\
<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>\
<wp:docPr id=\"{shape_id}\" name=\"Picture {shape_id}\"{extra_attrs}>{docpr_extra}</wp:docPr>\
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>\
<a:graphic><a:graphicData uri=\"{NS_PIC}\">\
<pic:pic>\
<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"{fname}\"{extra_attrs}>{cnvpr_extra}</pic:cNvPr><pic:cNvPicPr/></pic:nvPicPr>\
<pic:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>\
<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>\
</pic:pic>\
</a:graphicData></a:graphic></wp:inline>",
        cx = cx,
        cy = cy,
        shape_id = shape_id,
        docpr_extra = docpr_extra,
        extra_attrs = extra_attrs,
        fname = escape_xml_attr(&display_name),
        cnvpr_extra = cnvpr_extra,
        rid = rid,
    );

    Ok(format!("<w:drawing>{}</w:drawing>", pic))
}
