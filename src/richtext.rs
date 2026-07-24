//! Port of docxtpl RichText / RichTextParagraph / Listing XML generation.

/// html escape like Python's html.escape (escapes & < > " ')
pub fn html_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#x27;"),
            _ => out.push(c),
        }
    }
    out
}

#[derive(Debug, Clone, Default)]
pub struct TextProps {
    pub style: Option<String>,
    pub color: Option<String>,
    pub highlight: Option<String>,
    pub size: Option<u32>,
    pub subscript: bool,
    pub superscript: bool,
    pub bold: bool,
    pub italic: bool,
    /// false = no underline; Some string holds underline value ("single" etc.) or bool-ish
    pub underline: Option<String>,
    pub strike: bool,
    pub font: Option<String>,
    pub url_id: Option<String>,
    pub rtl: bool,
    pub lang: Option<String>,
}

pub fn richtext_run(text: &str, p: &TextProps) -> String {
    let text = html_escape(text);
    let mut prop = String::new();

    if let Some(style) = &p.style {
        prop += &format!("<w:rStyle w:val=\"{}\"/>", style);
    }
    if let Some(color) = &p.color {
        let color = color.strip_prefix('#').unwrap_or(color);
        prop += &format!("<w:color w:val=\"{}\"/>", color);
    }
    if let Some(highlight) = &p.highlight {
        let highlight = highlight.strip_prefix('#').unwrap_or(highlight);
        prop += &format!("<w:shd w:fill=\"{}\"/>", highlight);
    }
    if let Some(size) = p.size {
        prop += &format!("<w:sz w:val=\"{}\"/>", size);
        prop += &format!("<w:szCs w:val=\"{}\"/>", size);
    }
    if p.subscript {
        prop += "<w:vertAlign w:val=\"subscript\"/>";
    }
    if p.superscript {
        prop += "<w:vertAlign w:val=\"superscript\"/>";
    }
    if p.bold {
        prop += "<w:b/>";
        if p.rtl {
            prop += "<w:bCs/>";
        }
    }
    if p.italic {
        prop += "<w:i/>";
        if p.rtl {
            prop += "<w:iCs/>";
        }
    }
    if let Some(underline) = &p.underline {
        let valid = [
            "single", "double", "thick", "dotted", "dash", "dotDash", "dotDotDash", "wave",
        ];
        let u = if valid.contains(&underline.as_str()) {
            underline.as_str()
        } else {
            "single"
        };
        prop += &format!("<w:u w:val=\"{}\"/>", u);
    }
    if p.strike {
        prop += "<w:strike/>";
    }
    if let Some(font) = &p.font {
        let mut regional_font = String::new();
        let mut font_name = font.as_str();
        if let Some((region, f)) = font.split_once(':') {
            font_name = f;
            regional_font = format!(" w:{}=\"{}\"", region, f);
        }
        prop += &format!(
            "<w:rFonts w:ascii=\"{f}\" w:hAnsi=\"{f}\" w:cs=\"{f}\"{regional}/>",
            f = font_name,
            regional = regional_font
        );
    }
    if p.rtl {
        prop += "<w:rtl w:val=\"true\"/>";
    }
    if let Some(lang) = &p.lang {
        prop += &format!("<w:lang w:val=\"{}\"/>", lang);
    }

    let mut xml = String::from("<w:r>");
    if !prop.is_empty() {
        xml += &format!("<w:rPr>{}</w:rPr>", prop);
    }
    xml += &format!("<w:t xml:space=\"preserve\">{}</w:t></w:r>", text);
    if let Some(url_id) = &p.url_id {
        xml = format!(
            "<w:hyperlink r:id=\"{}\" w:tgtFrame=\"_blank\">{}</w:hyperlink>",
            url_id, xml
        );
    }
    xml
}

pub fn richtext_paragraph(runs_xml: &str, parastyle: Option<&str>) -> String {
    let mut prop = String::new();
    if let Some(ps) = parastyle {
        prop += &format!("<w:pStyle w:val=\"{}\"/>", ps);
    }
    let mut xml = String::from("<w:p>");
    if !prop.is_empty() {
        xml += &format!("<w:pPr>{}</w:pPr>", prop);
    }
    xml += runs_xml;
    xml += "</w:p>";
    xml
}

pub fn listing_xml(text: &str) -> String {
    html_escape(text)
}
