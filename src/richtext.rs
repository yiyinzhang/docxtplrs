//! RichText implementation for styled text in Word documents

use crate::types::DocxTplError;
use crate::xml_utils::escape_xml;
use pyo3::prelude::*;
use pyo3::types::PyDict;


/// Text styling options
#[derive(Debug, Clone, Default)]
pub struct TextStyle {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    pub font: Option<String>,
    pub font_size: Option<i64>, // In half-points
    pub color: Option<String>,  // Hex color without #
    pub highlight: Option<String>,
    pub caps: Option<bool>,
    pub small_caps: Option<bool>,
}

impl TextStyle {
    /// Create a new empty style
    pub fn new() -> Self {
        Self::default()
    }

    /// Apply style settings from a Python dict
    pub fn from_pydict(dict: &Bound<'_, PyDict>) -> PyResult<Self> {
        let mut style = Self::new();

        if let Some(v) = dict.get_item("bold")? {
            style.bold = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("italic")? {
            style.italic = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("underline")? {
            style.underline = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("strike")? {
            style.strike = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("font")? {
            style.font = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("font_size")? {
            style.font_size = Some(v.extract::<i64>()? * 2); // Convert points to half-points
        }
        if let Some(v) = dict.get_item("color")? {
            style.color = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("highlight")? {
            style.highlight = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("caps")? {
            style.caps = Some(v.extract()?);
        }
        if let Some(v) = dict.get_item("small_caps")? {
            style.small_caps = Some(v.extract()?);
        }

        Ok(style)
    }

    /// Convert style to WordprocessingML rPr element
    pub fn to_xml(&self) -> String {
        let mut props = Vec::new();

        if self.bold == Some(true) {
            props.push("<w:b/>".to_string());
        }
        if self.italic == Some(true) {
            props.push("<w:i/>".to_string());
        }
        if self.underline == Some(true) {
            props.push("<w:u w:val=\"single\"/>".to_string());
        }
        if self.strike == Some(true) {
            props.push("<w:strike/>".to_string());
        }
        if let Some(ref font) = self.font {
            // Handle font with region prefix (e.g., "eastAsia:微软雅黑")
            if font.contains(':') {
                let parts: Vec<&str> = font.splitn(2, ':').collect();
                props.push(format!(
                    "<w:rFonts w:{}=\"{}\"/>",
                    parts[0],
                    escape_xml(parts[1])
                ));
            } else {
                props.push(format!("<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>", 
                    escape_xml(font), escape_xml(font)));
            }
        }
        if let Some(size) = self.font_size {
            props.push(format!("<w:sz w:val=\"{}\"/>", size));
            props.push(format!("<w:szCs w:val=\"{}\"/>", size));
        }
        if let Some(ref color) = self.color {
            let color = color.trim_start_matches('#');
            props.push(format!("<w:color w:val=\"{}\"/>", escape_xml(color)));
        }
        if let Some(ref highlight) = self.highlight {
            props.push(format!("<w:highlight w:val=\"{}\"/>", escape_xml(highlight)));
        }
        if self.caps == Some(true) {
            props.push("<w:caps/>".to_string());
        }
        if self.small_caps == Some(true) {
            props.push("<w:smallCaps/>".to_string());
        }

        if props.is_empty() {
            String::new()
        } else {
            format!("<w:rPr>{}</w:rPr>", props.join(""))
        }
    }
}

/// A single text fragment with style
#[derive(Debug, Clone)]
pub struct TextFragment {
    pub text: String,
    pub style: TextStyle,
    pub url_id: Option<String>,
}

impl TextFragment {
    pub fn new(text: impl Into<String>, style: TextStyle) -> Self {
        Self {
            text: text.into(),
            style,
            url_id: None,
        }
    }

    pub fn with_url(mut self, url_id: impl Into<String>) -> Self {
        self.url_id = Some(url_id.into());
        self
    }
}

/// RichText for styled text in Word documents
///
/// This struct allows building complex styled text that can be inserted
/// into Word document templates using the {{r variable }} syntax.
#[pyclass(name = "RichText")]
#[derive(Debug, Clone)]
pub struct RichText {
    fragments: Vec<TextFragment>,
}

#[pymethods]
impl RichText {
    /// Create a new RichText object
    ///
    /// Args:
    ///     text: Initial text content
    ///     **kwargs: Style options (bold, italic, underline, strike, font, 
    ///               font_size, color, highlight, caps, small_caps)
    #[new]
    #[pyo3(signature = (text=None, **kwargs))]
    fn new(text: Option<String>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let style = if let Some(dict) = kwargs {
            TextStyle::from_pydict(dict)?
        } else {
            TextStyle::new()
        };

        let mut rt = Self {
            fragments: Vec::new(),
        };

        if let Some(t) = text {
            rt.add_fragment(t, style);
        }

        Ok(rt)
    }

    /// Add text with optional styling
    ///
    /// Args:
    ///     text: Text to add
    ///     **kwargs: Style options for this text fragment
    #[pyo3(signature = (text, **kwargs))]
    fn add(&mut self, text: String, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let style = if let Some(dict) = kwargs {
            TextStyle::from_pydict(dict)?
        } else {
            TextStyle::new()
        };

        self.add_fragment(text, style);
        Ok(())
    }

    /// Add a hyperlink
    ///
    /// Args:
    ///     text: Link text
    ///     url_id: The URL ID (obtained from DocxTemplate.build_url_id())
    ///     **kwargs: Style options for this link
    #[pyo3(signature = (text, url_id, **kwargs))]
    fn add_link(
        &mut self,
        text: String,
        url_id: String,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let style = if let Some(dict) = kwargs {
            let mut s = TextStyle::from_pydict(dict)?;
            // Links are typically underlined and blue by default
            if s.underline.is_none() {
                s.underline = Some(true);
            }
            if s.color.is_none() {
                s.color = Some("0000FF".to_string());
            }
            s
        } else {
            TextStyle {
                underline: Some(true),
                color: Some("0000FF".to_string()),
                ..Default::default()
            }
        };

        let fragment = TextFragment::new(text, style).with_url(url_id);
        self.fragments.push(fragment);
        Ok(())
    }

    /// Add a newline (line break within paragraph)
    fn add_newline(&mut self) {
        self.fragments.push(TextFragment {
            text: "\n".to_string(),
            style: TextStyle::new(),
            url_id: None,
        });
    }

    /// Add a new paragraph
    fn add_paragraph(&mut self) {
        self.fragments.push(TextFragment {
            text: "\x0C".to_string(), // Form feed as paragraph separator
            style: TextStyle::new(),
            url_id: None,
        });
    }

    /// Add a tab character
    fn add_tab(&mut self) {
        self.fragments.push(TextFragment {
            text: "\t".to_string(),
            style: TextStyle::new(),
            url_id: None,
        });
    }

    /// Add a page break
    fn add_page_break(&mut self) {
        self.fragments.push(TextFragment {
            text: "\x0C".to_string(), // Form feed
            style: TextStyle::new(),
            url_id: None,
        });
    }

    /// Convert RichText to XML string
    pub fn to_xml(&self) -> String {
        let mut runs = Vec::new();

        for fragment in &self.fragments {
            let style_xml = fragment.style.to_xml();

            // Handle special characters
            let text = &fragment.text;
            let parts: Vec<&str> = text.split('\n').collect();

            for (i, part) in parts.iter().enumerate() {
                if i > 0 {
                    runs.push("<w:br/>".to_string());
                }

                // Split by tabs
                let tab_parts: Vec<&str> = part.split('\t').collect();
                for (j, tab_part) in tab_parts.iter().enumerate() {
                    if j > 0 {
                        runs.push("<w:tab/>".to_string());
                    }

                    if !tab_part.is_empty() {
                        let escaped = escape_xml(tab_part);
                        let preserve = if tab_part.starts_with(' ') || tab_part.ends_with(' ') {
                            " xml:space=\"preserve\""
                        } else {
                            ""
                        };

                        let run_content = if let Some(ref url_id) = fragment.url_id {
                            format!(
                                "<w:hyperlink r:id=\"{}\"><w:r>{}<w:t{}>{}</w:t></w:r></w:hyperlink>",
                                escape_xml(url_id),
                                style_xml,
                                preserve,
                                escaped
                            )
                        } else {
                            format!(
                                "<w:r>{}{}<w:t{}>{}</w:t></w:r>",
                                if style_xml.is_empty() {
                                    ""
                                } else {
                                    &style_xml
                                },
                                if style_xml.is_empty() { "" } else { "" },
                                preserve,
                                escaped
                            )
                        };

                        runs.push(run_content);
                    }
                }
            }
        }

        runs.join("")
    }

    fn __str__(&self) -> String {
        self.fragments
            .iter()
            .map(|f| f.text.clone())
            .collect::<String>()
    }

    fn __repr__(&self) -> String {
        format!(
            "RichText({})",
            self.fragments
                .iter()
                .map(|f| format!("{:?}", f.text))
                .collect::<Vec<_>>()
                .join(", ")
        )
    }
}

impl RichText {
    fn add_fragment(&mut self, text: impl Into<String>, style: TextStyle) {
        let fragment = TextFragment::new(text, style);
        self.fragments.push(fragment);
    }

    pub fn is_empty(&self) -> bool {
        self.fragments.is_empty()
            || self
                .fragments
                .iter()
                .all(|f| f.text.trim().is_empty())
    }
}

/// RichTextParagraph for styled paragraphs in Word documents
///
/// This struct allows building complex styled paragraphs that can be inserted
/// into Word document templates using the {{p variable }} syntax.
#[pyclass(name = "RichTextParagraph")]
#[derive(Debug, Clone)]
pub struct RichTextParagraph {
    rich_text: RichText,
    paragraph_style: Option<String>,
    alignment: Option<String>,
}

#[pymethods]
impl RichTextParagraph {
    /// Create a new RichTextParagraph object
    #[new]
    fn new() -> Self {
        Self {
            rich_text: RichText::new(None, None).unwrap(),
            paragraph_style: None,
            alignment: None,
        }
    }

    /// Add RichText to this paragraph
    fn add_rt(&mut self, rt: &RichText) {
        for fragment in &rt.fragments {
            self.rich_text.fragments.push(fragment.clone());
        }
    }

    /// Set paragraph style
    #[setter]
    fn set_style(&mut self, style: String) {
        self.paragraph_style = Some(style);
    }

    /// Set paragraph alignment
    #[setter]
    fn set_alignment(&mut self, alignment: String) {
        self.alignment = Some(alignment);
    }

    /// Convert to XML
    pub fn to_xml(&self) -> String {
        let ppr = if self.paragraph_style.is_some() || self.alignment.is_some() {
            let mut props = Vec::new();
            if let Some(ref style) = self.paragraph_style {
                props.push(format!("<w:pStyle w:val=\"{}\"/>", escape_xml(style)));
            }
            if let Some(ref align) = self.alignment {
                props.push(format!("<w:jc w:val=\"{}\"/>", escape_xml(align)));
            }
            format!("<w:pPr>{}</w:pPr>", props.join(""))
        } else {
            String::new()
        };

        format!("<w:p>{}</w:p>", ppr)
    }
}

/// Listing class for escaped text with formatting
#[pyclass(name = "Listing")]
#[derive(Debug, Clone)]
pub struct Listing {
    text: String,
}

#[pymethods]
impl Listing {
    /// Create a new Listing object
    ///
    /// Args:
    ///     text: The text content with special characters
    #[new]
    fn new(text: String) -> Self {
        Self { text }
    }

    /// Convert to XML
    pub fn to_xml(&self) -> String {
        let escaped = escape_xml(&self.text);
        // Handle newlines and paragraphs
        let parts: Vec<&str> = escaped.split("\\a").collect();
        let mut paragraphs = Vec::new();

        for (i, part) in parts.iter().enumerate() {
            let runs: Vec<String> = part
                .split('\n')
                .enumerate()
                .map(|(j, line)| {
                    if j > 0 {
                        format!("<w:br/>{}", escape_xml(line))
                    } else {
                        escape_xml(line)
                    }
                })
                .collect();

            if i == 0 && parts.len() == 1 {
                // Single paragraph - just return runs
                return format!(
                    "<w:r><w:t xml:space=\"preserve\">{}</w:t></w:r>",
                    runs.join("")
                );
            } else {
                paragraphs.push(format!(
                    "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                    runs.join("")
                ));
            }
        }

        paragraphs.join("")
    }

    fn __str__(&self) -> String {
        self.text.clone()
    }
}

/// Create a RichText shortcut (alias for RichText::new)
#[pyfunction(name = "R")]
#[pyo3(signature = (text=None, **kwargs))]
pub fn r_shortcut(text: Option<String>, kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<RichText> {
    RichText::new(text, kwargs)
}

/// Create a RichTextParagraph shortcut (alias for RichTextParagraph::new)
#[pyfunction(name = "RP")]
pub fn rp_shortcut() -> RichTextParagraph {
    RichTextParagraph::new()
}
