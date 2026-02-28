//! Sub-document handling for Word documents

use crate::types::{DocxTplError, Result};
use pyo3::prelude::*;
use std::fs;
use std::path::Path;

/// Subdoc for embedding documents within documents
///
/// A subdoc represents content that can be inserted into another document.
/// It can be created from scratch or loaded from an existing .docx file.
#[pyclass(name = "Subdoc")]
#[derive(Debug, Clone)]
pub struct Subdoc {
    pub content: String,           // XML content
    pub relationship_id: Option<String>,
    pub is_external: bool,         // True if loaded from external file
    pub source_path: Option<String>,
}

#[pymethods]
impl Subdoc {
    /// Create a new empty Subdoc
    #[new]
    pub fn new() -> Self {
        Self {
            content: String::new(),
            relationship_id: None,
            is_external: false,
            source_path: None,
        }
    }

    fn __repr__(&self) -> String {
        format!(
            "Subdoc(external={}, content_length={})",
            self.is_external,
            self.content.len()
        )
    }
}

impl Subdoc {
    /// Create a Subdoc from an existing .docx file
    pub fn from_file(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Err(DocxTplError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("Subdoc file not found: {}", path.display()),
            )));
        }

        let data = fs::read(path)?;
        let content = Self::extract_document_xml(&data)?;

        Ok(Self {
            content,
            relationship_id: None,
            is_external: true,
            source_path: Some(path.to_string_lossy().to_string()),
        })
    }

    /// Extract document.xml content from a .docx file
    fn extract_document_xml(data: &[u8]) -> Result<String> {
        use std::io::Cursor;
        use zip::ZipArchive;

        let reader = Cursor::new(data);
        let mut archive = ZipArchive::new(reader)?;

        let mut doc_xml = String::new();
        {
            let mut file = archive.by_name("word/document.xml")?;
            std::io::Read::read_to_string(&mut file, &mut doc_xml)?;
        }

        // Extract body content only (remove <?xml declaration and <w:document> wrapper)
        let body_start = doc_xml.find("<w:body>");
        let body_end = doc_xml.rfind("</w:body>");

        if let (Some(start), Some(end)) = (body_start, body_end) {
            Ok(doc_xml[start + 8..end].to_string())
        } else {
            // Try without namespace prefix
            let body_start = doc_xml.find("<body>");
            let body_end = doc_xml.rfind("</body>");

            if let (Some(start), Some(end)) = (body_start, body_end) {
                Ok(doc_xml[start + 6..end].to_string())
            } else {
                // Return full content if body tags not found
                Ok(doc_xml)
            }
        }
    }

    /// Get the XML content to insert
    pub fn to_xml(&self) -> String {
        if self.content.is_empty() {
            return "<w:p><w:r><w:t></w:t></w:r></w:p>".to_string();
        }

        // Wrap the content in paragraphs if needed
        if !self.content.trim().starts_with("<w:p") {
            format!("<w:p><w:r><w:t>{}</w:t></w:r></w:p>", self.content)
        } else {
            self.content.clone()
        }
    }
}

/// Cell color specification for table cells
#[pyclass(name = "CellColor")]
#[derive(Debug, Clone)]
pub struct CellColor {
    pub color: String,
}

#[pymethods]
impl CellColor {
    #[new]
    fn new(color: String) -> Self {
        // Remove # if present
        let color = color.trim_start_matches('#').to_string();
        Self { color }
    }

    /// Generate the XML for cell shading
    pub fn to_xml(&self) -> String {
        format!(
            "<w:tcPr><w:shd w:fill=\"{}\" w:val=\"clear\" w:color=\"auto\"/></w:tcPr>",
            self.color
        )
    }
}

/// Vertical merge specification for table cells
#[pyclass(name = "VerticalMerge")]
#[derive(Debug, Clone)]
pub struct VerticalMerge {
    pub merge: bool,
}

#[pymethods]
impl VerticalMerge {
    #[new]
    fn new(merge: bool) -> Self {
        Self { merge }
    }

    pub fn to_xml(&self) -> String {
        if self.merge {
            "<w:tcPr><w:vMerge/></w:tcPr>".to_string()
        } else {
            "<w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr>".to_string()
        }
    }
}

/// Column span specification for table cells
#[pyclass(name = "ColSpan")]
#[derive(Debug, Clone)]
pub struct ColSpan {
    pub span: u32,
}

#[pymethods]
impl ColSpan {
    #[new]
    fn new(span: u32) -> Self {
        Self { span }
    }

    pub fn to_xml(&self) -> String {
        format!("<w:tcPr><w:gridSpan w:val=\"{}\"/></w:tcPr>", self.span)
    }
}

/// Document builder for creating subdocs programmatically
///
/// This allows building Word document content using a fluent API.
#[pyclass(name = "DocumentBuilder")]
#[derive(Debug, Clone)]
pub struct DocumentBuilder {
    content: Vec<String>,
    current_paragraph: Option<String>,
}

#[pymethods]
impl DocumentBuilder {
    /// Create a new DocumentBuilder
    #[new]
    fn new() -> Self {
        Self {
            content: Vec::new(),
            current_paragraph: None,
        }
    }

    /// Add a paragraph
    fn add_paragraph(&mut self, text: &str) -> PyResult<()> {
        self.finish_current_paragraph();

        let escaped = text
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;");

        self.content.push(format!(
            "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
            escaped
        ));

        Ok(())
    }

    /// Add a heading
    fn add_heading(&mut self, text: &str, level: u8) -> PyResult<()> {
        self.finish_current_paragraph();

        let style = format!("Heading{}", level.min(9).max(1));
        let escaped = text
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;");

        self.content.push(format!(
            "<w:p><w:pPr><w:pStyle w:val=\"{}\"/></w:pPr><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
            style, escaped
        ));

        Ok(())
    }

    /// Add text to current paragraph
    fn add_run(&mut self, text: &str, bold: bool, italic: bool) -> PyResult<()> {
        let escaped = text
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;");

        let mut rpr = String::new();
        if bold {
            rpr.push_str("<w:b/>");
        }
        if italic {
            rpr.push_str("<w:i/>");
        }

        if self.current_paragraph.is_none() {
            self.current_paragraph = Some(format!(
                "<w:p>{}<w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
                if rpr.is_empty() {
                    String::new()
                } else {
                    format!("<w:rPr>{}</w:rPr>", rpr)
                },
                escaped
            ));
        } else {
            // Append to current paragraph
            if let Some(ref mut para) = self.current_paragraph {
                let run = format!(
                    "<w:r>{}<w:t xml:space=\"preserve\">{}</w:t></w:r>",
                    if rpr.is_empty() {
                        String::new()
                    } else {
                        format!("<w:rPr>{}</w:rPr>", rpr)
                    },
                    escaped
                );
                // Insert run before closing </w:p>
                if let Some(pos) = para.rfind("</w:p>") {
                    para.insert_str(pos, &run);
                }
            }
        }

        Ok(())
    }

    /// Build and return the subdoc
    fn build(&mut self) -> PyResult<Subdoc> {
        self.finish_current_paragraph();

        Ok(Subdoc {
            content: self.content.join(""),
            relationship_id: None,
            is_external: false,
            source_path: None,
        })
    }

    fn __repr__(&self) -> String {
        format!(
            "DocumentBuilder(paragraphs={}, current={})",
            self.content.len(),
            self.current_paragraph.is_some()
        )
    }
}

impl DocumentBuilder {
    fn finish_current_paragraph(&mut self) {
        if let Some(para) = self.current_paragraph.take() {
            self.content.push(para);
        }
    }
}
