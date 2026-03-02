//! Core DocxTemplate implementation

use crate::image::InlineImage;
use crate::jinja_env::JinjaEnv;
use crate::richtext::{Listing, RichText};
use crate::subdoc::{CellColor, ColSpan, Subdoc, VerticalMerge};
use crate::types::{DocxTplError, Result};
use crate::xml_utils::{
    convert_unicode_attributes, escape_xml, extract_template_variables,
    postprocess_xml_content, preprocess_xml_content,
};
use minijinja::Value;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use regex::Regex;
use std::collections::HashMap;
use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

/// DocxTemplate for rendering Word documents with Jinja2 templates
///
/// This is the main class for working with Word document templates.
/// It loads a .docx file as a template and allows you to render it
/// with context variables using Jinja2 syntax.
#[pyclass(name = "DocxTemplate")]
#[derive(Debug)]
pub struct DocxTemplate {
    template_path: PathBuf,
    xml_parts: HashMap<String, String>, // path -> content
    binary_parts: HashMap<String, Vec<u8>>, // path -> binary data (images, etc.)
    content_types: HashMap<String, String>, // extension -> content type
    relationships: HashMap<String, Vec<(String, String, String)>>, // part -> [(id, type, target)]
    next_rel_id: u32,
    images: HashMap<String, InlineImage>, // rel_id -> image
    image_counter: u32,
    media_replacements: HashMap<String, String>, // old_name -> new_path
    embedded_replacements: HashMap<String, String>, // old_name -> new_path
    undeclared_variables: Option<Vec<String>>,
}

#[pymethods]
impl DocxTemplate {
    /// Load a .docx template file
    ///
    /// Args:
    ///     template_path: Path to the .docx template file
    #[new]
    fn new(template_path: String) -> PyResult<Self> {
        let path = Path::new(&template_path);

        if !path.exists() {
            return Err(DocxTplError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("Template file not found: {}", template_path),
            ))
            .into());
        }

        let data = fs::read(path)?;
        let mut tpl = Self {
            template_path: path.to_path_buf(),
            xml_parts: HashMap::new(),
            binary_parts: HashMap::new(),
            content_types: HashMap::new(),
            relationships: HashMap::new(),
            next_rel_id: 1,
            images: HashMap::new(),
            image_counter: 1,
            media_replacements: HashMap::new(),
            embedded_replacements: HashMap::new(),
            undeclared_variables: None,
        };

        tpl.load_docx(&data)?;

        Ok(tpl)
    }

    /// Render the template with context variables
    ///
    /// Args:
    ///     context: Dictionary of variables for Jinja2 templating
    ///     jinja_env: Optional custom Jinja2 environment with custom filters
    ///     autoescape: Whether to autoescape special characters (default: False)
    ///
    /// Example:
    ///     from docxtplrs import DocxTemplate, JinjaEnv
    ///
    ///     def format_currency(value):
    ///         return f"${value:,.2f}"
    ///
    ///     env = JinjaEnv()
    ///     env.add_filter("currency", format_currency)
    ///
    ///     doc = DocxTemplate("template.docx")
    ///     doc.render(context, jinja_env=env)
    #[pyo3(signature = (context, jinja_env=None, autoescape=false))]
    fn render(
        &mut self,
        context: &Bound<'_, PyDict>,
        jinja_env: Option<PyRef<'_, JinjaEnv>>,
        autoescape: bool,
    ) -> PyResult<()> {
        // Convert Python context to HashMap
        let context_map = self.py_dict_to_context(context)?;

        // Convert JinjaEnv filters to an Arc so we can share them with the closure
        let filters: Arc<HashMap<String, Arc<PyObject>>> = if let Some(je) = jinja_env {
            je.get_filters_arc()
        } else {
            Arc::new(HashMap::new())
        };

        // Process each XML part
        let part_keys: Vec<String> = self.xml_parts.keys().cloned().collect();

        for part_path in part_keys {
            if !part_path.starts_with("word/") {
                continue;
            }
            
            // Skip non-XML files
            if !part_path.ends_with(".xml") && !part_path.ends_with(".rels") {
                continue;
            }

            let content = self.xml_parts.get(&part_path).unwrap().clone();

            // Preprocess XML
            let preprocessed = preprocess_xml_content(&content);

            // Process special tags ({%p %}, {%tr %}, etc.) before rendering
            let preprocessed = self.process_special_tags(&preprocessed)?;

            // Convert Unicode attribute access to bracket notation (minijinja limitation)
            let preprocessed = convert_unicode_attributes(&preprocessed);
            
            // Debug: Save preprocessed template for comparison
            if part_path == "word/document.xml" {
                std::fs::write("/tmp/rust_template.txt", &preprocessed).ok();
            }
            
            // Render with Jinja2
            let rendered =
                self.render_template(&preprocessed, &context_map, autoescape, filters.clone())?;

            // Postprocess
            let postprocessed = postprocess_xml_content(&rendered);

            // Update the part
            self.xml_parts.insert(part_path, postprocessed);
        }

        // Apply media replacements
        self.apply_media_replacements()?;

        Ok(())
    }

    /// Save the rendered document to a file
    ///
    /// Args:
    ///     output_path: Path where the generated document will be saved
    fn save(&self, output_path: String) -> PyResult<()> {
        let path = Path::new(&output_path);

        // Ensure parent directory exists
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }

        let mut buffer = Cursor::new(Vec::new());
        self.write_to_zip(&mut buffer)?;

        fs::write(path, buffer.into_inner())?;

        Ok(())
    }

    /// Get the set of undeclared template variables
    ///
    /// Args:
    ///     context: Optional context dict. If provided, returns variables not in context.
    ///
    /// Returns:
    ///     Set of variable names that need to be defined
    #[pyo3(signature = (context=None))]
    fn get_undeclared_template_variables(
        &self,
        context: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<Vec<String>> {
        let mut all_vars = std::collections::HashSet::new();

        // Collect variables from all document parts
        for (path, content) in &self.xml_parts {
            if path.starts_with("word/document")
                || path.starts_with("word/header")
                || path.starts_with("word/footer")
            {
                let vars = extract_template_variables(content);
                all_vars.extend(vars);
            }
        }

        // If context provided, filter out defined variables
        if let Some(ctx) = context {
            let keys_list: Vec<pyo3::Bound<'_, pyo3::PyAny>> = ctx.keys().into_iter().collect();
            let defined: std::collections::HashSet<String> = keys_list
                .into_iter()
                .filter_map(|k| k.extract::<String>().ok())
                .collect();
            let undeclared: Vec<String> = all_vars
                .into_iter()
                .filter(|v| !defined.contains(v))
                .collect();
            Ok(undeclared)
        } else {
            let vars: Vec<String> = all_vars.into_iter().collect();
            Ok(vars)
        }
    }

    /// Create a new subdoc for embedding
    #[pyo3(signature = (docx_path=None))]
    fn new_subdoc(&self, docx_path: Option<String>) -> PyResult<Subdoc> {
        match docx_path {
            Some(path) => {
                let subdoc = Subdoc::from_file(Path::new(&path))
                    .map_err(|e| DocxTplError::Other(e.to_string()))?;
                Ok(subdoc)
            }
            None => Ok(Subdoc::new()),
        }
    }

    /// Build URL ID for hyperlinks
    ///
    /// Args:
    ///     url: The URL to link to
    ///
    /// Returns:
    ///     A relationship ID for use with RichText.add_link()
    fn build_url_id(&mut self, url: String) -> String {
        let rel_id = format!("rId{}", self.next_rel_id);
        self.next_rel_id += 1;

        // Add to relationships
        let part_rels = self
            .relationships
            .entry("word/_rels/document.xml.rels".to_string())
            .or_insert_with(Vec::new);
        part_rels.push((
            rel_id.clone(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                .to_string(),
            url,
        ));

        rel_id
    }

    /// Replace a picture in the document
    ///
    /// Args:
    ///     dummy_pic_name: Name of the dummy picture in the template
    ///     new_pic_path: Path to the replacement picture
    fn replace_pic(&mut self, dummy_pic_name: String, new_pic_path: String) -> PyResult<()> {
        self.media_replacements
            .insert(dummy_pic_name, new_pic_path);
        Ok(())
    }

    /// Replace media in the document
    ///
    /// Similar to replace_pic but for any media type.
    fn replace_media(&mut self, dummy_media: String, new_media: String) -> PyResult<()> {
        self.media_replacements.insert(dummy_media, new_media);
        Ok(())
    }

    /// Replace embedded objects in the document
    ///
    /// Args:
    ///     dummy_embedded: Name of the dummy embedded object
    ///     new_embedded: Path to the replacement embedded object
    fn replace_embedded(&mut self, dummy_embedded: String, new_embedded: String) -> PyResult<()> {
        self.embedded_replacements
            .insert(dummy_embedded, new_embedded);
        Ok(())
    }

    /// Replace content by zip name
    ///
    /// Args:
    ///     zip_name: The internal zip path to replace
    ///     new_file: Path to the replacement file
    fn replace_zipname(&mut self, zip_name: String, new_file: String) -> PyResult<()> {
        // Read the new file
        let data = fs::read(&new_file)?;
        // Store in a special key for later processing
        self.xml_parts.insert(
            format!("__REPLACE__{}", zip_name),
            String::from_utf8_lossy(&data).to_string(),
        );
        Ok(())
    }

    /// Reset all replacements (for multiple renderings)
    fn reset_replacements(&mut self) {
        self.media_replacements.clear();
        self.embedded_replacements.clear();
    }

    /// Set updateFields to true in settings.xml
    ///
    /// This enables automatic update of fields (like table of contents, page numbers)
    /// when the document is opened.
    fn set_updatefields_true(&mut self) -> PyResult<()> {
        let settings_path = "word/settings.xml";
        
        if let Some(settings_content) = self.xml_parts.get_mut(settings_path) {
            // Check if updateFields already exists
            if !settings_content.contains("updateFields") {
                // Add updateFields element before closing </w:settings> tag
                let update_field_elem = r#"<w:updateFields w:val="true"/>"#;
                
                // Find the last </w:settings> tag and insert before it with proper formatting
                if let Some(pos) = settings_content.rfind("</w:settings>") {
                    settings_content.insert_str(pos, &format!("\n    {}\n", update_field_elem));
                } else {
                    // Try without namespace prefix
                    if let Some(pos) = settings_content.rfind("</settings>") {
                        settings_content.insert_str(pos, &format!("\n    {}\n", update_field_elem));
                    } else {
                        return Err(DocxTplError::Other(
                            "Could not find </w:settings> tag in settings.xml".to_string()
                        ).into());
                    }
                }
            }
            Ok(())
        } else {
            // Settings.xml doesn't exist, create it
            let settings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:updateFields w:val="true"/>
</w:settings>"#;
            self.xml_parts.insert(settings_path.to_string(), settings_xml.to_string());
            
            // Add content type for settings.xml
            self.content_types.insert(
                "/word/settings.xml".to_string(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml".to_string()
            );
            
            // Add relationship for settings.xml
            let doc_rels = self.relationships
                .entry("word/_rels/document.xml.rels".to_string())
                .or_insert_with(Vec::new);
            
            // Check if settings relationship already exists
            let has_settings_rel = doc_rels.iter().any(|(_, rel_type, _)| {
                rel_type.contains("officeDocument/2006/relationships/settings")
            });
            
            if !has_settings_rel {
                let rel_id = format!("rId{}", self.next_rel_id);
                self.next_rel_id += 1;
                doc_rels.push((
                    rel_id,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings".to_string(),
                    "settings.xml".to_string(),
                ));
            }
            
            Ok(())
        }
    }

    /// Get document core properties (metadata)
    ///
    /// Returns a dictionary with properties like:
    /// - author/creator
    /// - title
    /// - subject
    /// - keywords
    /// - description
    /// - last_modified_by
    /// - revision
    fn get_docx_properties(&self) -> PyResult<HashMap<String, String>> {
        let mut props = HashMap::new();
        
        if let Some(core_content) = self.xml_parts.get("docProps/core.xml") {
            // Extract creator/author
            let re = Regex::new(r"<dc:creator>([^<]+)</dc:creator>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("author".to_string(), cap[1].to_string());
                props.insert("creator".to_string(), cap[1].to_string());
            }
            
            // Extract title
            let re = Regex::new(r"<dc:title>([^<]+)</dc:title>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("title".to_string(), cap[1].to_string());
            }
            
            // Extract subject
            let re = Regex::new(r"<dc:subject>([^<]+)</dc:subject>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("subject".to_string(), cap[1].to_string());
            }
            
            // Extract description
            let re = Regex::new(r"<dc:description>([^<]+)</dc:description>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("description".to_string(), cap[1].to_string());
            }
            
            // Extract keywords
            let re = Regex::new(r"<cp:keywords>([^<]+)</cp:keywords>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("keywords".to_string(), cap[1].to_string());
            }
            
            // Extract last modified by
            let re = Regex::new(r"<cp:lastModifiedBy>([^<]+)</cp:lastModifiedBy>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("last_modified_by".to_string(), cap[1].to_string());
            }
            
            // Extract revision
            let re = Regex::new(r"<cp:revision>([^<]+)</cp:revision>").unwrap();
            if let Some(cap) = re.captures(core_content) {
                props.insert("revision".to_string(), cap[1].to_string());
            }
        }
        
        Ok(props)
    }

    /// Set document core properties (metadata)
    ///
    /// Args:
    ///     properties: Dictionary with properties to set:
    ///         - author/creator
    ///         - title
    ///         - subject
    ///         - keywords
    ///         - description
    ///         - last_modified_by
    ///         - revision
    #[pyo3(signature = (properties))]
    fn set_docx_properties(&mut self, properties: &Bound<'_, PyDict>) -> PyResult<()> {
        let core_path = "docProps/core.xml";
        
        // Get existing content or create new
        let mut core_content = if let Some(content) = self.xml_parts.get(core_path) {
            content.clone()
        } else {
            // Create new core.xml
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
</cp:coreProperties>"#.to_string()
        };
        
        // Helper function to update or insert element
        let mut update_element = |tag: &str, ns: &str, value: &str| {
            let pattern = format!(r"<{}:{}>[^<]*</{}:{}>", ns, tag, ns, tag);
            let replacement = format!("<{}:{}>{}</{}:{}>", ns, tag, value, ns, tag);
            let re = Regex::new(&pattern).unwrap();
            
            if re.is_match(&core_content) {
                // Update existing
                core_content = re.replace(&core_content, &replacement).to_string();
            } else {
                // Insert before closing tag
                let close_tag = "</cp:coreProperties>";
                if let Some(pos) = core_content.rfind(close_tag) {
                    core_content.insert_str(pos, &format!("<{}:{}>{}</{}:{}>", ns, tag, value, ns, tag));
                }
            }
        };
        
        Python::with_gil(|py| {
            for (key, value) in properties {
                let key_str: String = key.extract()?;
                let value_str: String = value.extract()?;
                
                match key_str.as_str() {
                    "author" | "creator" => update_element("creator", "dc", &value_str),
                    "title" => update_element("title", "dc", &value_str),
                    "subject" => update_element("subject", "dc", &value_str),
                    "description" => update_element("description", "dc", &value_str),
                    "keywords" => update_element("keywords", "cp", &value_str),
                    "last_modified_by" => update_element("lastModifiedBy", "cp", &value_str),
                    "revision" => update_element("revision", "cp", &value_str),
                    _ => {}
                }
            }
            Ok::<(), PyErr>(())
        })?;
        
        self.xml_parts.insert(core_path.to_string(), core_content);
        
        // Ensure content type is set
        self.content_types.insert(
            "/docProps/core.xml".to_string(),
            "application/vnd.openxmlformats-package.core-properties+xml".to_string()
        );
        
        Ok(())
    }

    /// Modify paragraph properties in the document
    ///
    /// Args:
    ///     paragraph_index: Index of the paragraph to modify (0-based)
    ///     style_id: Optional style ID to apply (e.g., "Heading1", "Normal")
    ///     alignment: Optional alignment ("left", "center", "right", "justify")
    ///     space_before: Optional space before paragraph (in twips)
    ///     space_after: Optional space after paragraph (in twips)
    #[pyo3(signature = (paragraph_index, style_id=None, alignment=None, space_before=None, space_after=None))]
    fn set_paragraph_properties(
        &mut self,
        paragraph_index: usize,
        style_id: Option<String>,
        alignment: Option<String>,
        space_before: Option<i32>,
        space_after: Option<i32>,
    ) -> PyResult<()> {
        let doc_path = "word/document.xml";
        
        // First, collect the information we need
        let modification = if let Some(doc_content) = self.xml_parts.get(doc_path) {
            // Find all paragraphs
            let re = Regex::new(r"<w:p[>\s][^>]*>").unwrap();
            let paragraphs: Vec<_> = re.find_iter(doc_content).collect();
            
            if paragraph_index >= paragraphs.len() {
                return Err(DocxTplError::InvalidArgument(
                    format!("Paragraph index {} out of range (max: {})", paragraph_index, paragraphs.len().saturating_sub(1))
                ).into());
            }
            
            let p_start = paragraphs[paragraph_index].start();
            let p_match_len = paragraphs[paragraph_index].as_str().len();
            
            // Find the paragraph content
            if let Some(p_close) = find_tag_end(doc_content, p_start) {
                let paragraph = &doc_content[p_start..p_close];
                
                // Check if pPr exists
                let has_ppr = paragraph.contains("<w:pPr>");
                
                let mut props_to_add = Vec::new();
                
                // Build pPr content
                if let Some(style) = style_id {
                    props_to_add.push(format!(r#"<w:pStyle w:val="{}"/>"#, style));
                }
                
                if let Some(align) = alignment {
                    let align_val = match align.as_str() {
                        "left" => "left",
                        "center" => "center",
                        "right" => "right",
                        "justify" => "both",
                        "both" => "both",
                        _ => &align,
                    };
                    props_to_add.push(format!(r#"<w:jc w:val="{}"/>"#, align_val));
                }
                
                if space_before.is_some() || space_after.is_some() {
                    let mut spacing = String::from("<w:spacing");
                    if let Some(before) = space_before {
                        spacing.push_str(&format!(r#" w:before="{}""#, before));
                    }
                    if let Some(after) = space_after {
                        spacing.push_str(&format!(r#" w:after="{}""#, after));
                    }
                    spacing.push_str("/>");
                    props_to_add.push(spacing);
                }
                
                if !props_to_add.is_empty() {
                    let ppr_content = format!("<w:pPr>{}</w:pPr>", props_to_add.join(""));
                    
                    if has_ppr {
                        // Replace existing pPr
                        let ppr_re = Regex::new(r"<w:pPr>.*?</w:pPr>").unwrap();
                        let new_paragraph = ppr_re.replace(paragraph, &ppr_content);
                        Some((p_start, p_close, new_paragraph.to_string()))
                    } else {
                        // Insert pPr after <w:p>
                        let insert_pos = p_start + p_match_len;
                        Some((insert_pos, insert_pos, ppr_content))
                    }
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };
        
        // Apply the modification
        if let Some((start, end, new_content)) = modification {
            if let Some(doc_content) = self.xml_parts.get_mut(doc_path) {
                if start == end {
                    doc_content.insert_str(start, &new_content);
                } else {
                    doc_content.replace_range(start..end, &new_content);
                }
            }
        }
        
        Ok(())
    }

    /// Get a preview of the document XML (for debugging)
    fn get_xml(&self) -> PyResult<String> {
        self.xml_parts
            .get("word/document.xml")
            .cloned()
            .ok_or_else(|| DocxTplError::Other("Document XML not found".to_string()).into())
    }

    fn __repr__(&self) -> String {
        format!(
            "DocxTemplate(path='{}', parts={})",
            self.template_path.display(),
            self.xml_parts.len()
        )
    }
}

impl DocxTemplate {
    /// Load a .docx file into memory
    fn load_docx(&mut self, data: &[u8]) -> Result<()> {
        let reader = Cursor::new(data);
        let mut archive = ZipArchive::new(reader)?;

        // Read [Content_Types].xml
        let mut content_types = String::new();
        {
            let mut file = archive.by_name("[Content_Types].xml")?;
            file.read_to_string(&mut content_types)?;
        }
        self.parse_content_types(&content_types)?;

        // Read all XML parts
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();

            // Skip binary files for now
            if name.ends_with(".xml") || name.ends_with(".rels") {
                let mut content = String::new();
                file.read_to_string(&mut content)?;
                self.xml_parts.insert(name, content);
            } else if name.starts_with("word/media/") || name.starts_with("word/embeddings/") {
                // Store binary files separately to avoid corruption
                let mut data = Vec::new();
                file.read_to_end(&mut data)?;
                self.binary_parts.insert(name, data);
            }
        }

        // Parse relationships
        let paths: Vec<String> = self.xml_parts.keys().cloned().collect();
        for path in paths {
            if path.ends_with(".rels") {
                if let Some(content) = self.xml_parts.get(&path) {
                    let content = content.clone();
                    self.parse_relationships(&path, &content)?;
                }
            }
        }

        Ok(())
    }

    /// Parse [Content_Types].xml
    fn parse_content_types(&mut self, content: &str) -> Result<()> {
        // Extract content types for extensions
        let re = Regex::new(r#"Extension="([^"]+)"[^>]*ContentType="([^"]+)""#)?;
        for caps in re.captures_iter(content) {
            self.content_types
                .insert(caps[1].to_string(), caps[2].to_string());
        }

        // Also handle PartName pattern
        let re2 = Regex::new(r#"PartName="([^"]+)"[^>]*ContentType="([^"]+)""#)?;
        for caps in re2.captures_iter(content) {
            self.content_types
                .insert(caps[1].to_string(), caps[2].to_string());
        }

        Ok(())
    }

    /// Parse relationships XML
    fn parse_relationships(&mut self, path: &str, content: &str) -> Result<()> {
        let mut rels = Vec::new();

        let re = Regex::new(
            r#"<Relationship[^>]*Id="([^"]+)"[^>]*Type="([^"]+)"[^>]*Target="([^"]+)""#,
        )?;
        for caps in re.captures_iter(content) {
            rels.push((caps[1].to_string(), caps[2].to_string(), caps[3].to_string()));

            // Track highest relationship ID for new IDs
            if let Some(num) = caps[1].strip_prefix("rId") {
                if let Ok(n) = num.parse::<u32>() {
                    if n >= self.next_rel_id {
                        self.next_rel_id = n + 1;
                    }
                }
            }
        }

        self.relationships.insert(path.to_string(), rels);
        Ok(())
    }

    /// Convert Python dict to context HashMap
    fn py_dict_to_context(&self, dict: &Bound<'_, PyDict>) -> PyResult<HashMap<String, Value>> {
        let mut context = HashMap::new();

        for (key, value) in dict {
            let key_str: String = key.extract()?;
            let val = self.py_to_value(&value)?;
            context.insert(key_str, val);
        }

        Ok(context)
    }

    /// Convert Python object to minijinja Value
    fn py_to_value(&self, obj: &Bound<'_, PyAny>) -> PyResult<Value> {
        // Handle None
        if obj.is_none() {
            return Ok(Value::from(()));
        }

        // Handle RichText
        if let Ok(rt) = obj.extract::<PyRef<RichText>>() {
            return Ok(Value::from_safe_string(rt.to_xml()));
        }

        // Handle InlineImage
        if let Ok(img) = obj.extract::<PyRef<InlineImage>>() {
            let img_clone = img.clone();
            return Ok(Value::from_safe_string(format!(
                "__INLINE_IMAGE__{}__",
                img_clone.image_path
            )));
        }

        // Handle Subdoc
        if let Ok(subdoc) = obj.extract::<PyRef<Subdoc>>() {
            return Ok(Value::from_safe_string(subdoc.to_xml()));
        }

        // Handle CellColor
        if let Ok(cc) = obj.extract::<PyRef<CellColor>>() {
            return Ok(Value::from_safe_string(cc.to_xml()));
        }

        // Handle ColSpan
        if let Ok(cs) = obj.extract::<PyRef<ColSpan>>() {
            return Ok(Value::from_safe_string(cs.to_xml()));
        }

        // Handle VerticalMerge
        if let Ok(vm) = obj.extract::<PyRef<VerticalMerge>>() {
            return Ok(Value::from_safe_string(vm.to_xml()));
        }

        // Handle Listing
        if let Ok(listing) = obj.extract::<PyRef<Listing>>() {
            return Ok(Value::from_safe_string(listing.to_xml()));
        }

        // Handle strings
        if let Ok(s) = obj.extract::<String>() {
            return Ok(Value::from(s));
        }

        // Handle integers
        if let Ok(i) = obj.extract::<i64>() {
            return Ok(Value::from(i));
        }

        // Handle floats
        if let Ok(f) = obj.extract::<f64>() {
            return Ok(Value::from(f));
        }

        // Handle booleans
        if let Ok(b) = obj.extract::<bool>() {
            return Ok(Value::from(b));
        }

        // Handle lists
        if let Ok(list) = obj.downcast::<PyList>() {
            let mut vec = Vec::new();
            for item in list {
                vec.push(self.py_to_value(&item)?);
            }
            return Ok(Value::from(vec));
        }

        // Handle dicts
        if let Ok(dict) = obj.downcast::<PyDict>() {
            let mut map = HashMap::new();
            for (k, v) in dict {
                let key: String = k.extract()?;
                map.insert(key, self.py_to_value(&v)?);
            }
            return Ok(Value::from(map));
        }

        // Default to string representation
        let s = obj.str()?.to_string_lossy().to_string();
        Ok(Value::from(s))
    }

    /// Render template with Jinja2
    fn render_template(
        &mut self,
        template: &str,
        context: &HashMap<String, Value>,
        _autoescape: bool,
        filters: Arc<HashMap<String, Arc<PyObject>>>,
    ) -> Result<String> {
        let mut env = minijinja::Environment::new();

        // IMPORTANT: Disable autoescape before adding templates
        // This prevents minijinja from escaping XML tags
        env.set_auto_escape_callback(|_| minijinja::AutoEscape::None);

        // Add built-in filters
        env.add_filter("e", |s: String| escape_xml(&s));
        env.add_filter("escape", |s: String| escape_xml(&s));

        // Add custom filters from JinjaEnv
        // Clone the filters Arc for each filter to avoid lifetime issues
        let filters_for_closure = Arc::clone(&filters);
        if !filters_for_closure.is_empty() {
            for (name, func) in filters_for_closure.iter() {
                let filter_name = name.clone();
                let func = Arc::clone(func);

                // Create a wrapper function that calls the Python filter
                env.add_filter(
                    filter_name.clone(),
                    move |value: minijinja::Value| -> minijinja::Value {
                        Python::with_gil(|py| {
                            // Convert value to Python
                            let py_value =
                                Self::minijinja_to_python(py, &value).unwrap_or_else(|_| py.None());

                            // Call the Python filter function
                            let result = func.call1(py, (py_value,));

                            // Convert result back to minijinja Value
                            match result {
                                Ok(obj) => Self::python_to_minijinja(py, &obj)
                                    .unwrap_or_else(|_| minijinja::Value::from(())),
                                Err(e) => {
                                    // Log error and return original value
                                    eprintln!("Filter '{}' error: {}", filter_name, e);
                                    value
                                }
                            }
                        })
                    },
                );
            }
        }

        // Add template
        if let Err(e) = env.add_template("doc", template) {
            return Err(e.into());
        }

        // Render
        let tmpl = env.get_template("doc")?;
        let mut result = tmpl.render(context)?;

        // Unescape XML entities that minijinja might have escaped in loop iterations
        // This is necessary to preserve XML tag structure
        result = result.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", "\"").replace("&#39;", "'");

        // Handle inline images
        let result = self.process_inline_images(&result)?;

        Ok(result)
    }

    /// Convert minijinja Value to Python object
    fn minijinja_to_python(py: Python, value: &minijinja::Value) -> PyResult<PyObject> {
        use minijinja::value::ValueKind;

        match value.kind() {
            ValueKind::Undefined | ValueKind::None => Ok(py.None()),
            ValueKind::String => Ok(value
                .as_str()
                .map(|s| s.to_string().into_py(py))
                .unwrap_or_else(|| py.None())),
            ValueKind::Number => {
                // Try integer first, then float
                if let Some(i) = value.as_i64() {
                    Ok(i.into_py(py))
                } else {
                    // Try to get as f64 via string conversion
                    let s = value.to_string();
                    if let Ok(f) = s.parse::<f64>() {
                        Ok(f.into_py(py))
                    } else {
                        Ok(s.into_py(py))
                    }
                }
            }
            ValueKind::Bool => {
                // Boolean check via string representation for now
                let s = value.to_string();
                Ok(s.parse::<bool>().unwrap_or(false).into_py(py))
            }
            ValueKind::Seq => {
                // Convert sequence to list
                if let Some(obj) = value.as_object() {
                    if let Some(iter) = obj.try_iter() {
                        let items: Vec<PyObject> = iter
                            .filter_map(|item| Self::minijinja_to_python(py, &item).ok())
                            .collect();
                        return Ok(items.into_py(py));
                    }
                }
                Ok(py.None())
            }
            ValueKind::Map => {
                // Convert to dict - simplified version
                // For now, just convert to string representation
                // This could be enhanced to properly support map iteration
                let s = value.to_string();
                Ok(s.into_py(py))
            }
            _ => Ok(value.to_string().into_py(py)),
        }
    }

    /// Convert Python object to minijinja Value
    fn python_to_minijinja(py: Python, obj: &PyObject) -> PyResult<minijinja::Value> {
        let bound = obj.bind(py);

        // Handle None
        if bound.is_none() {
            return Ok(minijinja::Value::from(()));
        }

        // Handle strings
        if let Ok(s) = bound.extract::<String>() {
            return Ok(minijinja::Value::from(s));
        }

        // Handle integers
        if let Ok(i) = bound.extract::<i64>() {
            return Ok(minijinja::Value::from(i));
        }

        // Handle floats
        if let Ok(f) = bound.extract::<f64>() {
            return Ok(minijinja::Value::from(f));
        }

        // Handle booleans (must be before integers since bool is a subtype)
        if let Ok(b) = bound.extract::<bool>() {
            return Ok(minijinja::Value::from(b));
        }

        // Handle lists/tuples
        if let Ok(list) = bound.downcast::<PyList>() {
            let mut vec = Vec::new();
            for item in list {
                let py_obj: PyObject = item.into_py(py);
                vec.push(Self::python_to_minijinja(py, &py_obj)?);
            }
            return Ok(minijinja::Value::from(vec));
        }

        // Handle tuples
        if let Ok(tuple) = bound.downcast::<PyTuple>() {
            let mut vec = Vec::new();
            for item in tuple {
                let py_obj: PyObject = item.into_py(py);
                vec.push(Self::python_to_minijinja(py, &py_obj)?);
            }
            return Ok(minijinja::Value::from(vec));
        }

        // Handle dicts
        if let Ok(dict) = bound.downcast::<PyDict>() {
            let mut map = std::collections::HashMap::new();
            for (k, v) in dict {
                let key: String = k.extract()?;
                let py_obj: PyObject = v.into_py(py);
                let val = Self::python_to_minijinja(py, &py_obj)?;
                map.insert(key, val);
            }
            return Ok(minijinja::Value::from(map));
        }

        // Default: convert to string
        let s = bound.str()?.to_string_lossy().to_string();
        Ok(minijinja::Value::from(s))
    }

    /// Process special Jinja2 tags ({%p %}, {%tr %}, etc.)
    fn process_special_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Process paragraph tags {%p if ... %}, {%p for ... %}, etc.
        result = self.process_paragraph_tags(&result)?;

        // Process table row tags {%tr ... %}
        result = self.process_table_row_tags(&result)?;

        // Process table cell tags {%tc ... %}
        result = self.process_table_cell_tags(&result)?;

        // Process run tags {%r ... %}
        result = self.process_run_tags(&result)?;

        Ok(result)
    }

    /// Process {%p ... %} tags (paragraph-level)
    fn process_paragraph_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Simply replace {%p ... %} with {% ... %}
        let re = Regex::new(r"\{%p\s+(.+?)\s*%\}")?;
        let count = re.find_iter(&result).count();
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1].trim())
        }).to_string();
        


        Ok(result)
    }

    /// Process {%tr ... %} tags (table row level)
    fn process_table_row_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Simply replace {%tr ... %} with {% ... %}
        let re = Regex::new(r"\{%tr\s+(.+?)\s*%\}")?;
        let count_before = re.find_iter(&result).count();
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1].trim())
        }).to_string();
        let count_after = Regex::new(r"\{%tr\s+").unwrap().find_iter(&result).count();
        


        Ok(result)
    }

    /// Process {%tc ... %} tags (table cell level)
    fn process_table_cell_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Simply replace {%tc ... %} with {% ... %}
        let re = Regex::new(r"\{%tc\s+(.+?)\s*%\}")?;
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1].trim())
        }).to_string();

        Ok(result)
    }

    /// Process {%r ... %} tags (run level)
    fn process_run_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Replace {%r ... %} with {% ... %} (control flow tags)
        let re = Regex::new(r"\{%r\s+(.+?)\s*%\}")?;
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1].trim())
        }).to_string();

        // Also handle {{r ... }} for variables - replace with {{ ... }}
        let re2 = Regex::new(r"\{\{r\s+(.+?)\s*\}\}")?;
        result = re2.replace_all(&result, |caps: &regex::Captures| {
            format!("{{{{ {} }}}}", &caps[1].trim())
        }).to_string();

        Ok(result)
    }

    /// Process inline images in rendered content
    fn process_inline_images(&mut self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Pattern to match image placeholders
        let re = Regex::new(r"__INLINE_IMAGE__([^_]+)__")?;

        for caps in re.captures_iter(content) {
            let image_path = &caps[1];

            // Create relationship for image
            let rel_id = format!("rId{}", self.next_rel_id);
            self.next_rel_id += 1;

            let doc_pr_id = self.image_counter;
            self.image_counter += 1;

            // Build image XML
            let image_xml = format!(
                r#"<w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="5715000" cy="3810000"/>
                        <wp:docPr id="{}" name="Picture {}"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="0" name="{}"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="{}"/>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="5715000" cy="3810000"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>"#,
                doc_pr_id, doc_pr_id, image_path, rel_id
            );

            result = result.replace(&caps[0], &image_xml);

            // Add relationship
            let rels = self
                .relationships
                .entry("word/_rels/document.xml.rels".to_string())
                .or_insert_with(Vec::new);
            rels.push((
                rel_id,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                    .to_string(),
                format!("media/{}", image_path),
            ));
        }

        Ok(result)
    }

    /// Apply media replacements
    fn apply_media_replacements(&mut self) -> Result<()> {
        // This is handled during save
        Ok(())
    }

    /// Write document to ZIP
    fn write_to_zip<W: Write + std::io::Seek>(&self, writer: &mut W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o644);

        // Write all XML parts (skip [Content_Types].xml as we'll write our own)
        for (path, content) in &self.xml_parts {
            if path.starts_with("__REPLACE__") {
                continue; // Skip replacement placeholders
            }
            if path == "[Content_Types].xml" {
                continue; // Skip - we'll write our own
            }

            zip.start_file(path, options)?;
            zip.write_all(content.as_bytes())?;
        }

        // Write relationships (only if not already in xml_parts)
        for (path, rels) in &self.relationships {
            if !self.xml_parts.contains_key(path) {
                zip.start_file(path, options)?;
                zip.write_all(self.build_relationships(rels)?.as_bytes())?;
            }
        }

        // Write binary parts (images, embeddings) - skip if replaced
        for (path, data) in &self.binary_parts {
            // Skip if this media file is being replaced
            if path.starts_with("word/media/") {
                let media_name = path.strip_prefix("word/media/").unwrap_or(path);
                if self.media_replacements.contains_key(media_name) {
                    continue;
                }
            }
            zip.start_file(path, options)?;
            zip.write_all(data)?;
        }

        // Handle media replacements
        for (old_name, new_path) in &self.media_replacements {
            let media_path = format!("word/media/{}", old_name);
            if let Ok(data) = fs::read(new_path) {
                zip.start_file(&media_path, options)?;
                zip.write_all(&data)?;
            }
        }

        // Write [Content_Types].xml
        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(self.build_content_types()?.as_bytes())?;

        zip.finish()?;
        Ok(())
    }

    /// Build [Content_Types].xml content
    fn build_content_types(&self) -> Result<String> {
        // Start with the original or build from scratch
        let mut types = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );

        // Add defaults
        types.push_str(
            r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>"#,
        );

        // Add overrides from content_types map
        for (part, content_type) in &self.content_types {
            if part.starts_with('/') {
                types.push_str(&format!(
                    r#"<Override PartName="{}" ContentType="{}"/>"#,
                    part, content_type
                ));
            }
        }

        types.push_str("</Types>");
        Ok(types)
    }

    /// Build relationships XML
    fn build_relationships(&self, rels: &[(String, String, String)]) -> Result<String> {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        for (id, rel_type, target) in rels {
            xml.push_str(&format!(
                r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
                id, rel_type, target
            ));
        }

        xml.push_str("</Relationships>");
        Ok(xml)
    }
}

/// Helper function to find the end position of a tag (handles nested tags)
fn find_tag_end(content: &str, start_pos: usize) -> Option<usize> {
    // Find the tag name from the start position
    let tag_start = &content[start_pos..];
    if let Some(tag_end) = tag_start.find('>') {
        let opening_tag = &tag_start[..=tag_end];
        
        // Extract tag name (handles attributes)
        let tag_name = if let Some(space_pos) = opening_tag.find(' ') {
            &opening_tag[1..space_pos]
        } else {
            &opening_tag[1..opening_tag.len()-1]
        };
        
        // Handle self-closing tags
        if opening_tag.ends_with("/>") {
            return Some(start_pos + tag_end + 1);
        }
        
        // Find closing tag
        let close_tag = format!("</{}>", tag_name);
        let mut depth = 1;
        let mut pos = start_pos + tag_end + 1;
        
        while pos < content.len() {
            if let Some(open_pos) = content[pos..].find(&format!("<{}>", tag_name)) {
                let absolute_open = pos + open_pos;
                if let Some(close_match) = content[pos..].find(&close_tag) {
                    let absolute_close = pos + close_match;
                    
                    if absolute_open < absolute_close {
                        depth += 1;
                        pos = absolute_open + tag_name.len() + 2;
                    } else {
                        depth -= 1;
                        if depth == 0 {
                            return Some(absolute_close + close_tag.len());
                        }
                        pos = absolute_close + close_tag.len();
                    }
                } else {
                    return None;
                }
            } else {
                // No more opening tags, find closing tag
                if let Some(close_match) = content[pos..].find(&close_tag) {
                    return Some(pos + close_match + close_tag.len());
                }
                return None;
            }
        }
    }
    None
}
