//! Core DocxTemplate implementation

use crate::image::InlineImage;
use crate::richtext::{Listing, RichText};
use crate::subdoc::{ColSpan, Subdoc, VerticalMerge, CellColor};
use crate::types::{DocxTplError, Result};
use crate::xml_utils::{
    escape_xml, extract_template_variables, postprocess_xml_content, preprocess_xml_content,
};
use minijinja::Value;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use regex::Regex;
use std::collections::HashMap;
use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
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
    ///     jinja_env: Optional custom Jinja2 environment with filters
    ///     autoescape: Whether to autoescape special characters (default: False)
    #[pyo3(signature = (context, jinja_env=None, autoescape=false))]
    fn render(
        &mut self,
        context: &Bound<'_, PyDict>,
        jinja_env: Option<PyObject>,
        autoescape: bool,
    ) -> PyResult<()> {
        // Convert Python context to HashMap
        let context_map = self.py_dict_to_context(context)?;

        // TODO: Support Python custom filters from jinja_env
        // For now, we ignore jinja_env parameter
        let _py_filters: HashMap<String, PyObject> = HashMap::new();

        // Process each XML part
        let part_keys: Vec<String> = self.xml_parts.keys().cloned().collect();

        for part_path in part_keys {
            if !part_path.starts_with("word/") {
                continue;
            }

            let content = self.xml_parts.get(&part_path).unwrap().clone();

            // Preprocess XML
            let preprocessed = preprocess_xml_content(&content);

            // Render with Jinja2
            let rendered = self.render_template(&preprocessed, &context_map, autoescape, HashMap::new())?;

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
            if path.starts_with("word/document") || path.starts_with("word/header") || path.starts_with("word/footer") {
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
        let part_rels = self.relationships
            .entry("word/_rels/document.xml.rels".to_string())
            .or_insert_with(Vec::new);
        part_rels.push((
            rel_id.clone(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink".to_string(),
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
        self.xml_parts.insert(format!("__REPLACE__{}", zip_name), String::from_utf8_lossy(&data).to_string());
        Ok(())
    }

    /// Reset all replacements (for multiple renderings)
    fn reset_replacements(&mut self) {
        self.media_replacements.clear();
        self.embedded_replacements.clear();
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
                // Store binary files
                let mut data = Vec::new();
                file.read_to_end(&mut data)?;
                self.xml_parts.insert(name, String::from_utf8_lossy(&data).to_string());
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
        py_filters: HashMap<String, PyObject>,
    ) -> Result<String> {
        let mut env = minijinja::Environment::new();

        // IMPORTANT: Disable autoescape before adding templates
        // This prevents minijinja from escaping XML tags
        env.set_auto_escape_callback(|_| minijinja::AutoEscape::None);

        // Add built-in filters
        env.add_filter("e", |s: String| escape_xml(&s));
        env.add_filter("escape", |s: String| escape_xml(&s));

        // TODO: Add Python custom filters
        // This requires complex bridging between Python functions and Rust closures
        // For now, we only support built-in filters

        // Add template
        env.add_template("doc", template)?;

        // Render
        let tmpl = env.get_template("doc")?;
        let result = tmpl.render(context)?;

        // Post-process: handle special tags ({%p %}, {%tr %}, etc.)
        let result = self.process_special_tags(&result)?;

        // Handle inline images
        let result = self.process_inline_images(&result)?;

        Ok(result)
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

        // Pattern to match {%p ... %} tags
        let re = Regex::new(r"<w:p[^>]*>.*?<w:t[^>]*>\s*\{%p\s+(.+?)%\}\s*</w:t>.*?</w:p>")?;

        // Replace with normal {% ... %} and mark paragraph for removal
        result = re
            .replace_all(&result, |caps: &regex::Captures| {
                format!("{{{{ % {} % }}}}", &caps[1])
            })
            .to_string();

        Ok(result)
    }

    /// Process {%tr ... %} tags (table row level)
    fn process_table_row_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Pattern to match {%tr ... %} tags
        let re = Regex::new(r"<w:tr[^>]*>.*?<w:t[^>]*>\s*\{%tr\s+(.+?)%\}\s*</w:t>.*?</w:tr>")?;

        result = re
            .replace_all(&result, |caps: &regex::Captures| {
                format!("{{{{ % {} % }}}}", &caps[1])
            })
            .to_string();

        Ok(result)
    }

    /// Process {%tc ... %} tags (table cell level)
    fn process_table_cell_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        let re = Regex::new(r"<w:tc[^>]*>.*?<w:t[^>]*>\s*\{%tc\s+(.+?)%\}\s*</w:t>.*?</w:tc>")?;

        result = re
            .replace_all(&result, |caps: &regex::Captures| {
                format!("{{{{ % {} % }}}}", &caps[1])
            })
            .to_string();

        Ok(result)
    }

    /// Process {%r ... %} tags (run level)
    fn process_run_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Pattern to match {%r ... %} tags
        let re = Regex::new(r"<w:t[^>]*>\s*\{%r\s+(.+?)%\}\s*</w:t>")?;

        result = re
            .replace_all(&result, |caps: &regex::Captures| {
                format!("{{{{ % {} % }}}}", &caps[1])
            })
            .to_string();

        // Also handle {{r ... }} for variables
        let re2 = Regex::new(r"<w:t[^>]*>\s*\{\{r\s+(.+?)\}\}\s*</w:t>")?;

        result = re2
            .replace_all(&result, |caps: &regex::Captures| {
                format!("{{{{ {} }}}}", &caps[1])
            })
            .to_string();

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
                doc_pr_id,
                doc_pr_id,
                image_path,
                rel_id
            );

            result = result.replace(&caps[0], &image_xml);

            // Add relationship
            let rels = self.relationships
                .entry("word/_rels/document.xml.rels".to_string())
                .or_insert_with(Vec::new);
            rels.push((
                rel_id,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".to_string(),
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

        // Write [Content_Types].xml
        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(self.build_content_types()?.as_bytes())?;

        // Write relationships (only if not already in xml_parts)
        for (path, rels) in &self.relationships {
            if !self.xml_parts.contains_key(path) {
                zip.start_file(path, options)?;
                zip.write_all(self.build_relationships(rels)?.as_bytes())?;
            }
        }

        // Handle media replacements
        for (old_name, new_path) in &self.media_replacements {
            let media_path = format!("word/media/{}", old_name);
            if let Ok(data) = fs::read(new_path) {
                zip.start_file(&media_path, options)?;
                zip.write_all(&data)?;
            }
        }

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
