//! XML utilities for parsing and modifying Word documents

use crate::types::{namespaces, TagType};
use quick_xml::events::BytesStart;

use regex::Regex;
use std::collections::HashMap;

/// XML namespace declarations for Word documents
pub fn get_namespace_declarations() -> Vec<(&'static str, &'static str)> {
    vec![
        ("w", namespaces::W),
        ("wp", namespaces::WP),
        ("a", namespaces::A),
        ("pic", namespaces::PIC),
        ("r", namespaces::R),
    ]
}

/// Convert a tag with prefix to regular Jinja2 format
pub fn normalize_tag(tag: &str) -> String {
    // Handle special prefixes: {%p %}, {%tr %}, {%tc %}, {%r %}
    let re = Regex::new(r"\{%([ptrc]?)\s+").unwrap();
    re.replace_all(tag, |caps: &regex::Captures| {
        let prefix = caps.get(1).map_or("", |m| m.as_str());
        if prefix.is_empty() {
            "{% ".to_string()
        } else {
            "{% ".to_string()
        }
    })
    .to_string()
}

/// Extract the tag type from a Jinja2 tag
pub fn extract_tag_type(tag: &str) -> TagType {
    let re = Regex::new(r"\{%([ptrc]?)\s+").unwrap();
    if let Some(caps) = re.captures(tag) {
        let prefix = caps.get(1).map_or("", |m| m.as_str());
        TagType::from_prefix(prefix)
    } else {
        TagType::Normal
    }
}

/// Check if a tag is a Jinja2 tag
pub fn is_jinja_tag(text: &str) -> bool {
    text.contains("{%") || text.contains("{{")
}

/// Extract all variable names from a Jinja2 template
pub fn extract_template_variables(template: &str) -> Vec<String> {
    let mut variables = Vec::new();

    // Match {{ variable }} patterns
    let var_re = Regex::new(r"\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*[^}]*\}\}").unwrap();
    for caps in var_re.captures_iter(template) {
        if let Some(var) = caps.get(1) {
            variables.push(var.as_str().to_string());
        }
    }

    // Match {% for x in y %} patterns
    let for_re = Regex::new(r"\{%\s*for\s+\w+\s+in\s+([a-zA-Z_][a-zA-Z0-9_]*)").unwrap();
    for caps in for_re.captures_iter(template) {
        if let Some(var) = caps.get(1) {
            variables.push(var.as_str().to_string());
        }
    }

    // Match {% if x %} patterns
    let if_re = Regex::new(r"\{%\s*if\s+([a-zA-Z_][a-zA-Z0-9_]*)").unwrap();
    for caps in if_re.captures_iter(template) {
        if let Some(var) = caps.get(1) {
            variables.push(var.as_str().to_string());
        }
    }

    // Remove duplicates while preserving order
    let mut seen = std::collections::HashSet::new();
    variables.retain(|v| seen.insert(v.clone()));

    variables
}

/// Clean XML text content for use in Jinja2 templates
pub fn clean_xml_text(text: &str) -> String {
    // Handle XML entities that might interfere with Jinja2
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
}

/// Escape special characters for XML
pub fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
}

/// Unescape XML entities
pub fn unescape_xml(text: &str) -> String {
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
        .replace("&#xA;", "\n")
        .replace("&#xD;", "\r")
        .replace("&#10;", "\n")
        .replace("&#13;", "\r")
}

/// Parse XML attributes from a BytesStart element
pub fn parse_attributes(elem: &BytesStart) -> HashMap<String, String> {
    let mut attrs = HashMap::new();
    for attr in elem.attributes() {
        if let Ok(a) = attr {
            let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&a.value).to_string();
            attrs.insert(key, value);
        }
    }
    attrs
}

/// Build attribute string for XML element
pub fn build_attributes(attrs: &HashMap<String, String>) -> String {
    attrs
        .iter()
        .map(|(k, v)| format!("{}=\"{}\"", escape_xml(k), escape_xml(v)))
        .collect::<Vec<_>>()
        .join(" ")
}

/// Get the local name of an element (without namespace prefix)
pub fn local_name(elem: &BytesStart) -> String {
    let name_ref = elem.name();
    let name = String::from_utf8_lossy(name_ref.as_ref());
    name.split(':').last().unwrap_or(&name).to_string()
}

/// Check if an element has a specific local name
pub fn has_local_name(elem: &BytesStart, name: &str) -> bool {
    local_name(elem) == name
}

/// Split text around Jinja2 tags
pub fn split_jinja_tags(text: &str) -> Vec<(String, bool)> {
    let mut result = Vec::new();
    let re = Regex::new(r"(\{%.*?%\}|\{\{.*?\}\})").unwrap();

    let mut last_end = 0;
    for mat in re.find_iter(text) {
        if mat.start() > last_end {
            result.push((text[last_end..mat.start()].to_string(), false));
        }
        result.push((mat.as_str().to_string(), true));
        last_end = mat.end();
    }

    if last_end < text.len() {
        result.push((text[last_end..].to_string(), false));
    }

    if result.is_empty() {
        result.push((text.to_string(), false));
    }

    result
}

/// Convert plain text to Word runs with proper newline handling
pub fn text_to_runs(text: &str) -> Vec<(String, bool)> {
    let mut runs = Vec::new();

    for (i, part) in text.split('\n').enumerate() {
        if i > 0 {
            runs.push(("<w:br/>".to_string(), true));
        }
        if !part.is_empty() {
            // Check if part contains tab characters
            for (j, tab_part) in part.split('\t').enumerate() {
                if j > 0 {
                    runs.push(("<w:tab/>".to_string(), true));
                }
                if !tab_part.is_empty() {
                    runs.push((escape_xml(tab_part), false));
                }
            }
        }
    }

    runs
}

/// Merge adjacent text segments
pub fn merge_text_segments(segments: Vec<(String, bool)>) -> Vec<(String, bool)> {
    if segments.is_empty() {
        return segments;
    }

    let mut result = Vec::new();
    let mut current_text = String::new();
    let mut current_is_tag = segments[0].1;

    for (text, is_tag) in segments {
        if is_tag == current_is_tag && !is_tag {
            // Merge adjacent text
            current_text.push_str(&text);
        } else {
            if !current_text.is_empty() {
                result.push((current_text, current_is_tag));
            }
            current_text = text;
            current_is_tag = is_tag;
        }
    }

    if !current_text.is_empty() {
        result.push((current_text, current_is_tag));
    }

    result
}

/// Remove XML declaration from content
pub fn remove_xml_declaration(content: &str) -> String {
    let re = Regex::new(r"<\?xml[^?]*\?>\s*").unwrap();
    re.replace(content, "").to_string()
}

/// Find all paragraph elements in XML content
pub fn find_paragraphs(xml: &str) -> Vec<(usize, usize)> {
    let mut paragraphs = Vec::new();
    let p_open = Regex::new(r"<w:p[\s>]").unwrap();
    let p_close = Regex::new(r"</w:p>").unwrap();

    let mut stack = Vec::new();

    for mat in p_open.find_iter(xml) {
        stack.push(mat.start());
    }

    for mat in p_close.find_iter(xml) {
        if let Some(start) = stack.pop() {
            paragraphs.push((start, mat.end()));
        }
    }

    paragraphs
}

/// Find all table row elements in XML content
pub fn find_table_rows(xml: &str) -> Vec<(usize, usize)> {
    let mut rows = Vec::new();
    let tr_open = Regex::new(r"<w:tr[\s>]").unwrap();
    let tr_close = Regex::new(r"</w:tr>").unwrap();

    let mut stack = Vec::new();

    for mat in tr_open.find_iter(xml) {
        stack.push(mat.start());
    }

    for mat in tr_close.find_iter(xml) {
        if let Some(start) = stack.pop() {
            rows.push((start, mat.end()));
        }
    }

    rows
}

/// Find all table cell elements in XML content
pub fn find_table_cells(xml: &str) -> Vec<(usize, usize)> {
    let mut cells = Vec::new();
    let tc_open = Regex::new(r"<w:tc[\s>]").unwrap();
    let tc_close = Regex::new(r"</w:tc>").unwrap();

    let mut stack = Vec::new();

    for mat in tc_open.find_iter(xml) {
        stack.push(mat.start());
    }

    for mat in tc_close.find_iter(xml) {
        if let Some(start) = stack.pop() {
            cells.push((start, mat.end()));
        }
    }

    cells
}

/// Find all run elements in XML content
pub fn find_runs(xml: &str) -> Vec<(usize, usize)> {
    let mut runs = Vec::new();
    let r_open = Regex::new(r"<w:r[\s>]").unwrap();
    let r_close = Regex::new(r"</w:r>").unwrap();

    let mut stack = Vec::new();

    for mat in r_open.find_iter(xml) {
        stack.push(mat.start());
    }

    for mat in r_close.find_iter(xml) {
        if let Some(start) = stack.pop() {
            runs.push((start, mat.end()));
        }
    }

    runs
}

/// Extract all text from <w:t> elements and rebuild the document
/// This is a simpler approach that handles split Jinja2 tags
fn rebuild_text_content(xml: &str) -> String {
    // Check if there are any split Jinja2 tags
    // A split tag looks like: {{... in one <w:t> and ...}} in another
    
    // First, let's identify if there's a split
    // Count {{ and }}
    let open_count = xml.matches("{{").count();
    let close_count = xml.matches("}}").count();
    
    if open_count == 0 || open_count != close_count {
        // Tags are either not present or mismatched - need reconstruction
        // Extract all text content and rebuild
        
        // Find all <w:t> elements and their content
        let re = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>").unwrap();
        let texts: Vec<&str> = re.captures_iter(xml)
            .filter_map(|cap| cap.get(1).map(|m| m.as_str()))
            .collect();
        
        // Join them to see the full text
        let full_text: String = texts.join("");
        
        // Check for split Jinja tags in the joined text
        if full_text.contains("{{") && full_text.contains("}}") {
            // Need to reconstruct - replace all <w:t>content</w:t> with merged version
            // This is a simplification - we replace the first <w:t>...</w:t> with merged content
            // and remove subsequent ones until we have complete tags
            
            let mut result = xml.to_string();
            
            // Simple approach: just remove bookmark elements and let the merge happen
            // at a higher level
            result = Regex::new(r"<w:bookmarkStart[^>]*/>\s*<w:bookmarkEnd[^>]*/>")
                .unwrap()
                .replace_all(&result, "")
                .to_string();
            
            return result;
        }
    }
    
    xml.to_string()
}

/// Merge Jinja2 tags that are split across multiple Word runs
/// 
/// Word sometimes inserts bookmarks that split Jinja2 tags across runs:
/// <w:r><w:t>{{ var</w:t></w:r><w:bookmarkStart/><w:bookmarkEnd/><w:r><w:t> }}</w:t></w:r>
/// 
/// This function merges the text content across these elements
fn merge_split_tags(xml: &str) -> String {
    let mut result = xml.to_string();
    
    // Step 1: Remove bookmark elements that Word uses for cursor tracking
    // Handle both with and without whitespace between tags
    result = Regex::new(r"<w:bookmarkStart[^>]*/>\s*<w:bookmarkEnd[^>]*/>")
        .unwrap()
        .replace_all(&result, "")
        .to_string();
    
    // Remove any other bookmark elements
    result = Regex::new(r"<w:bookmark[^>]*/>")
        .unwrap()
        .replace_all(&result, "")
        .to_string();
    
    // Step 2: Merge {{...}} patterns split across runs
    // Handle pattern: {{text</w:t></w:r><w:r><w:t>text}}
    // The key is to match across </w:t></w:r>...<w:r>...<w:t> boundaries
    loop {
        let new_result = Regex::new(
            r"(\{\{[^}]*?)</w:t>\s*</w:r>\s*<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?\s*<w:t[^>]*>([^}]*?\}\})"
        ).unwrap()
        .replace_all(&result, |caps: &regex::Captures| {
            format!("{}{}", &caps[1], &caps[2])
        })
        .to_string();
        
        if new_result == result {
            break;
        }
        result = new_result;
    }
    
    // Step 3: Merge {%-...%} patterns split across runs  
    loop {
        let new_result = Regex::new(
            r"(\{%[^%]*?)</w:t>\s*</w:r>\s*<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?\s*<w:t[^>]*>([^%]*?%\})"
        ).unwrap()
        .replace_all(&result, |caps: &regex::Captures| {
            format!("{}{}", &caps[1], &caps[2])
        })
        .to_string();
        
        if new_result == result {
            break;
        }
        result = new_result;
    }
    
    result
}

/// Preprocess XML to extract text content from runs for template processing
pub fn preprocess_xml_content(xml: &str) -> String {
    let mut result = xml.to_string();

    // Convert Chinese quotes to ASCII quotes in Jinja2 tags
    // These can appear in Word documents when user types quotes
    // Handle both Unicode characters and XML entity encodings
    // U+2018 LEFT SINGLE QUOTATION MARK -> '
    // U+2019 RIGHT SINGLE QUOTATION MARK -> '
    // U+201C LEFT DOUBLE QUOTATION MARK -> "
    // U+201D RIGHT DOUBLE QUOTATION MARK -> "
    
    // Unicode characters
    result = result.replace('\u{2018}', "'").replace('\u{2019}', "'");
    result = result.replace('\u{201C}', "\"").replace('\u{201D}', "\"");
    
    // XML decimal entities (8216 = 0x2018, 8217 = 0x2019, 8220 = 0x201C, 8221 = 0x201D)
    result = result.replace("&#8216;", "'").replace("&#8217;", "'");
    result = result.replace("&#8220;", "\"").replace("&#8221;", "\"");
    
    // XML hexadecimal entities
    result = result.replace("&#x2018;", "'").replace("&#x2019;", "'");
    result = result.replace("&#x201C;", "\"").replace("&#x201D;", "\"");

    // First, merge Jinja2 tags that are split across multiple runs
    result = merge_split_tags(&result);

    // Find text elements and unescape their content temporarily
    let t_re = Regex::new(r"(<w:t[^>]*>)(.*?)(</w:t>)").unwrap();
    result = t_re
        .replace_all(&result, |caps: &regex::Captures| {
            let prefix = &caps[1];
            let content = &caps[2];
            let suffix = &caps[3];
            
            // Special handling for Jinja2 tags: don't escape quotes inside them
            // Pattern to match Jinja2 tags
            // \{\{[^{}]*\}\} matches {{ ... }}
            // \{%[^{}%]*%\} matches {% ... %}
            let jinja_re = Regex::new(r"\{\{[^{}]*\}\}|\{%[^{}%]*%\}").unwrap();
            let mut new_content = String::new();
            let mut last_end = 0;
            
            for mat in jinja_re.find_iter(content) {
                // Process content before Jinja2 tag
                let before = &content[last_end..mat.start()];
                new_content.push_str(&escape_xml(&unescape_xml(before)));
                
                // Process Jinja2 tag itself - unescape but don't re-escape quotes
                let tag = &content[mat.start()..mat.end()];
                let unescaped_tag = unescape_xml(tag);
                // Don't escape single quotes in Jinja2 tags
                let processed_tag = unescaped_tag.replace('&', "&amp;")
                    .replace('<', "&lt;")
                    .replace('>', "&gt;")
                    .replace('"', "&quot;");
                // Note: we don't escape ' here
                new_content.push_str(&processed_tag);
                
                last_end = mat.end();
            }
            
            // Process remaining content after last Jinja2 tag
            if last_end < content.len() {
                let after = &content[last_end..];
                new_content.push_str(&escape_xml(&unescape_xml(after)));
            }
            
            // If no Jinja2 tags found, process normally
            if last_end == 0 {
                new_content = escape_xml(&unescape_xml(content));
            }
            
            format!("{}{}{}", prefix, new_content, suffix)
        })
        .to_string();

    result
}

/// Postprocess XML after template rendering
pub fn postprocess_xml_content(xml: &str) -> String {
    // Ensure proper XML structure
    let mut result = xml.to_string();

    // Remove truly empty runs (only whitespace between tags)
    let empty_run_re = Regex::new(r"<w:r>\s*</w:r>").unwrap();
    result = empty_run_re.replace_all(&result, "").to_string();

    // Remove runs with only rPr (no actual content)
    let empty_run_with_rpr_re = Regex::new(r"<w:r>\s*<w:rPr>.*?</w:rPr>\s*</w:r>").unwrap();
    result = empty_run_with_rpr_re.replace_all(&result, "").to_string();

    // Remove empty table rows (rows with no visible text content)
    // These are typically leftover from {%tr for %} and {%tr endfor %} tags
    // A row is empty if it has no text content or only whitespace
    result = remove_empty_table_rows(&result);

    result
}

/// Remove table rows that have no visible text content
fn remove_empty_table_rows(xml: &str) -> String {
    let mut result = xml.to_string();
    let tr_start_re = Regex::new(r"<w:tr\b[^>]*>").unwrap();
    let tr_end = "</w:tr>";
    
    // Find all table rows and check if they're empty
    let mut rows_to_remove: Vec<(usize, usize)> = Vec::new();
    
    for mat in tr_start_re.find_iter(&result) {
        let start = mat.start();
        if let Some(end) = result[start..].find(tr_end) {
            let end = start + end + tr_end.len();
            let row_content = &result[start..end];
            
            // Check if row has any non-empty text
            // Extract all <w:t> contents
            let t_re = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>").unwrap();
            let has_content = t_re.captures_iter(row_content).any(|caps| {
                caps.get(1).map(|m| m.as_str().trim()).unwrap_or("").len() > 0
            });
            
            if !has_content {
                rows_to_remove.push((start, end));
            }
        }
    }
    
    // Remove rows in reverse order to maintain indices
    for (start, end) in rows_to_remove.into_iter().rev() {
        result.replace_range(start..end, "");
    }
    
    result
}
