//! OPC (docx zip) package handling: entries, relationships, content types.

use crate::xmldom::{Document, Node};
use crc32fast::Hasher as Crc32;
use sha1::{Digest, Sha1};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

pub mod rel_type {
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
}

pub fn crc32(data: &[u8]) -> u32 {
    let mut h = Crc32::new();
    h.update(data);
    h.finalize()
}

pub fn sha1_hex(data: &[u8]) -> String {
    let mut h = Sha1::new();
    h.update(data);
    hex::encode(h.finalize())
}

mod hex {
    pub fn encode(bytes: impl AsRef<[u8]>) -> String {
        bytes.as_ref().iter().map(|b| format!("{:02x}", b)).collect()
    }
}

/// One relationship parsed from a .rels file
#[derive(Debug, Clone)]
pub struct Rel {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub is_external: bool,
}

/// Relationships for one part, backed by its .rels zip entry.
#[derive(Debug, Clone, Default)]
pub struct Rels {
    pub rels: Vec<Rel>,
}

impl Rels {
    pub fn from_xml(xml: &str) -> Rels {
        let mut out = Rels::default();
        if let Ok(doc) = Document::parse(xml) {
            for child in &doc.root.children {
                if let Node::Elem(e) = child {
                    if e.name == "Relationship" {
                        out.rels.push(Rel {
                            id: e.get_attr("Id").unwrap_or("").to_string(),
                            rel_type: e.get_attr("Type").unwrap_or("").to_string(),
                            target: e.get_attr("Target").unwrap_or("").to_string(),
                            is_external: e.get_attr("TargetMode") == Some("External"),
                        });
                    }
                }
            }
        }
        out
    }

    pub fn to_xml(&self) -> String {
        let mut s = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        );
        s.push_str("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for r in &self.rels {
            s.push_str("<Relationship Id=\"");
            s.push_str(&r.id);
            s.push_str("\" Type=\"");
            s.push_str(&r.rel_type);
            s.push_str("\" Target=\"");
            s.push_str(&escape_xml_attr(&r.target));
            s.push_str("\"");
            if r.is_external {
                s.push_str(" TargetMode=\"External\"");
            }
            s.push_str("/>");
        }
        s.push_str("</Relationships>");
        s
    }

    /// next available rId number
    pub fn next_rid(&self) -> String {
        let mut max = 0u32;
        for r in &self.rels {
            if let Some(rest) = r.id.strip_prefix("rId") {
                if let Ok(n) = rest.parse::<u32>() {
                    max = max.max(n);
                }
            }
        }
        format!("rId{}", max + 1)
    }

    pub fn add(&mut self, rel_type: &str, target: &str, is_external: bool) -> String {
        // python-docx reuses existing relationships to the same target
        if let Some(existing) = self
            .rels
            .iter()
            .find(|r| r.rel_type == rel_type && r.target == target && r.is_external == is_external)
        {
            return existing.id.clone();
        }
        let id = self.next_rid();
        self.rels.push(Rel {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            is_external,
        });
        id
    }

    pub fn by_type<'a>(&'a self, rel_type: &'a str) -> impl Iterator<Item = &'a Rel> + 'a {
        self.rels.iter().filter(move |r| r.rel_type == rel_type)
    }

    pub fn get(&self, id: &str) -> Option<&Rel> {
        self.rels.iter().find(|r| r.id == id)
    }
}

pub fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

#[allow(dead_code)]
pub fn escape_xml_text(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}

/// path of the .rels entry for a part, e.g. word/document.xml -> word/_rels/document.xml.rels
pub fn rels_path_for(part: &str) -> String {
    match part.rfind('/') {
        Some(i) => format!("{}/_rels/{}.rels", &part[..i], &part[i + 1..]),
        None => format!("_rels/{}.rels", part),
    }
}

/// directory containing the part ("word/document.xml" -> "word")
pub fn part_dir(part: &str) -> &str {
    match part.rfind('/') {
        Some(i) => &part[..i],
        None => "",
    }
}

/// Resolve a relative target against the part's directory (no leading slash)
pub fn resolve_target(part: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let dir = part_dir(part);
    let combined = if dir.is_empty() {
        target.to_string()
    } else {
        format!("{}/{}", dir, target)
    };
    normalize_path(&combined)
}

/// Compute relative path from part directory to an absolute package path
pub fn relative_target(part: &str, abs_path: &str) -> String {
    let dir = part_dir(part);
    if dir.is_empty() {
        return abs_path.to_string();
    }
    let dir_segs: Vec<&str> = dir.split('/').collect();
    let path_segs: Vec<&str> = abs_path.split('/').collect();
    let mut common = 0;
    while common < dir_segs.len()
        && common < path_segs.len()
        && dir_segs[common] == path_segs[common]
    {
        common += 1;
    }
    let mut rel = String::new();
    for _ in common..dir_segs.len() {
        rel.push_str("../");
    }
    rel.push_str(&path_segs[common..].join("/"));
    rel
}

fn normalize_path(p: &str) -> String {
    let mut segs: Vec<&str> = Vec::new();
    for s in p.split('/') {
        match s {
            "" | "." => {}
            ".." => {
                segs.pop();
            }
            _ => segs.push(s),
        }
    }
    segs.join("/")
}

/// In-memory docx package
#[derive(Debug, Clone)]
pub struct Package {
    /// ordered zip entries (name, bytes, original compression method, mtime)
    pub entries: Vec<(String, Vec<u8>, zip::CompressionMethod, Option<zip::DateTime>)>,
    pub index: HashMap<String, usize>,
}

/// Detect the encoding of an XML part from BOM / xml declaration.
pub fn detect_encoding(blob: &[u8]) -> String {
    if blob.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return "utf-8".to_string();
    }
    if blob.starts_with(&[0xFF, 0xFE]) {
        return "utf-16le".to_string();
    }
    if blob.starts_with(&[0xFE, 0xFF]) {
        return "utf-16be".to_string();
    }
    let head = String::from_utf8_lossy(&blob[..blob.len().min(200)]);
    if let Some(pos) = head.find("encoding") {
        let rest = &head[pos + "encoding".len()..];
        let rest = rest.trim_start_matches(['=', ' ', '\'', '"']);
        let enc: String = rest
            .chars()
            .take_while(|c| c.is_ascii_alphanumeric() || *c == '-' || *c == '_')
            .collect();
        if !enc.is_empty() {
            return enc.to_lowercase();
        }
    }
    "utf-8".to_string()
}

/// Decode bytes to a string using the part's declared encoding.
pub fn decode_part(blob: &[u8]) -> String {
    let enc = detect_encoding(blob);
    let (blob, bom_len) = match enc.as_str() {
        "utf-16le" | "utf-16be" => (&blob[blob.len().min(2)..], 2),
        _ if blob.starts_with(&[0xEF, 0xBB, 0xBF]) => (&blob[3..], 3),
        _ => (blob, 0),
    };
    let _ = bom_len;
    let encoding = encoding_rs::Encoding::for_label(enc.as_bytes())
        .unwrap_or(encoding_rs::UTF_8);
    let (cow, _, _) = encoding.decode(blob);
    cow.to_string()
}

/// Encode a string for storage, honoring the part's declared encoding.
pub fn encode_part(content: &str, encoding: &str) -> Vec<u8> {
    let enc = encoding.to_lowercase();
    if enc == "utf-8" || enc == "utf8" || enc.is_empty() {
        return content.as_bytes().to_vec();
    }
    let encoding = encoding_rs::Encoding::for_label(enc.as_bytes())
        .unwrap_or(encoding_rs::UTF_8);
    let (cow, _, _) = encoding.encode(content);
    let bytes = cow.to_vec();
    if enc == "utf-16le" || enc == "utf-16" {
        let mut out = vec![0xFF, 0xFE];
        out.extend_from_slice(&bytes);
        out
    } else if enc == "utf-16be" {
        let mut out = vec![0xFE, 0xFF];
        out.extend_from_slice(&bytes);
        out
    } else {
        bytes
    }
}

impl Package {
    pub fn from_bytes(data: &[u8]) -> Result<Package, String> {
        let cursor = Cursor::new(data);
        let mut zip = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;
        let mut entries = Vec::with_capacity(zip.len());
        let mut index = HashMap::new();
        for i in 0..zip.len() {
            let mut f = zip.by_index(i).map_err(|e| e.to_string())?;
            let name = f.name().to_string();
            let compression = f.compression();
            let mtime = f.last_modified();
            let mut buf = Vec::with_capacity(f.size() as usize);
            f.read_to_end(&mut buf).map_err(|e| e.to_string())?;
            index.insert(name.clone(), entries.len());
            entries.push((name, buf, compression, mtime));
        }
        Ok(Package { entries, index })
    }

    pub fn get(&self, name: &str) -> Option<&[u8]> {
        self.index.get(name).map(|&i| self.entries[i].1.as_slice())
    }

    pub fn get_string(&self, name: &str) -> Option<String> {
        self.get(name).map(|b| decode_part(b))
    }

    /// declared encoding of a part's current bytes
    pub fn encoding_of(&self, name: &str) -> String {
        self.get(name).map(detect_encoding).unwrap_or_else(|| "utf-8".to_string())
    }

    pub fn set(&mut self, name: &str, data: Vec<u8>) {
        if let Some(&i) = self.index.get(name) {
            self.entries[i].1 = data;
        } else {
            self.index.insert(name.to_string(), self.entries.len());
            self.entries.push((
                name.to_string(),
                data,
                zip::CompressionMethod::Deflated,
                None,
            ));
        }
    }

    pub fn contains(&self, name: &str) -> bool {
        self.index.contains_key(name)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            let options: zip::write::SimpleFileOptions = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for (name, data, compression, mtime) in &self.entries {
                if name.ends_with('/') {
                    continue; // skip directory entries (python-docx doesn't write them)
                }
                let mut options = options.compression_method(*compression);
                if let Some(dt) = mtime {
                    options = options.last_modified_time(*dt);
                }
                writer.start_file(name, options).map_err(|e| e.to_string())?;
                writer.write_all(data).map_err(|e| e.to_string())?;
            }
            writer.finish().map_err(|e| e.to_string())?;
        }
        Ok(cursor.into_inner())
    }

    /// Load rels for a part (empty if no .rels entry)
    pub fn rels(&self, part: &str) -> Rels {
        match self.get_string(&rels_path_for(part)) {
            Some(xml) => Rels::from_xml(&xml),
            None => Rels::default(),
        }
    }

    pub fn save_rels(&mut self, part: &str, rels: &Rels) {
        self.set(&rels_path_for(part), rels.to_xml().into_bytes());
    }

    /// Add a relationship to a part, returning the new rId
    pub fn add_rel(&mut self, part: &str, rel_type: &str, target: &str, is_external: bool) -> String {
        let mut rels = self.rels(part);
        let id = rels.add(rel_type, target, is_external);
        self.save_rels(part, &rels);
        id
    }

    // ---- content types ----

    pub fn ensure_content_type_default(&mut self, ext: &str, content_type: &str) {
        let name = "[Content_Types].xml";
        let Some(xml) = self.get_string(name) else {
            return;
        };
        // check if a Default for this extension already exists
        let pat = format!("Extension=\"{}\"", ext);
        if xml.contains(&pat) {
            return;
        }
        let insertion = format!(
            "<Default Extension=\"{}\" ContentType=\"{}\"/>",
            ext, content_type
        );
        let new_xml = xml.replace("</Types>", &(insertion + "</Types>"));
        self.set(name, new_xml.into_bytes());
    }

    pub fn ensure_content_type_override(&mut self, part_name: &str, content_type: &str) {
        let name = "[Content_Types].xml";
        let Some(xml) = self.get_string(name) else {
            return;
        };
        let pat = format!("PartName=\"/{}\"", part_name);
        if xml.contains(&pat) {
            return;
        }
        let insertion = format!(
            "<Override PartName=\"/{}\" ContentType=\"{}\"/>",
            part_name, content_type
        );
        let new_xml = xml.replace("</Types>", &(insertion + "</Types>"));
        self.set(name, new_xml.into_bytes());
    }

    // ---- image parts ----

    /// Find existing image part path by sha1, if any
    pub fn find_image_by_sha1(&self, sha1: &str) -> Option<String> {
        for (name, data, _, _) in &self.entries {
            if name.starts_with("word/media/") && sha1_hex(data) == sha1 {
                return Some(name.clone());
            }
        }
        None
    }

    /// Next available /word/media/imageN.<ext> partname
    pub fn next_image_partname(&self, ext: &str) -> String {
        let mut used: Vec<u32> = Vec::new();
        for (name, _, _, _) in &self.entries {
            if let Some(rest) = name.strip_prefix("word/media/image") {
                let num_str: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
                if let Ok(n) = num_str.parse::<u32>() {
                    used.push(n);
                }
            }
        }
        let mut n = 1u32;
        while used.contains(&n) {
            n += 1;
        }
        format!("word/media/image{}.{}", n, ext)
    }

    /// Add image part (dedup by sha1), returns package path like "word/media/image3.png"
    pub fn get_or_add_image(&mut self, blob: &[u8], ext: &str, content_type: &str) -> String {
        let sha = sha1_hex(blob);
        if let Some(existing) = self.find_image_by_sha1(&sha) {
            return existing;
        }
        let partname = self.next_image_partname(ext);
        self.set(&partname, blob.to_vec());
        self.ensure_content_type_default(ext, content_type);
        partname
    }
}
