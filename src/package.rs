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
    pub const ENDNOTE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
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
        const HEX: &[u8; 16] = b"0123456789abcdef";
        let b = bytes.as_ref();
        let mut s = String::with_capacity(b.len() * 2);
        for &x in b {
            s.push(HEX[(x >> 4) as usize] as char);
            s.push(HEX[(x & 0x0f) as usize] as char);
        }
        s
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

/// One zip entry: name, payload, original compression method, mtime.
///
/// The payload is `Arc`-shared, so cloning a `Package` (per-render reload)
/// is a refcount bump per entry instead of a full memcpy of every blob.
/// Payloads of unmodified, already-compressed media entries (png/jpg/…) are
/// kept as the raw deflate stream (`Packed`) and inflated lazily on first
/// access: loading a media-heavy docx no longer inflates every image, and
/// `to_bytes` raw-copies the untouched stream without inflating it either.
#[derive(Debug, Clone)]
pub struct Entry {
    pub name: String,
    data: EntryData,
    pub compression: zip::CompressionMethod,
    pub mtime: Option<zip::DateTime>,
    /// true while the payload is byte-identical to the source zip entry it
    /// was loaded from (set()/swap_bytes() clear it): such entries can be
    /// raw-copied at save time without de/compression
    pristine: bool,
}

#[derive(Debug, Clone)]
enum EntryData {
    Ready(std::sync::Arc<[u8]>),
    Packed {
        raw: std::sync::Arc<[u8]>,
        /// uncompressed size (capacity hint)
        size: usize,
        cell: std::cell::OnceCell<std::sync::Arc<[u8]>>,
    },
}

/// Inflate a raw deflate stream. Corrupt streams yield an empty payload
/// (they can only come from a corrupt source zip; previously such a zip
/// failed at load time instead).
fn inflate_raw(raw: &[u8], size: usize) -> std::sync::Arc<[u8]> {
    let mut dec = flate2::read::DeflateDecoder::new(raw);
    let mut out = Vec::with_capacity(size);
    if dec.read_to_end(&mut out).is_err() {
        out.clear();
    }
    out.into()
}

impl Entry {
    /// decompressed payload (inflates packed entries on first access)
    pub fn bytes(&self) -> &[u8] {
        match &self.data {
            EntryData::Ready(b) => b,
            EntryData::Packed { raw, size, cell } => {
                cell.get_or_init(|| inflate_raw(raw, *size)).as_ref()
            }
        }
    }

    /// uncompressed length without forcing inflation
    fn data_len(&self) -> usize {
        match &self.data {
            EntryData::Ready(b) => b.len(),
            EntryData::Packed { size, .. } => *size,
        }
    }

    /// true while the payload can be raw-copied from the source archive:
    /// either still the untouched raw deflate stream, or any uncompressed
    /// entry that was never modified since load
    fn raw_copyable(&self) -> bool {
        match &self.data {
            EntryData::Packed { cell, .. } => cell.get().is_none(),
            EntryData::Ready(_) => self.pristine,
        }
    }

    /// decompressed payload as a shared Arc (inflates packed entries on
    /// first access): callers hold no borrow of the Package
    pub fn bytes_arc(&self) -> std::sync::Arc<[u8]> {
        match &self.data {
            EntryData::Ready(b) => b.clone(),
            EntryData::Packed { raw, size, cell } => {
                cell.get_or_init(|| inflate_raw(raw, *size)).clone()
            }
        }
    }

    fn set_bytes(&mut self, data: Vec<u8>) {
        self.data = EntryData::Ready(data.into());
        self.pristine = false;
    }

    /// replace the payload, returning the previous (materialized) bytes
    pub fn swap_bytes(&mut self, data: Vec<u8>) -> Vec<u8> {
        self.pristine = false;
        let old = std::mem::replace(&mut self.data, EntryData::Ready(data.into()));
        match old {
            EntryData::Ready(b) => b.to_vec(),
            EntryData::Packed { raw, size, cell } => match cell.into_inner() {
                Some(b) => b.to_vec(),
                None => inflate_raw(&raw, size).to_vec(),
            },
        }
    }
}

/// In-memory docx package
#[derive(Debug, Clone)]
pub struct Package {
    /// ordered zip entries
    pub entries: Vec<Entry>,
    pub index: HashMap<String, usize>,
    /// lazily built sha1 hex per word/media/* entry (image dedup); kept in
    /// sync by `set`
    media_sha1: HashMap<String, String>,
    /// parsed .rels per part (parse once per part; write-through on
    /// save_rels, invalidated by direct `set` of the .rels entry)
    rels_cache: std::cell::RefCell<HashMap<String, std::rc::Rc<Rels>>>,
    /// original zip bytes: lets `to_bytes` raw-copy untouched packed media
    /// entries without inflating them
    source: Option<std::sync::Arc<[u8]>>,
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
/// Zero-copy on the UTF-8 fast path (nearly every docx part).
pub fn decode_part_cow(blob: &[u8]) -> std::borrow::Cow<'_, str> {
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
    cow
}

/// Decode bytes to a string using the part's declared encoding.
pub fn decode_part(blob: &[u8]) -> String {
    decode_part_cow(blob).into_owned()
}

/// Encode a string for storage, honoring the part's declared encoding.
pub fn encode_part(content: &str, encoding: &str) -> Vec<u8> {
    let enc = encoding.to_lowercase();
    if enc == "utf-8" || enc == "utf8" || enc.is_empty() {
        return content.as_bytes().to_vec();
    }
    if enc == "utf-16le" || enc == "utf-16" || enc == "utf-16be" {
        // encoding_rs follows the WHATWG spec where UTF-16 encoders emit UTF-8,
        // so encode UTF-16 manually (BOM + real UTF-16 payload).
        let little = enc != "utf-16be";
        let mut out = Vec::with_capacity(2 + content.len() * 2);
        let bom: [u8; 2] = if little { [0xFF, 0xFE] } else { [0xFE, 0xFF] };
        out.extend_from_slice(&bom);
        for unit in content.encode_utf16() {
            let bytes = if little { unit.to_le_bytes() } else { unit.to_be_bytes() };
            out.extend_from_slice(&bytes);
        }
        return out;
    }
    let encoding = encoding_rs::Encoding::for_label(enc.as_bytes())
        .unwrap_or(encoding_rs::UTF_8);
    let (cow, _, _) = encoding.encode(content);
    cow.to_vec()
}

/// True for entries whose payload is already compressed (png/jpg/...):
/// deflating them gains ~0% but costs the full save CPU on multi-MB media,
/// so they are written with `Stored` instead.
fn incompressible_entry(name: &str) -> bool {
    let ext = name.rsplit('.').next().unwrap_or("");
    ext.eq_ignore_ascii_case("png")
        || ext.eq_ignore_ascii_case("jpg")
        || ext.eq_ignore_ascii_case("jpeg")
        || ext.eq_ignore_ascii_case("gif")
        || ext.eq_ignore_ascii_case("webp")
        || ext.eq_ignore_ascii_case("mp3")
        || ext.eq_ignore_ascii_case("mp4")
        || ext.eq_ignore_ascii_case("zip")
}

/// encode_part taking ownership of the content: the UTF-8 path (nearly
/// every docx part) is zero-copy via String::into_bytes.
pub fn encode_part_owned(content: String, encoding: &str) -> Vec<u8> {
    let enc = encoding.to_lowercase();
    if enc == "utf-8" || enc == "utf8" || enc.is_empty() {
        return content.into_bytes();
    }
    encode_part(&content, encoding)
}

/// Build a minimal valid docx (content types + root rels + document.xml)
/// wrapping the given body xml. Used to back the Subdoc query facade for
/// programmatically built (bound) subdocs.
pub fn minimal_docx(body_xml: &str) -> Vec<u8> {
    let ct = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let doc = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}<w:sectPr/></w:body></w:document>",
        body_xml
    );
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        let options: zip::write::SimpleFileOptions = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, data) in [
            ("[Content_Types].xml", &ct[..]),
            ("_rels/.rels", &rels[..]),
            ("word/document.xml", doc.as_bytes()),
        ] {
            // infallible for an in-memory cursor
            let _ = writer.start_file(name, options);
            let _ = writer.write_all(data);
        }
        let _ = writer.finish();
    }
    cursor.into_inner()
}

impl Package {
    pub fn from_bytes(data: &[u8]) -> Result<Package, String> {
        Self::from_archive(data, None)
    }

    /// from_bytes sharing the original zip bytes (no copy): untouched packed
    /// media entries can be raw-copied from it at save time
    pub fn from_bytes_arc(data: std::sync::Arc<[u8]>) -> Result<Package, String> {
        let src = data.clone();
        Self::from_archive(&data, Some(src))
    }

    fn from_archive(data: &[u8], source: Option<std::sync::Arc<[u8]>>) -> Result<Package, String> {
        let cursor = Cursor::new(data);
        let mut zip = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;
        let mut entries = Vec::with_capacity(zip.len());
        let mut index = HashMap::new();
        for i in 0..zip.len() {
            let (name, compression, mtime, size) = {
                let f = zip.by_index(i).map_err(|e| e.to_string())?;
                (
                    f.name().to_string(),
                    f.compression(),
                    f.last_modified(),
                    f.size() as usize,
                )
            };
            // already-compressed media stays as the raw deflate stream until
            // first access: loading a media-heavy docx skips inflating it
            let lazy = compression == zip::CompressionMethod::Deflated
                && incompressible_entry(&name);
            let data = if lazy {
                let mut f = zip.by_index_raw(i).map_err(|e| e.to_string())?;
                let mut raw = Vec::with_capacity(f.compressed_size() as usize);
                f.read_to_end(&mut raw).map_err(|e| e.to_string())?;
                EntryData::Packed {
                    raw: raw.into(),
                    size,
                    cell: std::cell::OnceCell::new(),
                }
            } else {
                let mut f = zip.by_index(i).map_err(|e| e.to_string())?;
                let mut buf = Vec::with_capacity(size);
                f.read_to_end(&mut buf).map_err(|e| e.to_string())?;
                EntryData::Ready(buf.into())
            };
            index.insert(name.clone(), entries.len());
            entries.push(Entry {
                name,
                data,
                compression,
                mtime,
                pristine: true,
            });
        }
        Ok(Package {
            entries,
            index,
            media_sha1: HashMap::new(),
            rels_cache: std::cell::RefCell::new(HashMap::new()),
            source,
        })
    }

    pub fn get(&self, name: &str) -> Option<&[u8]> {
        self.index.get(name).map(|&i| self.entries[i].bytes())
    }

    pub fn get_string(&self, name: &str) -> Option<String> {
        self.get(name).map(|b| decode_part(b))
    }

    /// get_cow without the full-content copy on the UTF-8 fast path
    /// (render of a multi-MB document part otherwise memcpy's it once)
    pub fn get_cow(&self, name: &str) -> Option<std::borrow::Cow<'_, str>> {
        self.get(name).map(decode_part_cow)
    }

    /// entry payload as a shared Arc (refcount bump; the Package stays
    /// unborrowed, so callers can mutate it while reading the bytes)
    pub fn get_arc(&self, name: &str) -> Option<std::sync::Arc<[u8]>> {
        self.index.get(name).map(|&i| self.entries[i].bytes_arc())
    }

    /// declared encoding of a part's current bytes
    pub fn encoding_of(&self, name: &str) -> String {
        self.get(name).map(detect_encoding).unwrap_or_else(|| "utf-8".to_string())
    }

    pub fn set(&mut self, name: &str, data: Vec<u8>) {
        let media_sha = name
            .starts_with("word/media/")
            .then(|| sha1_hex(&data));
        self.set_inner(name, data, media_sha);
    }

    /// set() with a caller-computed media hash (get_or_add_image already
    /// hashed the blob for dedup — don't hash multi-MB media twice)
    fn set_inner(&mut self, name: &str, data: Vec<u8>, media_sha: Option<String>) {
        if let Some(h) = media_sha {
            // keep the dedup cache consistent (cheap: one hash per set)
            self.media_sha1.insert(name.to_string(), h);
        }
        if name.ends_with(".rels") {
            self.rels_cache.borrow_mut().remove(name);
        }
        if let Some(&i) = self.index.get(name) {
            self.entries[i].set_bytes(data);
        } else {
            let compression = if incompressible_entry(name) {
                zip::CompressionMethod::Stored
            } else {
                zip::CompressionMethod::Deflated
            };
            self.index.insert(name.to_string(), self.entries.len());
            self.entries.push(Entry {
                name: name.to_string(),
                data: EntryData::Ready(data.into()),
                compression,
                mtime: None,
                pristine: false,
            });
        }
    }

    pub fn contains(&self, name: &str) -> bool {
        self.index.contains_key(name)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        // rough output estimate: compressed sizes are unknown, but total
        // uncompressed size is a safe upper bound for most docx
        let est: usize = self.entries.iter().map(|e| e.data_len()).sum();
        let mut cursor = Cursor::new(Vec::with_capacity(est));
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            let options: zip::write::SimpleFileOptions = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            // source archive for raw-copying untouched entries (no inflate
            // + no re-deflate; opened only when needed)
            let mut src_zip = match &self.source {
                Some(bytes) if self.entries.iter().any(|e| e.raw_copyable()) => {
                    zip::ZipArchive::new(Cursor::new(&bytes[..])).ok()
                }
                _ => None,
            };
            for entry in &self.entries {
                if entry.name.ends_with('/') {
                    continue; // skip directory entries (python-docx doesn't write them)
                }
                if entry.raw_copyable() {
                    if let Some(z) = src_zip.as_mut() {
                        match z.by_name(&entry.name) {
                            Ok(f) => {
                                writer.raw_copy_file(f).map_err(|e| e.to_string())?;
                                continue;
                            }
                            Err(e) => return Err(e.to_string()),
                        }
                    }
                }
                // already-compressed payloads are stored verbatim: deflate
                // on them is ~0% gain for the dominant save cost
                let method = if incompressible_entry(&entry.name) {
                    zip::CompressionMethod::Stored
                } else {
                    entry.compression
                };
                let mut options = options.compression_method(method);
                if let Some(dt) = entry.mtime {
                    options = options.last_modified_time(dt);
                }
                writer
                    .start_file(&entry.name, options)
                    .map_err(|e| e.to_string())?;
                writer.write_all(entry.bytes()).map_err(|e| e.to_string())?;
            }
            writer.finish().map_err(|e| e.to_string())?;
        }
        Ok(cursor.into_inner())
    }

    /// Load rels for a part (empty if no .rels entry); parsed once per part
    /// and cached until `save_rels`/`set` touch the underlying entry.
    /// Rc-shared: repeated lookups (header/footer discovery per render, image
    /// inserts) are refcount bumps instead of cloning every Rel.
    pub fn rels(&self, part: &str) -> std::rc::Rc<Rels> {
        let path = rels_path_for(part);
        if let Some(r) = self.rels_cache.borrow().get(&path) {
            return r.clone();
        }
        let rels = std::rc::Rc::new(match self.get_string(&path) {
            Some(xml) => Rels::from_xml(&xml),
            None => Rels::default(),
        });
        self.rels_cache
            .borrow_mut()
            .insert(path, rels.clone());
        rels
    }

    pub fn save_rels(&mut self, part: &str, rels: &Rels) {
        let path = rels_path_for(part);
        self.set(&path, rels.to_xml().into_bytes());
        // set() invalidates the cache for .rels entries; repopulate so the
        // next rels() on this part skips re-parsing
        self.rels_cache
            .borrow_mut()
            .insert(path, std::rc::Rc::new(rels.clone()));
    }

    /// Add a relationship to a part, returning the new rId
    pub fn add_rel(&mut self, part: &str, rel_type: &str, target: &str, is_external: bool) -> String {
        let mut rels = (*self.rels(part)).clone();
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

    /// Find existing image part path by sha1, if any. Hashes of media
    /// entries are computed once and cached (updated by `set`), so inserting
    /// N images costs N blob hashes instead of N × all-media hashing.
    pub fn find_image_by_sha1(&mut self, sha1: &str) -> Option<String> {
        for entry in &self.entries {
            if entry.name.starts_with("word/media/") && !self.media_sha1.contains_key(&entry.name) {
                let h = sha1_hex(entry.bytes());
                self.media_sha1.insert(entry.name.clone(), h);
            }
        }
        self.media_sha1
            .iter()
            .find(|(_, h)| h.as_str() == sha1)
            .map(|(name, _)| name.clone())
    }

    /// Next available /word/media/imageN.<ext> partname
    pub fn next_image_partname(&self, ext: &str) -> String {
        let mut used: std::collections::HashSet<u32> = std::collections::HashSet::new();
        for entry in &self.entries {
            if let Some(rest) = entry.name.strip_prefix("word/media/image") {
                let num_str: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
                if let Ok(n) = num_str.parse::<u32>() {
                    used.insert(n);
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
        self.set_inner(&partname, blob.to_vec(), Some(sha));
        self.ensure_content_type_default(ext, content_type);
        partname
    }
}

const DEFAULT_CORE_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type=\"dcterms:W3CDTF\">2000-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2000-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>";

/// Create the core properties part if missing (python-docx always has one).
pub fn ensure_core_part(pkg: &mut Package) {
    if pkg.contains("docProps/core.xml") {
        return;
    }
    pkg.set("docProps/core.xml", DEFAULT_CORE_XML.as_bytes().to_vec());
    pkg.ensure_content_type_override(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
    );
    let rels_path = "_rels/.rels";
    let mut rels = pkg
        .get_string(rels_path)
        .map(|x| Rels::from_xml(&x))
        .unwrap_or_default();
    rels.add(
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "docProps/core.xml",
        false,
    );
    pkg.set(rels_path, rels.to_xml().into_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    const CT_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
        <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
        <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
        </Types>";

    fn build_zip(entries: &[(&str, &[u8], zip::CompressionMethod)]) -> Vec<u8> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut cursor);
            for (name, data, method) in entries {
                let opts =
                    zip::write::SimpleFileOptions::default().compression_method(*method);
                w.start_file(name, opts).unwrap();
                w.write_all(data).unwrap();
            }
            w.finish().unwrap();
        }
        cursor.into_inner()
    }

    fn rels_pkg(rels_xml: &str) -> Package {
        let zip = build_zip(&[(
            "word/_rels/document.xml.rels",
            rels_xml.as_bytes(),
            zip::CompressionMethod::Deflated,
        )]);
        Package::from_bytes(&zip).unwrap()
    }

    // ---- hashes ----

    #[test]
    fn test_crc32_known_vector() {
        assert_eq!(crc32(b"123456789"), 0xCBF43926);
        assert_eq!(crc32(b""), 0);
    }

    #[test]
    fn test_sha1_hex_known_vector() {
        assert_eq!(
            sha1_hex(b"abc"),
            "a9993e364706816aba3e25717850c26c9cd0d89d"
        );
    }

    // ---- Rels ----

    #[test]
    fn test_rels_add_dedup_same_triple_reuses_rid() {
        let mut rels = Rels::default();
        let first = rels.add(rel_type::IMAGE, "media/a.png", false);
        let second = rels.add(rel_type::IMAGE, "media/a.png", false);
        assert_eq!(first, "rId1");
        assert_eq!(second, "rId1");
        assert_eq!(rels.rels.len(), 1);
    }

    #[test]
    fn test_rels_add_different_target_gets_new_rid() {
        let mut rels = Rels::default();
        rels.add(rel_type::IMAGE, "media/a.png", false);
        let id = rels.add(rel_type::IMAGE, "media/b.png", false);
        assert_eq!(id, "rId2");
        assert_eq!(rels.rels.len(), 2);
    }

    #[test]
    fn test_rels_add_different_type_same_target_gets_new_rid() {
        let mut rels = Rels::default();
        rels.add(rel_type::IMAGE, "media/a.png", false);
        let id = rels.add(rel_type::HEADER, "media/a.png", false);
        assert_eq!(id, "rId2");
    }

    #[test]
    fn test_rels_add_different_external_flag_gets_new_rid() {
        let mut rels = Rels::default();
        rels.add(rel_type::HYPERLINK, "https://example.com", false);
        let id = rels.add(rel_type::HYPERLINK, "https://example.com", true);
        assert_eq!(id, "rId2");
        // same triple with external=true now dedups against the second entry
        let again = rels.add(rel_type::HYPERLINK, "https://example.com", true);
        assert_eq!(again, "rId2");
    }

    #[test]
    fn test_rels_next_rid_uses_max_numeric_suffix() {
        let mut rels = Rels::default();
        for id in ["rId1", "rId7", "plain", "rIdX", "rId3x"] {
            rels.rels.push(Rel {
                id: id.to_string(),
                rel_type: String::new(),
                target: String::new(),
                is_external: false,
            });
        }
        assert_eq!(rels.next_rid(), "rId8");
        assert_eq!(Rels::default().next_rid(), "rId1");
    }

    #[test]
    fn test_rels_from_xml_parses_fields_and_external() {
        let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
            <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
            <Relationship Id=\"rId1\" Type=\"http://x/image\" Target=\"media/a.png\"/>\
            <Relationship Id=\"rId2\" Type=\"http://x/hyperlink\" Target=\"https://e.com?a=1&amp;b=2\" TargetMode=\"External\"/>\
            </Relationships>";
        let rels = Rels::from_xml(xml);
        assert_eq!(rels.rels.len(), 2);
        assert_eq!(rels.rels[0].id, "rId1");
        assert_eq!(rels.rels[0].rel_type, "http://x/image");
        assert_eq!(rels.rels[0].target, "media/a.png");
        assert!(!rels.rels[0].is_external);
        assert!(rels.rels[1].is_external);
        // attribute entities are decoded on parse
        assert_eq!(rels.rels[1].target, "https://e.com?a=1&b=2");
    }

    #[test]
    fn test_rels_from_xml_invalid_xml_yields_empty() {
        assert!(Rels::from_xml("not xml at all").rels.is_empty());
    }

    #[test]
    fn test_rels_to_xml_escapes_target_and_roundtrips() {
        let mut rels = Rels::default();
        let id = rels.add(rel_type::HYPERLINK, "https://e.com/?a=1&b=\"2\"", true);
        let xml = rels.to_xml();
        assert!(xml.contains("Target=\"https://e.com/?a=1&amp;b=&quot;2&quot;\""));
        assert!(xml.contains("TargetMode=\"External\""));
        let parsed = Rels::from_xml(&xml);
        assert_eq!(parsed.rels.len(), 1);
        assert_eq!(parsed.rels[0].id, id);
        assert_eq!(parsed.rels[0].target, "https://e.com/?a=1&b=\"2\"");
        assert!(parsed.rels[0].is_external);
    }

    #[test]
    fn test_rels_by_type_and_get() {
        let mut rels = Rels::default();
        let img = rels.add(rel_type::IMAGE, "media/a.png", false);
        rels.add(rel_type::HYPERLINK, "https://e.com", true);
        let images: Vec<&Rel> = rels.by_type(rel_type::IMAGE).collect();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].id, img);
        assert_eq!(rels.get(&img).unwrap().target, "media/a.png");
        assert!(rels.get("rId99").is_none());
    }

    // ---- escaping ----

    #[test]
    fn test_escape_xml_attr_escapes_amp_lt_quote() {
        assert_eq!(escape_xml_attr("a&b<c\"d'>e"), "a&amp;b&lt;c&quot;d'>e");
    }

    #[test]
    fn test_escape_xml_text_escapes_amp_lt_gt() {
        assert_eq!(escape_xml_text("a&b<c>\"d\""), "a&amp;b&lt;c&gt;\"d\"");
    }

    // ---- path helpers ----

    #[test]
    fn test_rels_path_for_nested_and_root() {
        assert_eq!(
            rels_path_for("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        assert_eq!(rels_path_for("document.xml"), "_rels/document.xml.rels");
    }

    #[test]
    fn test_part_dir_nested_and_root() {
        assert_eq!(part_dir("word/document.xml"), "word");
        assert_eq!(part_dir("a/b/c.xml"), "a/b");
        assert_eq!(part_dir("document.xml"), "");
    }

    #[test]
    fn test_resolve_target_relative_dotdot_and_absolute() {
        assert_eq!(
            resolve_target("word/document.xml", "media/image1.png"),
            "word/media/image1.png"
        );
        assert_eq!(
            resolve_target("word/document.xml", "../docProps/core.xml"),
            "docProps/core.xml"
        );
        assert_eq!(
            resolve_target("word/document.xml", "/word/media/a.png"),
            "word/media/a.png"
        );
        assert_eq!(resolve_target("document.xml", "media/a.png"), "media/a.png");
    }

    #[test]
    fn test_relative_target_common_and_parent_dirs() {
        assert_eq!(
            relative_target("word/document.xml", "word/media/image1.png"),
            "media/image1.png"
        );
        assert_eq!(
            relative_target("word/document.xml", "docProps/core.xml"),
            "../docProps/core.xml"
        );
        assert_eq!(
            relative_target("docProps/core.xml", "docProps/thumbnail.jpeg"),
            "thumbnail.jpeg"
        );
        // root-level part: no directory, absolute path returned as-is
        assert_eq!(relative_target("document.xml", "word/media/a.png"), "word/media/a.png");
    }

    // ---- encoding detection / transcoding ----

    #[test]
    fn test_detect_encoding_boms() {
        assert_eq!(detect_encoding(b"\xEF\xBB\xBF<a/>"), "utf-8");
        assert_eq!(detect_encoding(b"\xFF\xFE<\x00"), "utf-16le");
        assert_eq!(detect_encoding(b"\xFE\xFF\x00<"), "utf-16be");
    }

    #[test]
    fn test_detect_encoding_declaration_double_and_single_quotes() {
        assert_eq!(
            detect_encoding(b"<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><r/>"),
            "iso-8859-1"
        );
        assert_eq!(
            detect_encoding(b"<?xml version=\"1.0\" encoding='Windows-1252'?><r/>"),
            "windows-1252"
        );
    }

    #[test]
    fn test_detect_encoding_defaults_to_utf8() {
        assert_eq!(detect_encoding(b"<?xml version=\"1.0\"?><r/>"), "utf-8");
        assert_eq!(detect_encoding(b""), "utf-8");
        assert_eq!(detect_encoding(b"<no-declaration/>"), "utf-8");
    }

    #[test]
    fn test_decode_part_strips_utf8_bom() {
        assert_eq!(decode_part(b"\xEF\xBB\xBF<a>hi</a>"), "<a>hi</a>");
    }

    #[test]
    fn test_decode_part_latin1_declared() {
        let blob = b"<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><a>caf\xE9</a>";
        assert_eq!(decode_part(blob), "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><a>caf\u{E9}</a>");
    }

    fn utf16le_bytes(s: &str) -> Vec<u8> {
        let mut out = vec![0xFF, 0xFE];
        for u in s.encode_utf16() {
            out.extend_from_slice(&u.to_le_bytes());
        }
        out
    }

    fn utf16be_bytes(s: &str) -> Vec<u8> {
        let mut out = vec![0xFE, 0xFF];
        for u in s.encode_utf16() {
            out.extend_from_slice(&u.to_be_bytes());
        }
        out
    }

    #[test]
    fn test_decode_part_utf16le_with_bom_and_cjk() {
        let text = "<a>\u{4E2D}\u{6587}</a>";
        assert_eq!(detect_encoding(&utf16le_bytes(text)), "utf-16le");
        assert_eq!(decode_part(&utf16le_bytes(text)), text);
    }

    #[test]
    fn test_decode_part_utf16be_with_bom() {
        let text = "<a>hi</a>";
        assert_eq!(detect_encoding(&utf16be_bytes(text)), "utf-16be");
        assert_eq!(decode_part(&utf16be_bytes(text)), text);
    }

    #[test]
    fn test_encode_part_utf8_passthrough() {
        assert_eq!(encode_part("héllo", "utf-8"), "héllo".as_bytes());
        assert_eq!(encode_part("héllo", ""), "héllo".as_bytes());
    }

    // encoding_rs's `encode` for the UTF-16 labels follows the WHATWG spec and
    // emits *UTF-8* payload bytes, so encode_part encodes UTF-16 manually:
    // BOM + real UTF-16 payload, which decode_part reads back correctly.
    #[test]
    fn test_encode_part_utf16_roundtrip() {
        let text = "<a>\u{4E2D}\u{6587} héllo</a>";
        assert_eq!(encode_part(text, "utf-16le"), utf16le_bytes(text));
        assert_eq!(encode_part(text, "utf-16"), utf16le_bytes(text));
        assert_eq!(encode_part(text, "utf-16be"), utf16be_bytes(text));
        assert_eq!(decode_part(&encode_part(text, "utf-16le")), text);
        assert_eq!(decode_part(&encode_part(text, "utf-16be")), text);
    }

    #[test]
    fn test_encode_part_unknown_label_falls_back_to_utf8() {
        assert_eq!(encode_part("abc", "klingon"), b"abc");
    }

    // ---- Package zip I/O ----

    #[test]
    fn test_package_roundtrip_preserves_entries_and_compression() {
        let zip = build_zip(&[
            ("word/document.xml", b"<doc/>", zip::CompressionMethod::Stored),
            ("word/styles.xml", b"<styles/>", zip::CompressionMethod::Deflated),
        ]);
        let pkg = Package::from_bytes(&zip).unwrap();
        assert_eq!(pkg.get("word/document.xml"), Some(&b"<doc/>"[..]));
        assert!(pkg.contains("word/styles.xml"));
        assert!(!pkg.contains("word/missing.xml"));
        assert!(pkg.get("word/missing.xml").is_none());
        assert_eq!(pkg.entries[0].compression, zip::CompressionMethod::Stored);
        assert_eq!(pkg.entries[1].compression, zip::CompressionMethod::Deflated);

        let out = pkg.to_bytes().unwrap();
        let mut za = zip::ZipArchive::new(Cursor::new(&out)).unwrap();
        let mut names: Vec<String> = (0..za.len()).map(|i| za.by_index(i).unwrap().name().to_string()).collect();
        names.sort();
        assert_eq!(names, ["word/document.xml", "word/styles.xml"]);
        assert_eq!(za.by_name("word/document.xml").unwrap().compression(), zip::CompressionMethod::Stored);
        assert_eq!(za.by_name("word/styles.xml").unwrap().compression(), zip::CompressionMethod::Deflated);
        let mut buf = Vec::new();
        za.by_name("word/document.xml").unwrap().read_to_end(&mut buf).unwrap();
        assert_eq!(buf, b"<doc/>");
    }

    #[test]
    fn test_package_to_bytes_skips_directory_entries() {
        let zip = build_zip(&[
            ("word/", b"", zip::CompressionMethod::Stored),
            ("word/document.xml", b"<doc/>", zip::CompressionMethod::Deflated),
        ]);
        let pkg = Package::from_bytes(&zip).unwrap();
        let out = pkg.to_bytes().unwrap();
        let mut za = zip::ZipArchive::new(Cursor::new(&out)).unwrap();
        let names: Vec<String> = (0..za.len()).map(|i| za.by_index(i).unwrap().name().to_string()).collect();
        assert_eq!(names, ["word/document.xml"]);
    }

    #[test]
    fn test_package_to_bytes_preserves_mtime() {
        let dt = zip::DateTime::from_date_and_time(2020, 1, 2, 3, 4, 6).unwrap();
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut cursor);
            let opts = zip::write::SimpleFileOptions::default().last_modified_time(dt);
            w.start_file("word/document.xml", opts).unwrap();
            w.write_all(b"<doc/>").unwrap();
            w.finish().unwrap();
        }
        let pkg = Package::from_bytes(&cursor.into_inner()).unwrap();
        let out = pkg.to_bytes().unwrap();
        let mut za = zip::ZipArchive::new(Cursor::new(&out)).unwrap();
        assert_eq!(za.by_index(0).unwrap().last_modified(), Some(dt));
    }

    #[test]
    fn test_package_set_appends_deflated_and_overwrite_keeps_method() {
        let zip = build_zip(&[("a.xml", b"old", zip::CompressionMethod::Stored)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        // overwrite: data changes, compression method and entry order kept
        pkg.set("a.xml", b"new".to_vec());
        assert_eq!(pkg.get("a.xml"), Some(&b"new"[..]));
        assert_eq!(pkg.entries.len(), 1);
        assert_eq!(pkg.entries[0].compression, zip::CompressionMethod::Stored);
        // new entry: appended at the end with Deflated
        pkg.set("b.xml", b"x".to_vec());
        assert_eq!(pkg.entries.len(), 2);
        assert_eq!(pkg.entries[1].name, "b.xml");
        assert_eq!(pkg.entries[1].compression, zip::CompressionMethod::Deflated);
        assert_eq!(pkg.get("b.xml"), Some(&b"x"[..]));
    }

    #[test]
    fn test_package_get_string_decodes_declared_encoding() {
        let latin1 = b"<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><a>caf\xE9</a>";
        let zip = build_zip(&[("a.xml", latin1, zip::CompressionMethod::Deflated)]);
        let pkg = Package::from_bytes(&zip).unwrap();
        assert!(pkg.get_string("a.xml").unwrap().ends_with("caf\u{E9}</a>"));
        assert!(pkg.get_string("missing.xml").is_none());
    }

    #[test]
    fn test_package_encoding_of_declared_and_missing() {
        let latin1 = b"<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><a/>";
        let zip = build_zip(&[("a.xml", latin1, zip::CompressionMethod::Deflated)]);
        let pkg = Package::from_bytes(&zip).unwrap();
        assert_eq!(pkg.encoding_of("a.xml"), "iso-8859-1");
        assert_eq!(pkg.encoding_of("missing.xml"), "utf-8");
    }

    // ---- Package rels ----

    #[test]
    fn test_package_add_rel_writes_rels_entry_and_dedups() {
        let zip = build_zip(&[("word/document.xml", b"<doc/>", zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        let id1 = pkg.add_rel("word/document.xml", rel_type::IMAGE, "media/a.png", false);
        let id2 = pkg.add_rel("word/document.xml", rel_type::IMAGE, "media/a.png", false);
        assert_eq!(id1, "rId1");
        assert_eq!(id2, "rId1"); // dedup across save/load cycle
        let xml = pkg.get_string("word/_rels/document.xml.rels").unwrap();
        assert!(xml.contains("Id=\"rId1\""));
        assert!(xml.contains("Target=\"media/a.png\""));
        assert_eq!(xml.matches("<Relationship ").count(), 1);
    }

    #[test]
    fn test_package_rels_cache_invalidated_by_direct_set() {
        let xml_a = Rels::from_xml("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId5\" Type=\"http://x/image\" Target=\"a.png\"/></Relationships>");
        let mut pkg = rels_pkg(&xml_a.to_xml());
        assert!(pkg.rels("word/document.xml").get("rId5").is_some());
        // direct set of the .rels entry must drop the cached parse
        let xml_b = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId9\" Type=\"http://x/image\" Target=\"b.png\"/></Relationships>";
        pkg.set("word/_rels/document.xml.rels", xml_b.as_bytes().to_vec());
        let rels = pkg.rels("word/document.xml");
        assert!(rels.get("rId5").is_none());
        assert!(rels.get("rId9").is_some());
    }

    #[test]
    fn test_package_pristine_xml_raw_copied_and_dirty_rewritten() {
        let zip = build_zip(&[
            ("word/document.xml", b"<doc/>", zip::CompressionMethod::Deflated),
            ("word/styles.xml", b"<styles/>", zip::CompressionMethod::Deflated),
        ]);
        let mut pkg = Package::from_bytes_arc(zip.as_slice().into()).unwrap();
        // touch only document.xml; styles.xml stays pristine
        pkg.set("word/document.xml", b"<doc>changed</doc>".to_vec());
        let out = pkg.to_bytes().unwrap();
        let mut za = zip::ZipArchive::new(Cursor::new(&out)).unwrap();
        let mut buf = Vec::new();
        za.by_name("word/document.xml").unwrap().read_to_end(&mut buf).unwrap();
        assert_eq!(buf, b"<doc>changed</doc>");
        let mut buf2 = Vec::new();
        za.by_name("word/styles.xml").unwrap().read_to_end(&mut buf2).unwrap();
        assert_eq!(buf2, b"<styles/>");
        // both deflated (styles raw-copied from the deflated source)
        assert_eq!(
            za.by_name("word/styles.xml").unwrap().compression(),
            zip::CompressionMethod::Deflated
        );

        // a fully pristine package raw-copies everything: document.xml keeps
        // the source's exact deflate stream
        let pkg2 = Package::from_bytes_arc(zip.as_slice().into()).unwrap();
        assert_eq!(pkg2.to_bytes().unwrap(), zip);
    }

    #[test]
    fn test_package_packed_media_lazy_roundtrip() {
        // incompressible payload so zip deflate is stored-ish but still Deflated
        let payload: Vec<u8> = (0..100_000u32)
            .map(|i| i.wrapping_mul(2654435761) as u8)
            .collect();
        let zip = build_zip(&[
            ("word/document.xml", b"<doc/>", zip::CompressionMethod::Deflated),
            ("word/media/big.png", &payload, zip::CompressionMethod::Deflated),
        ]);
        let pkg = Package::from_bytes_arc(zip.as_slice().into()).unwrap();
        // media entry stays packed until accessed
        let idx = pkg.index["word/media/big.png"];
        assert!(pkg.entries[idx].raw_copyable());
        // first access inflates lazily and yields the original bytes
        assert_eq!(pkg.get("word/media/big.png"), Some(&payload[..]));

        // untouched packed entries are raw-copied at save: identical payload,
        // and the png keeps its Deflated method (raw stream copied verbatim)
        let pkg2 = Package::from_bytes_arc(zip.as_slice().into()).unwrap();
        let out = pkg2.to_bytes().unwrap();
        let mut za = zip::ZipArchive::new(Cursor::new(&out)).unwrap();
        let mut buf = Vec::new();
        za.by_name("word/media/big.png").unwrap().read_to_end(&mut buf).unwrap();
        assert_eq!(buf, payload);
        assert_eq!(
            za.by_name("word/media/big.png").unwrap().compression(),
            zip::CompressionMethod::Deflated
        );

        // overwritten media falls back to the normal Stored write path
        let mut pkg3 = Package::from_bytes_arc(zip.as_slice().into()).unwrap();
        pkg3.set("word/media/big.png", b"newpng".to_vec());
        let out3 = pkg3.to_bytes().unwrap();
        let mut za3 = zip::ZipArchive::new(Cursor::new(&out3)).unwrap();
        assert_eq!(
            za3.by_name("word/media/big.png").unwrap().compression(),
            zip::CompressionMethod::Stored
        );
        let mut buf3 = Vec::new();
        za3.by_name("word/media/big.png").unwrap().read_to_end(&mut buf3).unwrap();
        assert_eq!(buf3, b"newpng");
    }

    #[test]
    fn test_package_rels_missing_entry_yields_empty() {
        let zip = build_zip(&[("word/document.xml", b"<doc/>", zip::CompressionMethod::Deflated)]);
        let pkg = Package::from_bytes(&zip).unwrap();
        assert!(pkg.rels("word/document.xml").rels.is_empty());
    }

    // ---- content types ----

    #[test]
    fn test_ensure_content_type_default_inserts_before_closing_tag() {
        let zip = build_zip(&[("[Content_Types].xml", CT_XML.as_bytes(), zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        pkg.ensure_content_type_default("png", "image/png");
        let xml = pkg.get_string("[Content_Types].xml").unwrap();
        assert!(xml.contains("<Default Extension=\"png\" ContentType=\"image/png\"/></Types>"));
    }

    #[test]
    fn test_ensure_content_type_default_idempotent_for_existing_ext() {
        let zip = build_zip(&[("[Content_Types].xml", CT_XML.as_bytes(), zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        pkg.ensure_content_type_default("xml", "something/else");
        let xml = pkg.get_string("[Content_Types].xml").unwrap();
        assert!(!xml.contains("something/else"));
        assert_eq!(xml.matches("Extension=\"xml\"").count(), 1);
    }

    #[test]
    fn test_ensure_content_type_default_noop_without_content_types_part() {
        let zip = build_zip(&[("word/document.xml", b"<doc/>", zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        pkg.ensure_content_type_default("png", "image/png");
        assert!(!pkg.contains("[Content_Types].xml"));
    }

    #[test]
    fn test_ensure_content_type_override_inserts_and_is_idempotent() {
        let zip = build_zip(&[("[Content_Types].xml", CT_XML.as_bytes(), zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        pkg.ensure_content_type_override(
            "word/comments.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        );
        pkg.ensure_content_type_override("word/comments.xml", "other/type");
        let xml = pkg.get_string("[Content_Types].xml").unwrap();
        assert!(xml.contains("<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/></Types>"));
        assert!(!xml.contains("other/type"));
    }

    // ---- image parts ----

    #[test]
    fn test_next_image_partname_picks_lowest_free_number() {
        let zip = build_zip(&[
            ("word/media/image1.png", b"a", zip::CompressionMethod::Deflated),
            ("word/media/image3.png", b"b", zip::CompressionMethod::Deflated),
            ("word/media/image10.jpeg", b"c", zip::CompressionMethod::Deflated),
            ("word/media/imagesmith.png", b"d", zip::CompressionMethod::Deflated),
        ]);
        let pkg = Package::from_bytes(&zip).unwrap();
        assert_eq!(pkg.next_image_partname("png"), "word/media/image2.png");
    }

    #[test]
    fn test_find_image_by_sha1_hit_and_miss() {
        let blob = b"\x89PNG fake";
        let zip = build_zip(&[("word/media/image1.png", blob, zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        assert_eq!(
            pkg.find_image_by_sha1(&sha1_hex(blob)),
            Some("word/media/image1.png".to_string())
        );
        assert_eq!(pkg.find_image_by_sha1(&sha1_hex(b"other")), None);
    }

    #[test]
    fn test_get_or_add_image_dedups_by_sha1() {
        let zip = build_zip(&[("[Content_Types].xml", CT_XML.as_bytes(), zip::CompressionMethod::Deflated)]);
        let mut pkg = Package::from_bytes(&zip).unwrap();
        let p1 = pkg.get_or_add_image(b"img-bytes", "png", "image/png");
        let p2 = pkg.get_or_add_image(b"img-bytes", "png", "image/png");
        let p3 = pkg.get_or_add_image(b"different", "png", "image/png");
        assert_eq!(p1, "word/media/image1.png");
        assert_eq!(p2, p1); // dedup: same blob reused, no new part
        assert_eq!(p3, "word/media/image2.png");
        let media: Vec<&String> = pkg
            .entries
            .iter()
            .map(|e| &e.name)
            .filter(|n| n.starts_with("word/media/"))
            .collect();
        assert_eq!(media.len(), 2);
        // content type default registered exactly once
        let ct = pkg.get_string("[Content_Types].xml").unwrap();
        assert_eq!(ct.matches("Extension=\"png\"").count(), 1);
    }
}
