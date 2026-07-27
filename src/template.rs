//! Core template engine: render pipeline, fix_tables, docPr ids,
//! headers/footers, properties, footnotes, replacements, save.

use crate::package::{crc32, rel_type, resolve_target, Package};
use crate::patch::{decode_text_entities, patch_xml, resolve_listing, sub_str};
use crate::xmldom::{Document, Element, Node};
use minijinja::{AutoEscape, Environment, Value};
use pyo3::prelude::PyAnyMethods as _;
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

pub const DOCUMENT_PART: &str = "word/document.xml";

/// A context value whose XML must be materialized lazily (only if actually
/// printed by the template), like docxtpl's __str__-based InlineImage/Subdoc.
/// Blob fields are Arc-shared so registering/cloning a Deferred is O(1).
#[derive(Debug, Clone)]
pub enum Deferred {
    Image {
        blob: std::sync::Arc<[u8]>,
        filename: Option<String>,
        width: Option<i64>,
        height: Option<i64>,
        anchor: Option<String>,
        title: Option<String>,
        descr: Option<String>,
    },
    Subdoc {
        bytes: Option<std::sync::Arc<[u8]>>,
    },
    SubdocBlocks {
        blocks: std::sync::Arc<Vec<crate::subdocbuilder::Block>>,
    },
}

#[derive(Debug, Default)]
pub struct TplCore {
    pub original_bytes: Vec<u8>,
    pub package: Option<Package>,
    pub is_rendered: bool,
    pub is_saved: bool,
    pub allow_missing_pics: bool,
    pub crc_to_new_media: HashMap<u32, Vec<u8>>,
    pub crc_to_new_embedded: HashMap<u32, Vec<u8>>,
    pub zipname_to_replace: HashMap<String, Vec<u8>>,
    pub pics_to_replace: HashMap<String, Vec<u8>>,
    /// filename -> (target_ref, target_partname)
    pub pic_map: HashMap<String, (String, String)>,
    pub docx_ids_index: u32,
    /// lazily materialized context values (InlineImage / Subdoc)
    pub deferred: Vec<Deferred>,
    /// python object id -> deferred index (per render)
    pub deferred_by_oid: HashMap<usize, usize>,
    /// user-registered jinja customizations (python callables)
    pub custom_filters: Vec<(String, pyo3::Py<pyo3::PyAny>)>,
    pub custom_tests: Vec<(String, pyo3::Py<pyo3::PyAny>)>,
    pub custom_functions: Vec<(String, pyo3::Py<pyo3::PyAny>)>,
    pub custom_globals: Vec<(String, pyo3::Py<pyo3::PyAny>)>,
    pub template_loader: Option<pyo3::Py<pyo3::PyAny>>,
    /// set when a subdoc was materialized during the current render
    pub used_subdoc: bool,
    /// jinja environment options (duck-typed from jinja_env)
    pub env_options: EnvOptions,
    /// installed gettext catalog for {% trans %} support
    pub gettext_catalog: Option<std::sync::Arc<crate::gettext::Catalog>>,
    /// context lines of the last template error (docx_context attribute)
    pub last_error_context: Vec<String>,
    /// cached parse of word/document.xml for the docmodel facade
    pub doc_cache: Option<crate::xmldom::Document>,
    /// doc_cache has unpersisted mutations
    pub doc_dirty: bool,
    /// cached parses of auxiliary xml parts (word/styles.xml, word/settings.xml,
    /// word/comments.xml, docProps/core.xml, ...) for the docmodel facade;
    /// word/document.xml uses doc_cache instead
    pub part_caches: HashMap<String, crate::xmldom::Document>,
    /// auxiliary part caches with unpersisted mutations
    pub parts_dirty: HashSet<String>,
    /// per-part next wp:docPr shape id (avoids rescanning the part xml for
    /// every inserted image; seeded from the part's current max id)
    pub next_shape_ids: HashMap<String, u32>,
}

/// jinja2 environment options supported via duck-typing.
#[derive(Debug, Clone, Default)]
pub struct EnvOptions {
    pub trim_blocks: Option<bool>,
    pub lstrip_blocks: Option<bool>,
    pub keep_trailing_newline: Option<bool>,
    /// "lenient" | "chainable" | "semistrict" | "strict"
    pub undefined_behavior: Option<String>,
}

/// placeholder token for a deferred value inside rendered xml
pub fn deferred_token(idx: usize) -> String {
    format!("\u{1}DTPLD{}\u{1}", idx)
}

pub type CtxFn<'a> = dyn Fn(&mut TplCore, &str) -> Result<Value, String> + 'a;

impl TplCore {
    pub fn new(template_bytes: Vec<u8>) -> TplCore {
        TplCore {
            original_bytes: template_bytes,
            ..Default::default()
        }
    }

    pub fn init_docx(&mut self, reload: bool) -> Result<(), String> {
        if self.package.is_none() || (self.is_rendered && reload) {
            self.package = Some(Package::from_bytes(&self.original_bytes)?);
            self.is_rendered = false;
            self.invalidate_doc();
            self.invalidate_parts();
        }
        Ok(())
    }

    /// Parsed word/document.xml for the docmodel facade, parsed once and
    /// cached; reparsed after invalidation, on parse failure retried per call.
    pub fn document_dom(&mut self) -> Result<&mut crate::xmldom::Document, String> {
        self.init_docx(false)?;
        if self.doc_cache.is_none() {
            let xml = self
                .package
                .as_ref()
                .and_then(|p| p.get_string(DOCUMENT_PART))
                .ok_or_else(|| "word/document.xml not found".to_string())?;
            self.doc_cache = Some(crate::xmldom::Document::parse(&xml)?);
        }
        Ok(self.doc_cache.as_mut().unwrap())
    }

    pub fn mark_doc_dirty(&mut self) {
        self.doc_dirty = true;
    }

    /// Serialize the cached document DOM back into the package if mutated.
    pub fn flush_doc(&mut self) -> Result<(), String> {
        if !self.doc_dirty {
            return Ok(());
        }
        let Some(dom) = &self.doc_cache else {
            self.doc_dirty = false;
            return Ok(());
        };
        let out = dom.serialize();
        let pkg = self.pkg()?;
        let enc = pkg.encoding_of(DOCUMENT_PART);
        pkg.set(DOCUMENT_PART, crate::package::encode_part(&out, &enc));
        self.doc_dirty = false;
        Ok(())
    }

    /// Drop the cached document DOM (external code replaced the part).
    pub fn invalidate_doc(&mut self) {
        self.doc_cache = None;
        self.doc_dirty = false;
    }

    /// Parsed DOM of an auxiliary part (styles/settings/comments/core/...),
    /// parsed once and cached; reparsed after invalidation.
    /// word/document.xml is served by document_dom instead.
    pub fn part_dom(&mut self, part: &str) -> Result<&mut crate::xmldom::Document, String> {
        if part == DOCUMENT_PART {
            return self.document_dom();
        }
        self.init_docx(false)?;
        if !self.part_caches.contains_key(part) {
            let xml = self
                .package
                .as_ref()
                .and_then(|p| p.get_string(part))
                .ok_or_else(|| format!("part {} not found", part))?;
            self.part_caches
                .insert(part.to_string(), crate::xmldom::Document::parse(&xml)?);
        }
        Ok(self.part_caches.get_mut(part).unwrap())
    }

    /// Mark a cached part DOM as mutated; it is written back on flush_parts.
    pub fn mark_part_dirty(&mut self, part: &str) {
        if part == DOCUMENT_PART {
            self.doc_dirty = true;
        } else {
            self.parts_dirty.insert(part.to_string());
        }
    }

    /// Serialize all dirty cached DOMs (document + auxiliary parts) back
    /// into the package.
    pub fn flush_parts(&mut self) -> Result<(), String> {
        self.flush_doc()?;
        if self.parts_dirty.is_empty() {
            return Ok(());
        }
        let dirty: Vec<String> = self.parts_dirty.drain().collect();
        for part in dirty {
            let Some(out) = self.part_caches.get(&part).map(|d| d.serialize()) else {
                continue;
            };
            let enc = self.pkg()?.encoding_of(&part);
            let bytes = crate::package::encode_part(&out, &enc);
            self.pkg()?.set(&part, bytes);
        }
        Ok(())
    }

    /// Drop one cached part DOM (external code replaced that part).
    pub fn invalidate_part(&mut self, part: &str) {
        if part == DOCUMENT_PART {
            self.invalidate_doc();
            return;
        }
        self.part_caches.remove(part);
        self.parts_dirty.remove(part);
    }

    /// Drop all cached auxiliary part DOMs (package was reloaded).
    pub fn invalidate_parts(&mut self) {
        self.part_caches.clear();
        self.parts_dirty.clear();
    }

    fn pkg(&mut self) -> Result<&mut Package, String> {
        self.package.as_mut().ok_or_else(|| "package not loaded".to_string())
    }

    pub fn get_xml(&mut self) -> Result<String, String> {
        self.init_docx(false)?;
        self.flush_parts()?;
        self.pkg()?
            .get_string(DOCUMENT_PART)
            .ok_or_else(|| "word/document.xml not found".to_string())
    }

    // ---------------- render pipeline ----------------

    pub fn render(&mut self, autoescape: bool, make_ctx: &CtxFn) -> Result<(), String> {
        self.render_init()?;
        // persist docmodel edits made since the last render/reload
        self.flush_parts()?;

        // Build the context and the jinja environment ONCE per render and
        // reuse them for every part (Value/Environment clones are cheap);
        // previously each of the K parts rebuilt both from scratch.
        let ctx = make_ctx(self, DOCUMENT_PART)?;
        let ctx_once: &CtxFn = &|_core, _part| Ok(ctx.clone());
        let env = make_env(autoescape, self);

        // Body
        let src = self.pkg()?.get_string(DOCUMENT_PART)
            .ok_or_else(|| "word/document.xml not found".to_string())?;
        // fix tables / docPr / cNvPr ids only after every part has rendered:
        // used_subdoc is final by then, so the body is parsed+serialized once
        let body_rendered = self.render_part_with(DOCUMENT_PART, &src, &env, ctx_once)?;

        // Headers & footers
        for uri in [rel_type::HEADER, rel_type::FOOTER] {
            let parts = self.header_footer_parts(uri);
            for part in parts {
                let src = match self.pkg()?.get_string(&part) {
                    Some(s) => s,
                    None => continue,
                };
                let rendered = self.render_part_with(&part, &src, &env, ctx_once)?;
                self.set_part_rendered(&part, rendered);
            }
        }

        // core properties always render without autoescape (docxtpl parity)
        let props_env;
        let props_env = if autoescape {
            props_env = make_env(false, self);
            &props_env
        } else {
            &env
        };
        self.render_properties(props_env, ctx_once)?;
        self.render_footnotes(&env, ctx_once)?;

        // like docxcompose's renumber_nvpicpr_ids: when subdocs were merged,
        // make pic:cNvPr ids unique (body first, then headers/footers). The
        // body's cNvPr renumbering rides along with its fix_tables pass.
        if self.used_subdoc {
            let mut next_id: u32 = 1;
            let fixed = fix_tables_docpr_cnvpr(
                &body_rendered,
                &mut self.docx_ids_index,
                Some(&mut next_id),
            )?;
            self.set_part_rendered(DOCUMENT_PART, fixed);
            let hf_parts: Vec<String> = [rel_type::HEADER, rel_type::FOOTER]
                .into_iter()
                .flat_map(|uri| self.header_footer_parts(uri))
                .collect();
            for part in hf_parts {
                let Some(xml) = self.pkg()?.get_string(&part) else {
                    continue;
                };
                let enc = self.pkg()?.encoding_of(&part);
                if let Some(fixed) = renumber_cnvpr(&xml, &mut next_id) {
                    let bytes = crate::package::encode_part(&fixed, &enc);
                    self.pkg()?.set(&part, bytes);
                }
            }
        } else {
            // fix tables and docPr ids on body only
            let fixed = fix_tables_and_docpr(&body_rendered, &mut self.docx_ids_index)?;
            self.set_part_rendered(DOCUMENT_PART, fixed);
        }

        self.is_rendered = true;
        Ok(())
    }

    fn render_init(&mut self) -> Result<(), String> {
        self.init_docx(true)?;
        self.pic_map.clear();
        self.docx_ids_index = 1000;
        self.is_saved = false;
        self.deferred.clear();
        self.deferred_by_oid.clear();
        self.used_subdoc = false;
        self.last_error_context.clear();
        self.next_shape_ids.clear();
        Ok(())
    }

    fn header_footer_parts(&mut self, uri: &str) -> Vec<String> {
        let rels = match &self.package {
            Some(p) => p.rels(DOCUMENT_PART),
            None => return Vec::new(),
        };
        rels.by_type(uri)
            .filter(|r| !r.is_external)
            .map(|r| resolve_target(DOCUMENT_PART, &r.target))
            .collect()
    }

    /// patch_xml + jinja render + resolve_listing for one part
    pub fn render_part(
        &mut self,
        part: &str,
        src_xml: &str,
        autoescape: bool,
        make_ctx: &CtxFn,
    ) -> Result<String, String> {
        let env = make_env(autoescape, self);
        self.render_part_with(part, src_xml, &env, make_ctx)
    }

    /// render_part with a pre-built environment (shared across all parts of
    /// one render call).
    fn render_part_with(
        &mut self,
        part: &str,
        src_xml: &str,
        env: &Environment,
        make_ctx: &CtxFn,
    ) -> Result<String, String> {
        let decoded = decode_text_entities(src_xml);
        let patched = patch_xml(&decoded);
        let prev = crate::pybridge::set_current_render(self as *mut TplCore, part);
        let result = (|| {
            let ctx = make_ctx(self, part)?;
            let dst = render_xml_str_with(&patched, ctx, env, self)?;
            let dst = resolve_listing(&dst);
            self.materialize_deferred(part, dst.into_owned())
        })();
        crate::pybridge::restore_current_render(prev);
        result
    }

    /// Replace deferred-value placeholders actually printed by the template
    /// with their materialized XML for the given part. Single pass over the
    /// xml; each deferred value is materialized at most once per part.
    pub fn materialize_deferred(&mut self, part: &str, xml: String) -> Result<String, String> {
        if self.deferred.is_empty() || !xml.contains('\u{1}') {
            return Ok(xml);
        }
        let mut done: Vec<Option<String>> = (0..self.deferred.len()).map(|_| None).collect();
        let mut out = String::with_capacity(xml.len());
        let mut rest = xml.as_str();
        while let Some(start) = rest.find('\u{1}') {
            out.push_str(&rest[..start]);
            let after = &rest[start + 1..];
            let parsed = after.find('\u{1}').and_then(|e| {
                after[..e]
                    .strip_prefix("DTPLD")
                    .and_then(|d| d.parse::<usize>().ok())
                    .map(|idx| (idx, e))
            });
            match parsed {
                Some((idx, end)) if idx < self.deferred.len() => {
                    if done[idx].is_none() {
                        done[idx] = Some(self.materialize_one(part, idx)?);
                    }
                    out.push_str(done[idx].as_ref().unwrap());
                    rest = &after[end + 1..];
                }
                _ => {
                    out.push('\u{1}');
                    rest = after;
                }
            }
        }
        out.push_str(rest);
        Ok(out)
    }

    /// Materialize one deferred value (cheap to call: the Deferred clone only
    /// bumps Arc refcounts).
    fn materialize_one(&mut self, part: &str, idx: usize) -> Result<String, String> {
        match self.deferred[idx].clone() {
            Deferred::Image {
                blob,
                filename,
                width,
                height,
                anchor,
                title,
                descr,
            } => crate::inline_image::inline_image_xml(
                self,
                part,
                &blob,
                filename.as_deref(),
                width,
                height,
                anchor.as_deref(),
                title.as_deref(),
                descr.as_deref(),
            ),
            Deferred::Subdoc { bytes } => {
                self.used_subdoc = true;
                match bytes {
                    Some(b) => crate::subdoc::subdoc_xml(self, &b),
                    None => Ok(String::new()),
                }
            }
            Deferred::SubdocBlocks { blocks } => {
                self.used_subdoc = true;
                crate::subdocbuilder::serialize_blocks(self, part, &blocks)
            }
        }
    }

    fn render_properties(&mut self, env: &Environment, make_ctx: &CtxFn) -> Result<(), String> {
        let name = "docProps/core.xml";
        // python-docx always provides a core properties part
        if self.package.as_ref().map(|p| !p.contains(name)).unwrap_or(false) {
            crate::docmodel::ensure_core_part(self.package.as_mut().unwrap());
        }
        let src = match self.pkg()?.get_string(name) {
            Some(s) => s,
            None => return Ok(()),
        };
        let mut doc = match Document::parse(&src) {
            Ok(d) => d,
            Err(_) => return Ok(()),
        };
        // python-docx core property -> xml tag
        let props = [
            "dc:creator",     // author
            "dc:description", // comments
            "dc:identifier",
            "dc:language",
            "dc:subject",
            "dc:title",
        ];
        let prev = crate::pybridge::set_current_render(self as *mut TplCore, name);
        let result = (|| -> Result<(), String> {
            let ctx = make_ctx(self, name)?;
            let mut changed = false;
            for tag in props {
                if let Some(el) = doc.root.find_mut(tag) {
                    let initial = el.text_content();
                    if initial.is_empty() {
                        continue;
                    }
                    let rendered = env
                        .template_from_str(&initial)
                        .and_then(|t| t.render(&ctx))
                        .map_err(|e| format!("error rendering core property {}: {}", tag, e))?;
                    let rendered = self.materialize_deferred(name, rendered)?;
                    el.children = vec![Node::Text(rendered)];
                    changed = true;
                }
            }
            if changed {
                let xml = doc.serialize();
                self.pkg()?.set(name, xml.into_bytes());
                self.invalidate_part(name);
            }
            Ok(())
        })();
        crate::pybridge::restore_current_render(prev);
        result
    }

    fn render_footnotes(&mut self, env: &Environment, make_ctx: &CtxFn) -> Result<(), String> {
        let rels = match &self.package {
            Some(p) => p.rels(DOCUMENT_PART),
            None => return Ok(()),
        };
        let parts: Vec<String> = rels
            .by_type(rel_type::FOOTNOTES)
            .filter(|r| !r.is_external)
            .map(|r| resolve_target(DOCUMENT_PART, &r.target))
            .collect();
        for part in parts {
            let src = match self.pkg()?.get_string(&part) {
                Some(s) => s,
                None => continue,
            };
            let rendered = self.render_part_with(&part, &src, env, make_ctx)?;
            self.set_part_rendered(&part, rendered);
        }
        Ok(())
    }

    /// Store a rendered XML part re-encoded with its declared encoding
    /// (docxtpl get_headers_footers_encoding parity).
    fn set_part_rendered(&mut self, part: &str, content: String) {
        let pkg = self.package.as_mut().expect("package loaded");
        let enc = pkg.encoding_of(part);
        pkg.set(part, crate::package::encode_part(&content, &enc));
        // rendered xml replaces whatever the part cache held
        self.invalidate_part(part);
    }

    // ---------------- variables ----------------

    pub fn undeclared_variables(&mut self, context_keys: Option<HashSet<String>>) -> Result<Vec<String>, String> {
        // Build on a temporary package so current state is untouched
        let pkg = Package::from_bytes(&self.original_bytes)?;
        let mut xml = String::new();
        if let Some(doc) = pkg.get_string(DOCUMENT_PART) {
            xml.push_str(&patch_xml(&decode_text_entities(&doc)));
        }
        let rels = pkg.rels(DOCUMENT_PART);
        for uri in [rel_type::HEADER, rel_type::FOOTER] {
            for r in rels.by_type(uri).filter(|r| !r.is_external) {
                let part = resolve_target(DOCUMENT_PART, &r.target);
                if let Some(src) = pkg.get_string(&part) {
                    xml.push_str(&patch_xml(&decode_text_entities(&src)));
                }
            }
        }
        let env = Environment::new();
        let tmpl = env
            .template_from_str(&xml)
            .map_err(|e| e.to_string())?;
        let mut vars = tmpl.undeclared_variables(false);
        if let Some(keys) = context_keys {
            vars = vars.into_iter().filter(|v| !keys.contains(v)).collect();
        }
        Ok(vars.into_iter().collect())
    }

    // ---------------- replacements ----------------

    pub fn reset_replacements(&mut self) {
        self.crc_to_new_media.clear();
        self.crc_to_new_embedded.clear();
        self.zipname_to_replace.clear();
        self.pics_to_replace.clear();
    }

    pub fn build_url_id(&mut self, url: &str) -> Result<String, String> {
        self.init_docx(false)?;
        Ok(self.pkg()?.add_rel(DOCUMENT_PART, rel_type::HYPERLINK, url, true))
    }

    fn pre_processing(&mut self) -> Result<(), String> {
        if !self.pics_to_replace.is_empty() {
            self.replace_pics()?;
        }
        Ok(())
    }

    fn replace_pics(&mut self) -> Result<(), String> {
        let mut replaced: HashMap<String, bool> = self
            .pics_to_replace
            .keys()
            .map(|k| (k.clone(), false))
            .collect();

        let mut parts = vec![DOCUMENT_PART.to_string()];
        for uri in [rel_type::HEADER, rel_type::FOOTER] {
            parts.extend(self.header_footer_parts(uri));
        }

        for part in parts {
            self.replace_part_pics(&part, &mut replaced)?;
        }

        if !self.allow_missing_pics {
            for (img_id, was_replaced) in &replaced {
                if !was_replaced {
                    return Err(format!(
                        "Picture {} not found in the docx template",
                        img_id
                    ));
                }
            }
        }
        Ok(())
    }

    fn replace_part_pics(
        &mut self,
        part: &str,
        replaced: &mut HashMap<String, bool>,
    ) -> Result<(), String> {
        let src = match self.pkg()?.get_string(part) {
            Some(s) => s,
            None => return Ok(()),
        };
        let doc = match Document::parse(&src) {
            Ok(d) => d,
            Err(_) => return Ok(()),
        };
        let rels = self.pkg()?.rels(part);

        // gather all a:graphicData elements
        let mut gds: Vec<&Element> = Vec::new();
        doc.root.iter_descendants("a:graphicData", &mut gds);

        let pic_uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        let mut updates: Vec<(String, Vec<u8>)> = Vec::new();

        for gd in gds {
            if gd.get_attr("uri") != Some(pic_uri) {
                continue;
            }
            let mut blips: Vec<&Element> = Vec::new();
            gd.iter_descendants("a:blip", &mut blips);
            let Some(blip) = blips.first() else { continue };
            let Some(rid) = blip.get_attr("r:embed") else { continue };

            let mut cnvprs: Vec<&Element> = Vec::new();
            gd.iter_descendants("pic:cNvPr", &mut cnvprs);
            let Some(cnvpr) = cnvprs.first() else { continue };
            let filename = cnvpr.get_attr("name").unwrap_or("").to_string();
            let title = cnvpr.get_attr("title").unwrap_or("").to_string();
            let descr = cnvpr.get_attr("descr").unwrap_or("").to_string();

            if let Some(rel) = rels.get(rid) {
                let target_part = resolve_target(part, &rel.target);
                self.pic_map
                    .insert(filename.clone(), (rel.target.clone(), target_part.clone()));
            } else {
                continue;
            }

            for (img_id, img_data) in &self.pics_to_replace {
                if *img_id == filename || *img_id == title || *img_id == descr {
                    let rel = rels.get(rid).unwrap();
                    let target_part = resolve_target(part, &rel.target);
                    updates.push((target_part, img_data.clone()));
                    if let Some(v) = replaced.get_mut(img_id) {
                        *v = true;
                    }
                    break;
                }
            }
        }

        for (part_name, data) in updates {
            self.pkg()?.set(&part_name, data);
        }
        Ok(())
    }

    // ---------------- save ----------------

    pub fn save_bytes(&mut self) -> Result<Vec<u8>, String> {
        self.flush_parts()?;
        // load the package only if nothing has touched it yet; reloading
        // unconditionally here would discard pre-render docmodel edits
        self.init_docx(false)?;
        self.pre_processing()?;

        // post processing: zip-level replacements (rare path: only clone the
        // whole package, media blobs included, when replacements are pending)
        let out = if !self.crc_to_new_media.is_empty()
            || !self.crc_to_new_embedded.is_empty()
            || !self.zipname_to_replace.is_empty()
        {
            let mut pkg = self
                .package
                .clone()
                .ok_or_else(|| "package not loaded".to_string())?;
            for (name, data, _compression, _mtime) in pkg.entries.iter_mut() {
                if let Some(new) = self.zipname_to_replace.get(name) {
                    *data = new.clone();
                } else if name.starts_with("word/media/") {
                    let c = crc32(data);
                    if let Some(new) = self.crc_to_new_media.get(&c) {
                        *data = new.clone();
                    }
                } else if name.starts_with("word/embeddings/") {
                    let c = crc32(data);
                    if let Some(new) = self.crc_to_new_embedded.get(&c) {
                        *data = new.clone();
                    }
                }
            }
            pkg.to_bytes()?
        } else {
            self.package
                .as_ref()
                .ok_or_else(|| "package not loaded".to_string())?
                .to_bytes()?
        };

        self.is_saved = true;
        Ok(out)
    }
}

// ---------------- jinja rendering of one part ----------------

pub fn make_env(autoescape: bool, core: &TplCore) -> Environment<'static> {
    let mut env = Environment::new();
    if autoescape {
        env.set_auto_escape_callback(|_| AutoEscape::Html);
    } else {
        env.set_auto_escape_callback(|_| AutoEscape::None);
    }
    if let Some(v) = core.env_options.trim_blocks {
        env.set_trim_blocks(v);
    }
    if let Some(v) = core.env_options.lstrip_blocks {
        env.set_lstrip_blocks(v);
    }
    if let Some(v) = core.env_options.keep_trailing_newline {
        env.set_keep_trailing_newline(v);
    }
    if let Some(b) = &core.env_options.undefined_behavior {
        env.set_undefined_behavior(match b.as_str() {
            "chainable" => minijinja::UndefinedBehavior::Chainable,
            "strict" => minijinja::UndefinedBehavior::Strict,
            "semistrict" => minijinja::UndefinedBehavior::SemiStrict,
            _ => minijinja::UndefinedBehavior::Lenient,
        });
    }
    // jinja2 exposes Python str methods on template strings; emulate the
    // common ones for minijinja values.
    env.set_unknown_method_callback(|_state, value, method, args| {
        py_like_method(value, method, args)
    });

    // jinja2's Undefined supports len() == 0
    env.add_filter("length", |value: Value| -> usize {
        use minijinja::value::ValueKind;
        match value.kind() {
            ValueKind::Undefined | ValueKind::None => 0,
            ValueKind::String | ValueKind::Bytes => value.to_string().chars().count(),
            _ => {
                if let Some(object) = value.as_object() {
                    if let Some(len) = object.enumerator_len() {
                        return len;
                    }
                }
                value
                    .try_iter()
                    .map(|it| it.count())
                    .unwrap_or(0)
            }
        }
    });

    // jinja2 filters missing from minijinja
    env.add_filter("striptags", |value: Value| -> String {
        let s = value.to_string();
        let mut out = String::with_capacity(s.len());
        let mut in_tag = false;
        for c in s.chars() {
            match c {
                '<' => in_tag = true,
                '>' => in_tag = false,
                _ if !in_tag => out.push(c),
                _ => {}
            }
        }
        // jinja2 also collapses whitespace
        out.split_whitespace().collect::<Vec<_>>().join(" ")
    });
    env.add_filter("pyformat", |args: minijinja::value::Rest<Value>| -> Result<String, minijinja::Error> {
        let Some((fmt, rest)) = args.0.split_first() else {
            return Ok(String::new());
        };
        percent_format_positional(&fmt.to_string(), rest)
    });
    env.add_filter("filesizeformat", |value: Value, binary: Option<Value>| {
        let bytes = f64::try_from(value.clone()).unwrap_or(0.0);
        let binary = binary.map(|b| b.is_true()).unwrap_or(false);
        let (base, units) = if binary {
            (1024.0_f64, ["Bytes", "KiB", "MiB", "GiB", "TiB", "PiB", "EiB", "ZiB", "YiB"])
        } else {
            (1000.0_f64, ["Bytes", "kB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"])
        };
        if bytes.abs() < base {
            return format!("{} Bytes", bytes as i64);
        }
        let mut value = bytes;
        let mut unit = units[0];
        for (i, prefix) in units.iter().enumerate().skip(1) {
            unit = prefix;
            value = bytes / base.powi(i as i32);
            if value.abs() < base {
                break;
            }
        }
        format!("{:.1} {}", value, unit)
    });
    env.add_filter("wordcount", |value: Value| -> usize {
        let s = value.to_string();
        s.split(|c: char| !c.is_alphanumeric() && c != '_')
            .filter(|w| !w.is_empty())
            .count()
    });
    env.add_filter("center", |value: Value, width: Option<Value>| {
        let s = value.to_string();
        let width: usize = width
            .and_then(|w| w.to_string().parse().ok())
            .unwrap_or(80);
        let len = s.chars().count();
        if len >= width {
            return s;
        }
        let total = width - len;
        let left = total / 2;
        format!("{}{}{}", " ".repeat(left), s, " ".repeat(total - left))
    });
    env.add_filter("forceescape", |value: Value| {
        let s = value.to_string();
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&#34;")
            .replace('\'', "&#39;")
    });
    env.add_filter("truncate", |args: minijinja::value::Rest<Value>| {
        let s = args.0.first().map(|v| v.to_string()).unwrap_or_default();
        let length: usize = args.0.get(1).and_then(|v| v.to_string().parse().ok()).unwrap_or(255);
        let killwords = args.0.get(2).map(|v| v.is_true()).unwrap_or(false);
        let end = args.0.get(3).map(|v| v.to_string()).unwrap_or_else(|| "...".to_string());
        let leeway: usize = args.0.get(4).and_then(|v| v.to_string().parse().ok()).unwrap_or(5);
        let len = s.chars().count();
        if len <= length + leeway {
            return s;
        }
        let cut: String = s.chars().take(length.saturating_sub(end.chars().count())).collect();
        if killwords {
            return format!("{}{}", cut, end);
        }
        match cut.rfind(|c: char| c.is_whitespace()) {
            Some(pos) if pos > 0 => format!("{}{}", &cut[..pos], end),
            _ => format!("{}{}", cut, end),
        }
    });
    env.add_filter("xmlattr", |value: Value, autospace: Option<Value>| {
        let autospace = autospace.map(|v| v.is_true()).unwrap_or(true);
        let mut items: Vec<String> = Vec::new();
        if let Ok(iter) = value.try_iter() {
            for k in iter {
                if let Ok(v) = value.get_item(&k) {
                    let vs = v
                        .to_string()
                        .replace('&', "&amp;")
                        .replace('<', "&lt;")
                        .replace('>', "&gt;")
                        .replace('"', "&#34;");
                    items.push(format!("{}=\"{}\"", k, vs));
                }
            }
        }
        let joined = items.join(" ");
        if autospace && !joined.is_empty() {
            format!("{} ", joined)
        } else {
            joined
        }
    });
    env.add_filter("wordwrap", |args: minijinja::value::Rest<Value>| {
        let s = args.0.first().map(|v| v.to_string()).unwrap_or_default();
        let width: usize = args.0.get(1).and_then(|v| v.to_string().parse().ok()).unwrap_or(79);
        let break_long = args.0.get(2).map(|v| v.is_true()).unwrap_or(true);
        let wrapstring = args.0.get(3).map(|v| v.to_string()).unwrap_or_else(|| "\n".to_string());
        wordwrap(&s, width.max(1), break_long, &wrapstring)
    });
    env.add_filter("urlize", |args: minijinja::value::Rest<Value>| {
        let s = args.0.first().map(|v| v.to_string()).unwrap_or_default();
        urlize(&s)
    });
    env.add_filter("random", |value: Value| -> Result<Value, minijinja::Error> {
        let items: Vec<Value> = value
            .try_iter()
            .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?
            .collect();
        if items.is_empty() {
            return Err(minijinja::Error::new(
                minijinja::ErrorKind::InvalidOperation,
                "random from empty sequence",
            ));
        }
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.subsec_nanos() as usize)
            .unwrap_or(0);
        Ok(items[nanos % items.len()].clone())
    });

    // bool/None tests must match our jinja2-faithful wrappers
    env.add_test("true", |value: Value| -> bool {
        if matches!(value.kind(), minijinja::value::ValueKind::Bool) {
            return value.is_true();
        }
        value
            .as_object()
            .and_then(|o| o.downcast_ref::<crate::pybridge::PyBoolObj>())
            .map(|b| b.0)
            .unwrap_or(false)
    });
    env.add_test("false", |value: Value| -> bool {
        if matches!(value.kind(), minijinja::value::ValueKind::Bool) {
            return !value.is_true();
        }
        value
            .as_object()
            .and_then(|o| o.downcast_ref::<crate::pybridge::PyBoolObj>())
            .map(|b| !b.0)
            .unwrap_or(false)
    });
    env.add_test("boolean", |value: Value| -> bool {
        matches!(value.kind(), minijinja::value::ValueKind::Bool)
            || value
                .as_object()
                .map(|o| o.downcast_ref::<crate::pybridge::PyBoolObj>().is_some())
                .unwrap_or(false)
    });
    env.add_test("none", |value: Value| -> bool {
        matches!(value.kind(), minijinja::value::ValueKind::None)
            || value
                .as_object()
                .map(|o| o.downcast_ref::<crate::pybridge::PyNoneObj>().is_some())
                .unwrap_or(false)
    });
    env.add_test("eq_true", |value: Value| -> bool {
        test_eq_true(&value)
    });
    env.add_test("eq_false", |value: Value| -> bool {
        test_eq_false(&value)
    });
    env.add_test("callable", |value: Value| -> bool {
        if let Some(o) = value.as_object() {
            if let Some(w) = o.downcast_ref::<crate::pybridge::PyWrapper>() {
                return pyo3::Python::attach(|py| w.obj.bind(py).is_callable());
            }
            // minijinja objects (macros/functions) are callable
            return true;
        }
        false
    });

    // jinja2's `mapping`/`sequence` tests must also match wrapped python
    // dicts/lists (they are objects for laziness)
    env.add_test("mapping", |value: Value| -> bool {
        if matches!(value.kind(), minijinja::value::ValueKind::Map) {
            return true;
        }
        if let Some(object) = value.as_object() {
            if let Some(wrapper) = object.downcast_ref::<crate::pybridge::PyWrapper>() {
                return pyo3::Python::attach(|py| {
                    wrapper.obj.bind(py).cast::<pyo3::types::PyDict>().is_ok()
                });
            }
        }
        false
    });
    env.add_test("sequence", |value: Value| -> bool {
        use minijinja::value::ValueKind;
        if matches!(value.kind(), ValueKind::Seq | ValueKind::String | ValueKind::Bytes) {
            return true;
        }
        if let Some(object) = value.as_object() {
            if let Some(wrapper) = object.downcast_ref::<crate::pybridge::PyWrapper>() {
                return pyo3::Python::attach(|py| {
                    let o = wrapper.obj.bind(py);
                    o.cast::<pyo3::types::PyList>().is_ok()
                        || o.cast::<pyo3::types::PyTuple>().is_ok()
                        || o.cast::<pyo3::types::PyString>().is_ok()
                });
            }
            return value.try_iter().is_ok();
        }
        false
    });

    // gettext functions backing {% trans %} (identity when no catalog)
    {
        let catalog = core.gettext_catalog.clone();
        env.add_function(
            "__dtpl_gettext",
            move |state: &minijinja::State, msgid: String| {
                let pattern = catalog
                    .as_ref()
                    .map(|c| c.gettext(&msgid))
                    .unwrap_or(msgid);
                percent_format(state, &pattern)
            },
        );
        let catalog = core.gettext_catalog.clone();
        env.add_function(
            "__dtpl_pgettext",
            move |state: &minijinja::State, context: String, msgid: String| {
                let pattern = catalog
                    .as_ref()
                    .map(|c| c.pgettext(&context, &msgid))
                    .unwrap_or(msgid);
                percent_format(state, &pattern)
            },
        );
        let catalog = core.gettext_catalog.clone();
        env.add_function(
            "__dtpl_ngettext",
            move |state: &minijinja::State, singular: String, plural: String, count: Value| {
                let n = i64::try_from(count).unwrap_or(1);
                let pattern = catalog
                    .as_ref()
                    .map(|c| c.ngettext(&singular, &plural, n))
                    .unwrap_or_else(|| if n == 1 { singular } else { plural });
                percent_format(state, &pattern)
            },
        );
        let catalog = core.gettext_catalog.clone();
        env.add_function(
            "__dtpl_npgettext",
            move |state: &minijinja::State, context: String, singular: String, plural: String, count: Value| {
                let n = i64::try_from(count).unwrap_or(1);
                let pattern = catalog
                    .as_ref()
                    .map(|c| c.npgettext(&context, &singular, &plural, n))
                    .unwrap_or_else(|| if n == 1 { singular } else { plural });
                percent_format(state, &pattern)
            },
        );
    }

    // user-registered filters / tests / functions / globals (python callables)
    for (name, callable) in &core.custom_filters {
        let callable = pyo3::Python::attach(|py| callable.clone_ref(py));
        env.add_filter(name.clone(), move |args: minijinja::value::Rest<Value>, kwargs: minijinja::value::Kwargs| {
            call_python_variadic(&callable, &args.0, &kwargs)
        });
    }
    for (name, callable) in &core.custom_tests {
        let callable = pyo3::Python::attach(|py| callable.clone_ref(py));
        env.add_test(name.clone(), move |args: minijinja::value::Rest<Value>, kwargs: minijinja::value::Kwargs| {
            call_python_variadic(&callable, &args.0, &kwargs)
                .map(|v| v.is_true())
        });
    }
    for (name, callable) in &core.custom_functions {
        let callable = pyo3::Python::attach(|py| callable.clone_ref(py));
        env.add_function(name.clone(), move |args: minijinja::value::Rest<Value>, kwargs: minijinja::value::Kwargs| {
            call_python_variadic(&callable, &args.0, &kwargs)
        });
    }
    for (name, value) in &core.custom_globals {
        let v = crate::pybridge::py_to_value_global(value);
        env.add_global(name.clone(), v);
    }
    if let Some(loader) = &core.template_loader {
        let loader = pyo3::Python::attach(|py| loader.clone_ref(py));
        env.set_loader(make_py_loader(loader));
    }
    env
}

fn make_py_loader(
    loader: pyo3::Py<pyo3::PyAny>,
) -> impl Fn(&str) -> Result<Option<String>, minijinja::Error> + Send + Sync + 'static {
    move |name| {
        let name = name.to_string();
        pyo3::Python::attach(|py| -> Result<Option<String>, minijinja::Error> {
            let result = loader
                .call1(py, (name,))
                .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?;
            if result.is_none(py) {
                return Ok(None);
            }
            let source = result
                .extract::<String>(py)
                .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?;
            Ok(Some(source))
        })
    }
}

fn wordwrap(s: &str, width: usize, break_long: bool, wrapstring: &str) -> String {
    let mut out = String::new();
    let mut col = 0usize;
    for word in s.split(' ') {
        let wlen = word.chars().count();
        if col > 0 && col + 1 + wlen > width {
            out.push_str(wrapstring);
            if wlen > width && break_long {
                // break the long word itself
                let mut rest: Vec<char> = word.chars().collect();
                while rest.len() > width {
                    let chunk: String = rest.drain(..width).collect();
                    out.push_str(&chunk);
                    out.push_str(wrapstring);
                }
                out.push_str(&rest.iter().collect::<String>());
                col = rest.len();
                continue;
            }
            out.push_str(word);
            col = wlen;
            continue;
        }
        if col > 0 {
            out.push(' ');
            col += 1;
        }
        out.push_str(word);
        col += wlen;
    }
    out
}

fn urlize(s: &str) -> String {
    // link http(s):// and www. urls like jinja2's urlize (simplified policies)
    crate::patch::sub(
        r#"(https?://[^\s<>'"]+|www\.[^\s<>'"]+)"#,
        |m| {
            let url = m.get(1).unwrap().as_str();
            // strip trailing punctuation like jinja2 policies
            let (mut end, mut trail) = (url.len(), "");
            while let Some(c) = url[..end].chars().last() {
                if ".:;!?,)".contains(c) {
                    end -= c.len_utf8();
                    trail = &url[end..];
                } else {
                    break;
                }
            }
            let clean = &url[..end];
            let href = if clean.starts_with("www.") {
                format!("http://{}", clean)
            } else {
                clean.to_string()
            };
            format!(
                "<a href=\"{}\" rel=\"noopener\">{}</a>{}",
                href, clean, trail
            )
        },
        s,
    )
}

fn call_python_variadic(
    callable: &pyo3::Py<pyo3::PyAny>,
    args: &[Value],
    kwargs: &minijinja::value::Kwargs,
) -> Result<Value, minijinja::Error> {
    pyo3::Python::attach(|py| {
        let mut py_args = Vec::with_capacity(args.len());
        for a in args {
            py_args.push(
                crate::pybridge::value_to_py(py, a)
                    .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?,
            );
        }
        let tuple = pyo3::types::PyTuple::new(py, py_args)
            .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?;
        // keyword arguments -> PyDict
        let dict = pyo3::types::PyDict::new(py);
        for key in kwargs.args() {
            let v: Value = kwargs.get(key).map_err(|e| e)?;
            dict.set_item(
                key,
                crate::pybridge::value_to_py(py, &v).map_err(|e| {
                    minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string())
                })?,
            )
            .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?;
        }
        let result = callable
            .call(py, tuple, Some(&dict))
            .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))?;
        let result = result.bind(py);
        crate::pybridge::py_to_value_render(&result)
            .map_err(|e| minijinja::Error::new(minijinja::ErrorKind::InvalidOperation, e.to_string()))
    })
}

/// Python str.format() with format spec support
/// ([[fill]align][sign][#][0][width][,|_][.precision][type])
fn py_format(template: &str, args: &[Value]) -> String {
    let mut result = String::new();
    let mut rest = template;
    let mut auto_idx = 0usize;
    while let Some(pos) = rest.find(['{', '}']) {
        let ch = rest[pos..].chars().next().unwrap();
        result.push_str(&rest[..pos]);
        if ch == '}' {
            if rest[pos..].starts_with("}}") {
                result.push('}');
                rest = &rest[pos + 2..];
            } else {
                result.push('}');
                rest = &rest[pos + 1..];
            }
            continue;
        }
        if rest[pos..].starts_with("{{") {
            result.push('{');
            rest = &rest[pos + 2..];
            continue;
        }
        let end = rest[pos..].find('}').map(|e| pos + e).unwrap_or(rest.len());
        let field_spec = &rest[pos + 1..end];
        rest = &rest[(end + 1).min(rest.len())..];
        // split field / conversion / spec
        let (field_conv, spec) = match field_spec.split_once(':') {
            Some((f, sp)) => (f, Some(sp)),
            None => (field_spec, None),
        };
        let (field, repr_mode) = match field_conv.strip_suffix("!r") {
            Some(f) => (f, true),
            None => (field_conv.strip_suffix("!s").unwrap_or(field_conv.strip_suffix("!a").unwrap_or(field_conv)), false),
        };
        let idx: usize = if field.is_empty() {
            let i = auto_idx;
            auto_idx += 1;
            i
        } else {
            field.parse().unwrap_or(0)
        };
        let Some(value) = args.get(idx) else {
            continue;
        };
        let mut text = if repr_mode {
            match value.kind() {
                minijinja::value::ValueKind::String => format!("'{}'", value),
                _ => value.to_string(),
            }
        } else {
            value.to_string()
        };
        if let Some(spec) = spec {
            text = apply_format_spec(&text, value, spec);
        }
        result.push_str(&text);
    }
    result.push_str(rest);
    result
}

fn apply_format_spec(text: &str, value: &Value, spec: &str) -> String {
    let chars: Vec<char> = spec.chars().collect();
    let mut i = 0usize;
    let mut fill = ' ';
    let mut align = '\0';
    let mut sign = '-';
    let mut alt = false;
    let mut zero = false;
    let mut width: Option<usize> = None;
    let mut grouping = '\0';
    let mut precision: Option<usize> = None;
    let mut ty = '\0';

    if chars.len() >= 2 && ["<", ">", "=", "^"].contains(&chars[1].to_string().as_str()) {
        fill = chars[0];
        align = chars[1];
        i = 2;
    } else if !chars.is_empty() && ["<", ">", "=", "^"].contains(&chars[0].to_string().as_str()) {
        align = chars[0];
        i = 1;
    }
    if i < chars.len() && (chars[i] == '+' || chars[i] == '-' || chars[i] == ' ') {
        sign = chars[i];
        i += 1;
    }
    if i < chars.len() && chars[i] == '#' {
        alt = true;
        i += 1;
    }
    if i < chars.len() && chars[i] == '0' {
        zero = true;
        i += 1;
    }
    let width_start = i;
    while i < chars.len() && chars[i].is_ascii_digit() {
        i += 1;
    }
    if i > width_start {
        width = chars[width_start..i].iter().collect::<String>().parse().ok();
    }
    if i < chars.len() && (chars[i] == ',' || chars[i] == '_') {
        grouping = chars[i];
        i += 1;
    }
    if i < chars.len() && chars[i] == '.' {
        i += 1;
        let p_start = i;
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }
        precision = chars[p_start..i].iter().collect::<String>().parse().ok();
    }
    if i < chars.len() {
        ty = chars[i];
    }

    let is_neg = text.starts_with('-');
    let mut body = text.trim_start_matches(['-', '+']).to_string();

    // numeric reformatting
    let mut prefix = String::new();
    match ty {
        'd' | 'n' | '\0' if grouping != '\0' => {
            body = group_digits(&body, grouping);
        }
        'b' => {
            if let Some(v) = value.as_i64() {
                body = format!("{:b}", v);
                if alt { prefix = "0b".into(); }
            }
        }
        'o' => {
            if let Some(v) = value.as_i64() {
                body = format!("{:o}", v);
                if alt { prefix = "0o".into(); }
            }
        }
        'x' => {
            if let Some(v) = value.as_i64() {
                body = format!("{:x}", v);
                if alt { prefix = "0x".into(); }
            }
        }
        'X' => {
            if let Some(v) = value.as_i64() {
                body = format!("{:X}", v);
                if alt { prefix = "0X".into(); }
            }
        }
        'e' | 'E' => {
            if let Ok(v) = f64::try_from(value.clone()) {
                let p = precision.unwrap_or(6);
                body = py_exp_format(format!("{:.*e}", p, v));
                if ty == 'E' { body = body.to_uppercase(); }
            }
        }
        'f' | 'F' => {
            if let Ok(v) = f64::try_from(value.clone()) {
                let p = precision.unwrap_or(6);
                body = format!("{:.*}", p, v.abs());
                if grouping != '\0' { body = group_digits(&body, grouping); }
            }
        }
        '%' => {
            if let Ok(v) = f64::try_from(value.clone()) {
                let p = precision.unwrap_or(6);
                body = format!("{:.*}%", p, v.abs() * 100.0);
            }
        }
        'g' | 'G' => {
            if let Ok(v) = f64::try_from(value.clone()) {
                let p = precision.unwrap_or(6);
                body = format_g(v, p);
                if ty == 'G' { body = body.to_uppercase(); }
            }
        }
        'c' => {
            if let Some(v) = value.as_i64() {
                if let Some(c) = char::from_u32(v as u32) { body = c.to_string(); }
            }
        }
        's' | '\0' => {
            if let Some(p) = precision {
                body = body.chars().take(p).collect();
            }
        }
        _ => {}
    }
    if ty == 'd' && grouping != '\0' {
        // already handled above
    }

    let sign_str = if is_neg {
        "-"
    } else {
        match sign {
            '+' => "+",
            ' ' => " ",
            _ => "",
        }
    };

    let default_align = if matches!(value.kind(), minijinja::value::ValueKind::Number) { '>' } else { '<' };
    let align = if align == '\0' {
        if zero { '=' } else { default_align }
    } else {
        align
    };
    if zero && fill == ' ' {
        fill = '0';
    }

    let core = format!("{}{}{}", sign_str, prefix, body);
    let w = width.unwrap_or(0);
    let len = core.chars().count();
    if len >= w {
        return core;
    }
    let pad_total = w - len;
    match align {
        '>' => format!("{}{}", fill.to_string().repeat(pad_total), core),
        '^' => {
            let left = pad_total / 2;
            let right = pad_total - left;
            format!("{}{}{}", fill.to_string().repeat(left), core, fill.to_string().repeat(right))
        }
        '=' => format!("{}{}{}{}", sign_str, prefix, fill.to_string().repeat(pad_total), body),
        _ => format!("{}{}", core, fill.to_string().repeat(pad_total)),
    }
}

/// rust "1.2345e4" -> python "1.2345e+04"
fn py_exp_format(s: String) -> String {
    if let Some(epos) = s.find('e') {
        let (mantissa, exp) = (&s[..epos], &s[epos + 1..]);
        let (esign, edigits) = if let Some(d) = exp.strip_prefix('-') {
            ("-", d)
        } else {
            ("+", exp.trim_start_matches('+'))
        };
        return format!("{}e{}{:0>2}", mantissa, esign, edigits);
    }
    s
}

fn group_digits(s: &str, grouping: char) -> String {
    let (int_part, rest) = match s.find(['.', 'e', 'E']) {
        Some(p) => (&s[..p], &s[p..]),
        None => (s, ""),
    };
    let digits: Vec<char> = int_part.chars().collect();
    let mut out = String::new();
    for (i, c) in digits.iter().enumerate() {
        if i > 0 && (digits.len() - i) % 3 == 0 {
            out.push(grouping);
        }
        out.push(*c);
    }
    out.push_str(rest);
    out
}

fn format_g(v: f64, precision: usize) -> String {
    if v == 0.0 {
        return "0".to_string();
    }
    let exp = v.abs().log10().floor() as i32;
    if exp < -4 || exp >= precision as i32 {
        let s = format!("{:.*e}", precision.saturating_sub(1), v);
        // trim trailing zeros in mantissa
        if let Some(epos) = s.find('e') {
            let mantissa = s[..epos].trim_end_matches('0').trim_end_matches('.');
            return format!("{}{}", mantissa, &s[epos..]);
        }
        s
    } else {
        let decimals = (precision as i32 - 1 - exp).max(0) as usize;
        let s = format!("{:.*}", decimals, v);
        s.trim_end_matches('0').trim_end_matches('.').to_string()
    }
}

/// Python-like str methods for jinja compatibility (called for unknown methods).
fn py_like_method(
    value: &Value,
    method: &str,
    args: &[Value],
) -> Result<Value, minijinja::Error> {
    use minijinja::{Error, ErrorKind};
    let unsupported = || {
        Err(Error::new(
            ErrorKind::UnknownMethod,
            format!("object has no method named {}", method),
        ))
    };
    // dict methods
    if matches!(value.kind(), minijinja::value::ValueKind::Map) {
        match method {
            "keys" | "values" | "items" => {
                let iter = value.try_iter().map_err(|e| {
                    Error::new(ErrorKind::InvalidOperation, e.to_string())
                })?;
                let pairs: Vec<Value> = iter.collect();
                return Ok(match method {
                    "keys" => Value::from_serialize(&pairs),
                    "values" => {
                        let vals: Vec<Value> = pairs
                            .iter()
                            .filter_map(|k| value.get_item(k).ok())
                            .collect();
                        Value::from_serialize(&vals)
                    }
                    _ => {
                        let items: Vec<Value> = pairs
                            .iter()
                            .filter_map(|k| {
                                value
                                    .get_item(k)
                                    .ok()
                                    .map(|v| Value::from_serialize(&vec![k.clone(), v]))
                            })
                            .collect();
                        Value::from_serialize(&items)
                    }
                });
            }
            "get" => {
                let key = args.first().cloned().unwrap_or(Value::UNDEFINED);
                let default = args.get(1).cloned().unwrap_or(Value::from(()));
                return match value.get_item(&key) {
                    Ok(v) if !v.is_undefined() => Ok(v),
                    _ => Ok(default),
                };
            }
            _ => return unsupported(),
        }
    }
    let Some(s) = value.as_str().map(|s| s.to_string()) else {
        return unsupported();
    };
    let arg_str = |i: usize| -> String {
        args.get(i).map(|v| v.to_string()).unwrap_or_default()
    };
    let out: Value = match method {
        "upper" => Value::from(s.to_uppercase()),
        "lower" => Value::from(s.to_lowercase()),
        "capitalize" => {
            let mut chars = s.chars();
            let first = chars.next().map(|c| c.to_uppercase().collect::<String>());
            Value::from(format!(
                "{}{}",
                first.unwrap_or_default(),
                chars.as_str().to_lowercase()
            ))
        }
        "title" => Value::from(minijinja::filters::title(std::borrow::Cow::Borrowed(
            s.as_str(),
        ))),
        "strip" => Value::from(s.trim().to_string()),
        "lstrip" => Value::from(s.trim_start().to_string()),
        "rstrip" => Value::from(s.trim_end().to_string()),
        "replace" => Value::from(s.replace(&arg_str(0), &arg_str(1))),
        "startswith" => Value::from(s.starts_with(&arg_str(0))),
        "endswith" => Value::from(s.ends_with(&arg_str(0))),
        "split" => {
            let sep = args.first();
            let parts: Vec<String> = match sep {
                Some(v) if !v.is_none() && !v.is_undefined() => {
                    s.split(&v.to_string()).map(|p| p.to_string()).collect()
                }
                _ => s.split_whitespace().map(|p| p.to_string()).collect(),
            };
            Value::from_serialize(&parts)
        }
        "rsplit" => {
            let sep = args.first();
            let parts: Vec<String> = match sep {
                Some(v) if !v.is_none() && !v.is_undefined() => {
                    let mut p: Vec<String> =
                        s.rsplit(&v.to_string()).map(|p| p.to_string()).collect();
                    p.reverse();
                    p
                }
                _ => s.split_whitespace().map(|p| p.to_string()).collect(),
            };
            Value::from_serialize(&parts)
        }
        "join" => {
            let sep = s.clone();
            let mut parts: Vec<String> = Vec::new();
            if let Some(seq) = args.first() {
                if let Ok(iter) = seq.try_iter() {
                    for item in iter {
                        parts.push(item.to_string());
                    }
                }
            }
            Value::from(parts.join(&sep))
        }
        "zfill" => {
            let width: usize = arg_str(0).parse().unwrap_or(0);
            let (sign, digits) = if let Some(rest) = s.strip_prefix(['-', '+']) {
                (&s[..1], rest)
            } else {
                ("", s.as_str())
            };
            let pad = width.saturating_sub(sign.len() + digits.len());
            Value::from(format!("{}{}{}", sign, "0".repeat(pad), digits))
        }
        "ljust" => {
            let width: usize = arg_str(0).parse().unwrap_or(0);
            let fill = arg_str(1).chars().next().unwrap_or(' ');
            let pad = width.saturating_sub(s.chars().count());
            Value::from(format!("{}{}", s, fill.to_string().repeat(pad)))
        }
        "rjust" => {
            let width: usize = arg_str(0).parse().unwrap_or(0);
            let fill = arg_str(1).chars().next().unwrap_or(' ');
            let pad = width.saturating_sub(s.chars().count());
            Value::from(format!("{}{}", fill.to_string().repeat(pad), s))
        }
        "center" => {
            let width: usize = arg_str(0).parse().unwrap_or(0);
            let total = width.saturating_sub(s.chars().count());
            let left = total / 2;
            let right = total - left;
            Value::from(format!(
                "{}{}{}",
                " ".repeat(left),
                s,
                " ".repeat(right)
            ))
        }
        "count" => Value::from(s.matches(&arg_str(0)).count() as i64),
        "find" => Value::from(
            s.find(&arg_str(0)).map(|i| i as i64).unwrap_or(-1),
        ),
        "index" => match s.find(&arg_str(0)) {
            Some(i) => Value::from(i as i64),
            None => {
                return Err(Error::new(
                    ErrorKind::InvalidOperation,
                    "substring not found",
                ))
            }
        },
        "format" => Value::from(py_format(&s, args)),
        "isdigit" => Value::from(!s.is_empty() && s.chars().all(|c| c.is_ascii_digit())),
        "isalpha" => Value::from(!s.is_empty() && s.chars().all(|c| c.is_alphabetic())),
        "isalnum" => Value::from(!s.is_empty() && s.chars().all(|c| c.is_alphanumeric())),
        "isspace" => Value::from(!s.is_empty() && s.chars().all(|c| c.is_whitespace())),
        "isupper" => Value::from(s.chars().any(|c| c.is_uppercase()) && !s.chars().any(|c| c.is_lowercase())),
        "islower" => Value::from(s.chars().any(|c| c.is_lowercase()) && !s.chars().any(|c| c.is_uppercase())),
        "isnumeric" => Value::from(!s.is_empty() && s.chars().all(|c| c.is_numeric())),
        "isdecimal" => Value::from(!s.is_empty() && s.chars().all(|c| c.is_ascii_digit())),
        "istitle" => {
            let mut prev_cased = false;
            let mut ok = !s.is_empty();
            for c in s.chars() {
                let cased = c.is_alphabetic();
                if cased {
                    if prev_cased && c.is_uppercase() {
                        ok = false;
                        break;
                    }
                    if !prev_cased && c.is_lowercase() {
                        ok = false;
                        break;
                    }
                }
                prev_cased = cased;
            }
            Value::from(ok)
        }
        "isidentifier" => {
            let mut chars = s.chars();
            let first_ok = chars.next().map(|c| c.is_alphabetic() || c == '_').unwrap_or(false);
            Value::from(first_ok && chars.all(|c| c.is_alphanumeric() || c == '_'))
        }
        "isprintable" => Value::from(s.chars().all(|c| !c.is_control())),
        "swapcase" => Value::from(
            s.chars()
                .map(|c| {
                    if c.is_uppercase() {
                        c.to_lowercase().collect::<String>()
                    } else if c.is_lowercase() {
                        c.to_uppercase().collect::<String>()
                    } else {
                        c.to_string()
                    }
                })
                .collect::<String>(),
        ),
        "casefold" => Value::from(s.to_lowercase()),
        "expandtabs" => {
            let tabsize: usize = args.first().and_then(|v| v.to_string().parse().ok()).unwrap_or(8);
            let mut out = String::new();
            let mut col = 0usize;
            for c in s.chars() {
                match c {
                    '\t' => {
                        let spaces = tabsize - (col % tabsize.max(1));
                        out.push_str(&" ".repeat(spaces));
                        col += spaces;
                    }
                    '\n' | '\r' => {
                        out.push(c);
                        col = 0;
                    }
                    _ => {
                        out.push(c);
                        col += 1;
                    }
                }
            }
            Value::from(out)
        }
        "splitlines" => {
            let keepends = args.first().map(|v| v.is_true()).unwrap_or(false);
            let mut parts: Vec<String> = Vec::new();
            let mut cur = String::new();
            let mut chars = s.chars().peekable();
            while let Some(c) = chars.next() {
                let is_break = c == '\n' || c == '\r';
                if is_break {
                    if c == '\r' && chars.peek() == Some(&'\n') {
                        chars.next();
                        if keepends {
                            cur.push('\r');
                            cur.push('\n');
                        }
                    } else if keepends {
                        cur.push(c);
                    }
                    parts.push(std::mem::take(&mut cur));
                } else {
                    cur.push(c);
                }
            }
            if !cur.is_empty() {
                parts.push(cur);
            }
            Value::from_serialize(&parts)
        }
        "partition" => {
            let sep = arg_str(0);
            let (before, found, after) = match s.find(&sep) {
                Some(i) if !sep.is_empty() => (s[..i].to_string(), sep.clone(), s[i + sep.len()..].to_string()),
                _ => (s.clone(), String::new(), String::new()),
            };
            Value::from_serialize(&vec![before, found, after])
        }
        "rpartition" => {
            let sep = arg_str(0);
            let (before, found, after) = match s.rfind(&sep) {
                Some(i) if !sep.is_empty() => (s[..i].to_string(), sep.clone(), s[i + sep.len()..].to_string()),
                _ => (String::new(), String::new(), s.clone()),
            };
            Value::from_serialize(&vec![before, found, after])
        }
        "removeprefix" => {
            let p = arg_str(0);
            Value::from(s.strip_prefix(&p).unwrap_or(&s).to_string())
        }
        "removesuffix" => {
            let p = arg_str(0);
            Value::from(s.strip_suffix(&p).unwrap_or(&s).to_string())
        }
        "encode" => Value::from(s.clone()),
        _ => return unsupported(),
    };
    Ok(out)
}

/// Rewrite `{% trans %}` blocks into gettext function calls (jinja2 i18n
/// extension semantics). Without an installed catalog the functions fall
/// back to the untranslated, formatted msgid.
pub fn preprocess_trans(src: &str) -> Cow<'_, str> {
    if !src.contains("{% trans") {
        return Cow::Borrowed(src);
    }
    Cow::Owned(crate::patch::sub(
        r"(?s)\{%\s*trans\s*(.*?)%\}(.*?)(?:\{%\s*pluralize\s*(.*?)\s*%\}(.*?))?\{%\s*endtrans\s*%\}",
        |m| trans_rewrite(
            m.get(1).unwrap().as_str(),
            m.get(2).unwrap().as_str(),
            m.get(3).map(|g| g.as_str()),
            m.get(4).map(|g| g.as_str()),
        ),
        src,
    ))
}

fn trans_rewrite(
    assigns: &str,
    singular: &str,
    plz_expr: Option<&str>,
    plural: Option<&str>,
) -> String {
    // parse k=v assignments; context="x" selects pgettext
    let kv_re = crate::patch::re(r#"(\w+)\s*=\s*("(?:[^"]*)"|'(?:[^']*)'|[^,\s]+)"#);
    let mut sets = String::new();
    let mut context: Option<String> = None;
    let mut first_var: Option<String> = None;
    for cap in kv_re.captures_iter(assigns).flatten() {
        let key = cap[1].to_string();
        let value = cap[2].to_string();
        if key == "context" {
            context = Some(value.trim_matches(['"', '\'']).to_string());
            continue;
        }
        if first_var.is_none() {
            first_var = Some(key.clone());
        }
        sets.push_str(&format!("{{% set {} = {} %}}", key, value));
    }

    let escape_jinja_str = |s: &str| -> String {
        s.replace('\\', "\\\\").replace('"', "\\\"").replace('\n', "\\n")
    };
    let to_msgid = |body: &str| -> String {
        let interpolated = crate::patch::sub(
            r"\{\{\s*(.*?)\s*\}\}",
            |m| format!("%({})s", m.get(1).unwrap().as_str().trim()),
            body,
        );
        escape_jinja_str(&interpolated)
    };

    let singular_id = to_msgid(singular);
    match plural {
        Some(plural_body) => {
            let plural_id = to_msgid(plural_body);
            let count = plz_expr
                .map(|e| e.trim().to_string())
                .filter(|e| !e.is_empty())
                .or(first_var)
                .unwrap_or_else(|| "1".to_string());
            let func = if context.is_some() {
                "__dtpl_npgettext"
            } else {
                "__dtpl_ngettext"
            };
            let ctx_arg = context
                .map(|c| format!("\"{}\", ", escape_jinja_str(&c)))
                .unwrap_or_default();
            format!(
                "{}{{{{ {}({}\"{}\", \"{}\", ({})) }}}}",
                sets, func, ctx_arg, singular_id, plural_id, count
            )
        }
        None => {
            let (func, ctx_arg) = match &context {
                Some(c) => (
                    "__dtpl_pgettext",
                    format!("\"{}\", ", escape_jinja_str(c)),
                ),
                None => ("__dtpl_gettext", String::new()),
            };
            format!("{}{{{{ {}({}\"{}\") }}}}", sets, func, ctx_arg, singular_id)
        }
    }
}

/// Python %-style formatting against the template state
/// (supports %(name)s/%d/%f/%x/%% with flags, width, precision).
fn percent_format(state: &minijinja::State, pattern: &str) -> Result<String, minijinja::Error> {
    let mut out = String::new();
    let mut rest = pattern;
    while let Some(pos) = rest.find('%') {
        out.push_str(&rest[..pos]);
        let tail = &rest[pos + 1..];
        if let Some(stripped) = tail.strip_prefix('%') {
            out.push('%');
            rest = stripped;
            continue;
        }
        let Some(after) = tail.strip_prefix('(') else {
            out.push('%');
            rest = tail;
            continue;
        };
        let Some(end) = after.find(')') else {
            out.push('%');
            rest = tail;
            continue;
        };
        let name = &after[..end];
        let mut spec_part = &after[end + 1..];
        // parse %-spec: [-+ #0]*[width][.precision]type
        let mut align_left = false;
        let mut sign = "";
        let mut zero = false;
        let mut alt = false;
        loop {
            match spec_part.chars().next() {
                Some('-') => align_left = true,
                Some('+') => sign = "+",
                Some(' ') => sign = " ",
                Some('0') => zero = true,
                Some('#') => alt = true,
                _ => break,
            }
            spec_part = &spec_part[1..];
        }
        let mut width = String::new();
        while spec_part.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false) {
            width.push(spec_part.chars().next().unwrap());
            spec_part = &spec_part[1..];
        }
        let mut precision = String::new();
        if spec_part.starts_with('.') {
            spec_part = &spec_part[1..];
            while spec_part.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false) {
                precision.push(spec_part.chars().next().unwrap());
                spec_part = &spec_part[1..];
            }
        }
        let ty = spec_part.chars().next().unwrap_or('s');
        spec_part = &spec_part[1..];
        rest = spec_part;

        let value = state
            .lookup(name)
            .unwrap_or_else(|| Value::from(()));
        // build a {}-style spec for apply_format_spec
        let mut spec = String::new();
        if align_left {
            spec.push('<');
        }
        spec.push_str(sign);
        if alt {
            spec.push('#');
        }
        if zero {
            spec.push('0');
        }
        spec.push_str(&width);
        if !precision.is_empty() {
            spec.push('.');
            spec.push_str(&precision);
        }
        spec.push(ty);
        let text = value.to_string();
        out.push_str(&apply_format_spec(&text, &value, &spec));
    }
    out.push_str(rest);
    Ok(out)
}

/// Rewrite `== true` / `!= false` style comparisons to custom tests with
/// exact Python equality semantics (booleans are wrapped for Python-faithful
/// display, so minijinja's native `==` cannot compare them with `true`).
///
/// The rewrite is scoped to jinja tag contents and skips quoted strings, so
/// literal document text and string literals are never touched.
pub fn preprocess_bool_compare(src: &str) -> Cow<'_, str> {
    if !src.contains("true") && !src.contains("false") && !src.contains("none")
        && !src.contains("True") && !src.contains("False") && !src.contains("None")
    {
        return Cow::Borrowed(src);
    }
    const RULES: &[(&str, &str)] = &[
        ("== true", "is eq_true"),
        ("!= true", "is not eq_true"),
        ("== false", "is eq_false"),
        ("!= false", "is not eq_false"),
        ("== True", "is eq_true"),
        ("!= True", "is not eq_true"),
        ("== False", "is eq_false"),
        ("!= False", "is not eq_false"),
        ("== none", "is none"),
        ("!= none", "is not none"),
        ("== None", "is none"),
        ("!= None", "is not none"),
    ];

    // process only the inside of {% %} / {{ }} blocks (linear tag scan;
    // borrowed when no tag actually changed)
    rewrite_jinja_tags(src, |tag| {
        // cheap prefilter: every rule starts with == or !=
        if !tag.contains("==") && !tag.contains("!=") {
            return None;
        }
        let rewritten = rewrite_tag_bool_compare(tag, RULES);
        if rewritten != tag {
            Some(rewritten)
        } else {
            None
        }
    })
}

fn rewrite_tag_bool_compare(tag: &str, rules: &[(&str, &str)]) -> String {
    // walk the tag, rewriting only outside quoted regions (char-safe)
    let mut out = String::with_capacity(tag.len());
    let mut i = 0usize;
    while i < tag.len() {
        let c = tag[i..].chars().next().unwrap();
        if c == '"' || c == '\'' {
            // copy the quoted region verbatim
            let start = i;
            i += c.len_utf8();
            while i < tag.len() {
                let d = tag[i..].chars().next().unwrap();
                if d == '\\' {
                    i += d.len_utf8();
                    if i < tag.len() {
                        i += tag[i..].chars().next().unwrap().len_utf8();
                    }
                    continue;
                }
                i += d.len_utf8();
                if d == c {
                    break;
                }
            }
            out.push_str(&tag[start..i]);
            continue;
        }
        // try to match a rule at this position
        let mut matched = false;
        for (pat, rep) in rules {
            if tag[i..].starts_with(pat) {
                // boundary check: the char after the pattern must not be an
                // identifier char (avoid matching e.g. "== trueish")
                let after = tag[i + pat.len()..].chars().next();
                if after.map(|c| c.is_alphanumeric() || c == '_').unwrap_or(true) {
                    continue;
                }
                out.push_str(rep);
                i += pat.len();
                matched = true;
                break;
            }
        }
        if !matched {
            out.push(c);
            i += c.len_utf8();
        }
    }
    out
}

/// Engine-feature preprocessing applied to jinja tag contents:
/// `{% debug %}` -> `{{ debug() }}`, and printf-style `'fmt' % args` ->
/// `('fmt')|pyformat(args)` (minijinja has no % operator for strings).
pub fn preprocess_engine_features(src: &str) -> Cow<'_, str> {
    let needs_debug = src.contains("{% debug");
    let needs_printf = src.contains('%');
    if !needs_debug && !needs_printf {
        return Cow::Borrowed(src);
    }
    rewrite_jinja_tags(src, |tag| {
        let mut out: Cow<'_, str> = Cow::Borrowed(tag);
        if needs_debug && tag.contains("debug") {
            out = Cow::Owned(crate::patch::sub_str(
                r"\{%\s*debug\s*%\}",
                "{{ debug() }}",
                &out,
            ));
        }
        if needs_printf && (out.contains('%') && (out.contains('\'') || out.contains('"'))) {
            out = Cow::Owned(rewrite_printf_tags(&out));
        }
        match out {
            Cow::Owned(o) if o != tag => Some(o),
            _ => None,
        }
    })
}

/// Rewrite `STR % RHS` inside a tag to `(STR)|pyformat(RHS)` (char-safe).
fn rewrite_printf_tags(tag: &str) -> String {
    let mut out = String::with_capacity(tag.len());
    let mut i = 0usize;
    while i < tag.len() {
        let c = tag[i..].chars().next().unwrap();
        if c == '"' || c == '\'' {
            // capture the string literal
            let start = i;
            i += c.len_utf8();
            while i < tag.len() {
                let d = tag[i..].chars().next().unwrap();
                if d == '\\' {
                    i += d.len_utf8();
                    if i < tag.len() {
                        i += tag[i..].chars().next().unwrap().len_utf8();
                    }
                    continue;
                }
                i += d.len_utf8();
                if d == c {
                    break;
                }
            }
            let literal = &tag[start..i];
            // check for ` % rhs` after the literal
            let mut j = i;
            while j < tag.len() && tag[j..].chars().next().unwrap().is_whitespace() {
                j += tag[j..].chars().next().unwrap().len_utf8();
            }
            let has_printf = tag[j..].starts_with('%')
                && tag[j + 1..]
                    .chars()
                    .next()
                    .map(|c| c.is_whitespace() || c == '(' || c.is_alphabetic() || c == '_' || c.is_ascii_digit())
                    .unwrap_or(false);
            if has_printf {
                let mut k = j + 1;
                while k < tag.len() && tag[k..].chars().next().unwrap().is_whitespace() {
                    k += tag[k..].chars().next().unwrap().len_utf8();
                }
                // capture RHS: balanced parens or an atom
                let (rhs, end) = capture_rhs(tag, k);
                if !rhs.is_empty() {
                    out.push_str(&format!("({})|pyformat({})", literal, rhs));
                    i = end;
                    continue;
                }
            }
            out.push_str(literal);
            continue;
        }
        out.push(c);
        i += c.len_utf8();
    }
    out
}

/// Capture an expression atom starting at position k: a balanced-paren group
/// (without outer parens) or `name.attr[0].call(...)` style atom.
fn capture_rhs(tag: &str, k: usize) -> (String, usize) {
    if k >= tag.len() {
        return (String::new(), k);
    }
    if tag[k..].starts_with('(') {
        // balanced parens; content without outer parens
        let mut depth = 0i32;
        let mut j = k;
        while j < tag.len() {
            let c = tag[j..].chars().next().unwrap();
            match c {
                '(' | '[' | '{' => depth += 1,
                ')' | ']' | '}' => {
                    depth -= 1;
                    if depth == 0 {
                        return (tag[k + 1..j].to_string(), j + 1);
                    }
                }
                _ => {}
            }
            j += c.len_utf8();
        }
        return (String::new(), k);
    }
    // atom: dotted name possibly with calls / subscripts
    let mut j = k;
    while j < tag.len() {
        let c = tag[j..].chars().next().unwrap();
        if c.is_alphanumeric() || c == '_' || c == '.' {
            j += c.len_utf8();
        } else if c == '(' || c == '[' {
            let close = if c == '(' { ')' } else { ']' };
            let mut depth = 0i32;
            while j < tag.len() {
                let d = tag[j..].chars().next().unwrap();
                if d == c {
                    depth += 1;
                } else if d == close {
                    depth -= 1;
                    if depth == 0 {
                        j += d.len_utf8();
                        break;
                    }
                }
                j += d.len_utf8();
            }
        } else {
            break;
        }
    }
    (tag[k..j].to_string(), j)
}

/// Python % formatting with positional args.
fn percent_format_positional(fmt: &str, args: &[Value]) -> Result<String, minijinja::Error> {
    let mut out = String::new();
    let mut rest = fmt;
    let mut arg_iter = args.iter();
    while let Some(pos) = rest.find('%') {
        out.push_str(&rest[..pos]);
        let tail = &rest[pos + 1..];
        if let Some(stripped) = tail.strip_prefix('%') {
            out.push('%');
            rest = stripped;
            continue;
        }
        // parse spec inline: [-+ #0]*[width][.precision]type
        let mut spec_part = tail;
        let mut spec = String::new();
        loop {
            match spec_part.chars().next() {
                Some('-') => spec.push('<'),
                Some('+') => spec.push('+'),
                Some(' ') => spec.push(' '),
                Some('0') => spec.push('0'),
                Some('#') => spec.push('#'),
                _ => break,
            }
            spec_part = &spec_part[1..];
        }
        while spec_part.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false) {
            spec.push(spec_part.chars().next().unwrap());
            spec_part = &spec_part[1..];
        }
        if spec_part.starts_with('.') {
            spec.push('.');
            spec_part = &spec_part[1..];
            while spec_part.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false) {
                spec.push(spec_part.chars().next().unwrap());
                spec_part = &spec_part[1..];
            }
        }
        let ty = spec_part.chars().next().unwrap_or('s');
        spec_part = &spec_part[1..];
        rest = spec_part;
        spec.push(ty);
        // %(name)s is not valid in positional mode; treat '(' as literal
        if ty == '(' {
            out.push('%');
            continue;
        }
        let Some(value) = arg_iter.next() else {
            break;
        };
        let text = value.to_string();
        out.push_str(&apply_format_spec(&text, value, &spec));
    }
    out.push_str(rest);
    Ok(out)
}

/// Python `== True` semantics: true for booleans True and numbers equal to 1.
fn test_eq_true(value: &Value) -> bool {
    if let Some(o) = value.as_object() {
        if let Some(b) = o.downcast_ref::<crate::pybridge::PyBoolObj>() {
            return b.0;
        }
        return false;
    }
    match value.kind() {
        minijinja::value::ValueKind::Bool => value.is_true(),
        minijinja::value::ValueKind::Number => value.as_i64() == Some(1)
            || f64::try_from(value.clone()).map(|f| f == 1.0).unwrap_or(false),
        _ => false,
    }
}

/// Python `== False` semantics.
fn test_eq_false(value: &Value) -> bool {
    if let Some(o) = value.as_object() {
        if let Some(b) = o.downcast_ref::<crate::pybridge::PyBoolObj>() {
            return !b.0;
        }
        return false;
    }
    match value.kind() {
        minijinja::value::ValueKind::Bool => !value.is_true(),
        minijinja::value::ValueKind::Number => value.as_i64() == Some(0)
            || f64::try_from(value.clone()).map(|f| f == 0.0).unwrap_or(false),
        _ => false,
    }
}

/// Heuristic repair for ill-formed xml (lxml recover-mode equivalent):
/// escape stray `&` and `<` that are clearly not markup, so structured
/// fixes can still run when values were inserted unescaped.
pub fn recover_xml(xml: &str) -> String {
    // & not starting a valid entity -> &amp;
    let s = crate::patch::sub(
        r"&(?!(?:amp|lt|gt|quot|apos);|#[0-9]+;|#x[0-9a-fA-F]+;)",
        |_| "&amp;".to_string(),
        xml,
    );
    // < is markup only if it opens/closes a namespace-qualified element
    // (docx parts always use prefixes), a declaration or a comment/doctype;
    // everything else (e.g. text "a<b") is escaped, which also yields
    // schema-valid output instead of lxml's unknown-element recovery
    crate::patch::sub(
        r"<(?!(?:[a-zA-Z_][\w.-]*:)|[/?!])",
        |_| "&lt;".to_string(),
        &s,
    )
}

pub fn render_xml_str(src_xml: &str, ctx: Value, autoescape: bool, core: &mut TplCore) -> Result<String, String> {
    let env = make_env(autoescape, core);
    render_xml_str_with(src_xml, ctx, &env, core)
}

/// render_xml_str with a pre-built environment (shared across all parts of
/// one render call).
pub fn render_xml_str_with(src_xml: &str, ctx: Value, env: &Environment, core: &mut TplCore) -> Result<String, String> {
    // add newlines before paragraphs so template error line numbers are useful
    // (plain memmem scan; the fancy-regex original paid per-match capture
    // expansion over the whole part)
    let src = add_paragraph_newlines(src_xml);
    let src = preprocess_trans(&src);
    let src = preprocess_bool_compare(&src);
    let src = preprocess_engine_features(&src);

    let rendered = match env.template_from_str(&src).and_then(|t| t.render(&ctx)) {
        Ok(s) => s,
        Err(e) => {
            let mut msg = format!("{}", e);
            if let Some(line) = e.line() {
                let start = line.saturating_sub(4);
                let context: Vec<String> = src
                    .lines()
                    .skip(start)
                    .take(7)
                    .map(|l| sub_str(r"<[^>]+>", "", l))
                    .collect();
                core.last_error_context = context.clone();
                msg.push_str(&format!(
                    "\nContext (lines {}-{}):\n{}",
                    start + 1,
                    start + 7,
                    context.join("\n")
                ));
            }
            return Err(msg);
        }
    };

    let dst = remove_paragraph_newlines(&rendered);
    Ok(restore_escaped_delims(dst.into_owned()))
}

/// `<w:p([ >])` -> `\n<w:p$1` without regex: insert `\n` before every
/// `<w:p ` / `<w:p>` opening tag.
fn add_paragraph_newlines(src: &str) -> Cow<'_, str> {
    if !src.contains("<w:p") {
        return Cow::Borrowed(src);
    }
    // every match inserts exactly one '\n'
    let extra = src.matches("<w:p ").count() + src.matches("<w:p>").count();
    let mut out = String::with_capacity(src.len() + extra);
    let mut rest = src;
    while let Some(p) = rest.find("<w:p") {
        let after = p + 4;
        match rest.as_bytes().get(after) {
            Some(b' ') | Some(b'>') => {
                out.push_str(&rest[..p]);
                out.push_str("\n<w:p");
            }
            _ => {
                out.push_str(&rest[..after]);
            }
        }
        rest = &rest[after..];
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// `\n<w:p([ >])` -> `<w:p$1` without regex: drop a `\n` immediately before
/// `<w:p ` / `<w:p>`; borrowed when there is nothing to remove.
fn remove_paragraph_newlines(s: &str) -> Cow<'_, str> {
    if !s.contains("\n<w:p") {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(p) = rest.find("\n<w:p") {
        let after = p + 5;
        match rest.as_bytes().get(after) {
            Some(b' ') | Some(b'>') => {
                out.push_str(&rest[..p]); // drop the '\n'
                out.push_str("<w:p");
            }
            _ => {
                out.push_str(&rest[..after]);
            }
        }
        rest = &rest[after..];
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// Linear-scan equivalent of the tag-finding regex `(?s)(\{%.*?%\}|\{\{.*?\}\})`:
/// calls `f` with each `{% ... %}` / `{{ ... }}` tag (closer included, first
/// matching closer wins, non-overlapping left to right); `f` returns
/// Some(replacement) to splice it in or None to keep the tag verbatim.
/// Borrowed when no tag was replaced.
fn rewrite_jinja_tags<'a>(
    src: &'a str,
    mut f: impl FnMut(&str) -> Option<String>,
) -> Cow<'a, str> {
    let b = src.as_bytes();
    let mut out: Option<String> = None;
    let mut copied = 0usize;
    let mut i = 0usize;
    while i + 3 < b.len() {
        if b[i] == b'{' && (b[i + 1] == b'%' || b[i + 1] == b'{') {
            let closer = if b[i + 1] == b'%' { "%}" } else { "}}" };
            if let Some(rel) = src[i + 2..].find(closer) {
                let tag_end = i + 2 + rel + 2;
                if let Some(rep) = f(&src[i..tag_end]) {
                    let o = out.get_or_insert_with(|| String::with_capacity(src.len()));
                    o.push_str(&src[copied..i]);
                    o.push_str(&rep);
                    copied = tag_end;
                }
                i = tag_end;
                continue;
            }
        }
        i += 1;
    }
    match out {
        Some(mut o) => {
            o.push_str(&src[copied..]);
            Cow::Owned(o)
        }
        None => Cow::Borrowed(src),
    }
}

/// Single-pass equivalent of the four chained replaces
/// `{_{`→`{{`, `}_}`→`}}`, `{%`←`{_%`, `%}`←`%_}`; returns the input
/// unchanged (no copy) when none of the escape sequences are present.
fn restore_escaped_delims(s: String) -> String {
    if !s.contains("{_") && !s.contains("}_") && !s.contains("%_") {
        return s;
    }
    let b = s.as_bytes();
    let mut out = String::with_capacity(s.len());
    let mut last = 0usize;
    let mut i = 0usize;
    while i + 2 < b.len() {
        // all marker bytes are ascii, so slicing at i / i+3 is char-safe
        let rep = if b[i + 1] == b'_' {
            match (b[i], b[i + 2]) {
                (b'{', b'{') => Some("{{"),
                (b'}', b'}') => Some("}}"),
                (b'{', b'%') => Some("{%"),
                (b'%', b'}') => Some("%}"),
                _ => None,
            }
        } else {
            None
        };
        match rep {
            Some(r) => {
                out.push_str(&s[last..i]);
                out.push_str(r);
                i += 3;
                last = i;
            }
            None => i += 1,
        }
    }
    out.push_str(&s[last..]);
    out
}

// ---------------- fix tables & docPr ids ----------------

pub fn fix_tables_and_docpr(xml: &str, docx_ids_index: &mut u32) -> Result<String, String> {
    fix_tables_docpr_cnvpr(xml, docx_ids_index, None)
}

/// fix_tables + docPr renumber + (optionally) pic:cNvPr renumber in a single
/// DOM parse/serialize round-trip. The cNvPr pass (docxcompose
/// renumber_nvpicpr_ids parity) is only wanted when subdocs were merged;
/// folding it in here saves a second full parse+serialize of the body.
fn fix_tables_docpr_cnvpr(
    xml: &str,
    docx_ids_index: &mut u32,
    cnvpr_next: Option<&mut u32>,
) -> Result<String, String> {
    // nothing to fix without any table, drawing or (when requested) picture:
    // skip the full DOM parse+serialize round-trip
    let need_cnvpr = cnvpr_next.is_some();
    if !xml.contains("<w:tbl")
        && !xml.contains("wp:docPr")
        && !(need_cnvpr && xml.contains("pic:cNvPr"))
    {
        return Ok(xml.to_string());
    }
    // If the rendered xml is not well-formed (e.g. unescaped values without
    // autoescape), attempt recovery first (docxtpl uses an lxml recover-mode
    // parser). As a last resort, apply regex-based fixes so table grids and
    // docPr ids are corrected even for severely broken xml.
    let doc = match Document::parse(xml) {
        Ok(d) => Some(d),
        Err(_) => match Document::parse(&recover_xml(xml)) {
            Ok(d) => Some(d),
            Err(_) => None,
        },
    };
    match doc {
        Some(mut doc) => {
            fix_tables_elem(&mut doc.root);
            fix_docpr_elem(&mut doc.root, docx_ids_index);
            if let Some(next) = cnvpr_next {
                fix_cnvpr_elem(&mut doc.root, next);
            }
            // round-trip output stays close to the input size (fixes are local)
            Ok(doc.serialize_with_capacity(xml.len() + xml.len() / 8 + 64))
        }
        None => {
            let s = regex_fix_tables(xml);
            let mut s = regex_fix_docpr(&s, docx_ids_index);
            // cNvPr renumbering is DOM-based; attempt it on the regex-fixed
            // xml, exactly like the old fix-then-renumber sequence
            if let Some(next) = cnvpr_next {
                if let Some(fixed) = renumber_cnvpr(&s, next) {
                    s = fixed;
                }
            }
            Ok(s)
        }
    }
}

/// Renumber pic:cNvPr ids sequentially (docxcompose renumber_nvpicpr_ids).
fn renumber_cnvpr(xml: &str, next_id: &mut u32) -> Option<String> {
    if !xml.contains("pic:cNvPr") {
        return None;
    }
    let mut doc = Document::parse(xml).ok()?;
    fix_cnvpr_elem(&mut doc.root, next_id);
    Some(doc.serialize_with_capacity(xml.len() + 64))
}

fn fix_cnvpr_elem(el: &mut Element, next_id: &mut u32) {
    if el.name == "pic:cNvPr" {
        el.set_attr("id", &next_id.to_string());
        *next_id += 1;
    }
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            fix_cnvpr_elem(e, next_id);
        }
    }
}

/// Last-resort docPr renumbering for unparseable xml.
fn regex_fix_docpr(xml: &str, idx: &mut u32) -> String {
    crate::patch::sub(
        r#"<wp:docPr id="\d+""#,
        |m| {
            *idx += 1;
            let _ = m;
            format!("<wp:docPr id=\"{}\"", idx)
        },
        xml,
    )
}

/// Last-resort table grid fixing for unparseable xml: adjusts gridCol count
/// and widths to match the maximum cell count per row.
fn regex_fix_tables(xml: &str) -> String {
    // find top-level table regions by nesting depth
    let mut regions: Vec<(usize, usize)> = Vec::new();
    {
        let re_tbl = crate::patch::re(r"<w:tbl[ >]|</w:tbl>");
        let mut depth = 0i32;
        let mut start = 0usize;
        for m in re_tbl.find_iter(xml).flatten() {
            if m.as_str().starts_with("</") {
                depth -= 1;
                if depth == 0 {
                    regions.push((start, m.end()));
                }
            } else {
                if depth == 0 {
                    start = m.start();
                }
                depth += 1;
            }
        }
    }
    if regions.is_empty() {
        return xml.to_string();
    }

    let gridcol_re = crate::patch::re(r#"<w:gridCol w:w="(\d+)"/>"#);
    let mut out = String::with_capacity(xml.len());
    let mut last = 0usize;
    for (rs, re) in regions {
        out.push_str(&xml[last..rs]);
        let region = &xml[rs..re];
        let fixed = (|| {
            let grid_start = region.find("<w:tblGrid>")?;
            let grid_end = region.find("</w:tblGrid>")?;
            let grid = &region[grid_start..grid_end];
            let mut widths: Vec<i64> = gridcol_re
                .captures_iter(grid)
                .flatten()
                .map(|c| c[1].parse().unwrap_or(0))
                .collect();
            let n_cols = widths.len();
            // max cells in any row of the region
            let mut max_cells = 0usize;
            for row in region.split("<w:tr").skip(1) {
                let row_end = row.find("</w:tr>").unwrap_or(row.len());
                let cells = row[..row_end].matches("<w:tc>").count()
                    + row[..row_end].matches("<w:tc ").count();
                max_cells = max_cells.max(cells);
            }
            if max_cells == 0 || n_cols == 0 {
                return None;
            }
            let total: i64 = widths.iter().sum();
            if max_cells > n_cols && total > 0 {
                let old_avg = total as f64 / n_cols as f64;
                let new_avg = (total as f64 / max_cells as f64) as i64;
                for w in widths.iter_mut() {
                    *w = (*w as f64 * new_avg as f64 / old_avg) as i64;
                }
                while widths.len() < max_cells {
                    widths.push(new_avg);
                }
            } else if n_cols > max_cells {
                let removed: i64 = widths.drain(max_cells..).sum();
                let extra = removed / max_cells.max(1) as i64;
                for w in widths.iter_mut() {
                    *w += extra;
                }
            } else {
                return None;
            }
            // rebuild the region with the new grid
            let new_grid: String = widths
                .iter()
                .map(|w| format!("<w:gridCol w:w=\"{}\"/>", w))
                .collect();
            Some(format!(
                "{}<w:tblGrid>{}</w:tblGrid>{}",
                &region[..grid_start],
                new_grid,
                &region[grid_end + "</w:tblGrid>".len()..]
            ))
        })();
        out.push_str(fixed.as_deref().unwrap_or(region));
        last = re;
    }
    out.push_str(&xml[last..]);
    out
}

fn fix_docpr_elem(el: &mut Element, idx: &mut u32) {
    if el.name == "wp:docPr" {
        *idx += 1;
        el.set_attr("id", &idx.to_string());
    }
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            fix_docpr_elem(e, idx);
        }
    }
}

fn fix_tables_elem(el: &mut Element) {
    if el.name == "w:tbl" {
        fix_one_table(el);
    }
    for c in el.children.iter_mut() {
        if let Node::Elem(e) = c {
            fix_tables_elem(e);
        }
    }
}

fn fix_one_table(tbl: &mut Element) {
    let Some(tbl_grid_idx) = tbl
        .children
        .iter()
        .position(|c| matches!(c, Node::Elem(e) if e.name == "w:tblGrid"))
    else {
        return;
    };

    // snapshot column widths
    let grid = match &tbl.children[tbl_grid_idx] {
        Node::Elem(e) => e,
        _ => return,
    };
    let mut col_widths: Vec<f64> = grid
        .find_all("w:gridCol")
        .iter()
        .map(|c| {
            c.get_attr("w:w")
                .and_then(|w| w.parse::<f64>().ok())
                .unwrap_or(0.0)
        })
        .collect();
    let n_columns = col_widths.len();

    // count stats across all descendant rows (immutable pass)
    let mut rows: Vec<&Element> = Vec::new();
    tbl.iter_descendants("w:tr", &mut rows);
    let mut to_add = 0usize;
    for r in &rows {
        let cells = r.find_all("w:tc");
        if n_columns + to_add < cells.len() {
            to_add = cells.len() - n_columns;
        }
    }

    let get_cell_len = |cell: &Element| -> usize {
        if let Some(tcpr) = cell.find("w:tcPr") {
            if let Some(gs) = tcpr.find("w:gridSpan") {
                if let Some(v) = gs.get_attr("w:val").and_then(|v| v.parse::<usize>().ok()) {
                    return v;
                }
            }
        }
        1
    };
    let mut cells_len_max = 0usize;
    for r in &rows {
        let cells = r.find_all("w:tc");
        let len: usize = cells.iter().map(|c| get_cell_len(c)).sum();
        cells_len_max = cells_len_max.max(len);
    }
    drop(rows);

    if to_add > 0 {
        let width: f64 = col_widths.iter().sum();
        if width > 0.0 && n_columns > 0 {
            let old_average = width / n_columns as f64;
            let new_average = width / (n_columns + to_add) as f64;
            for w in col_widths.iter_mut() {
                *w = (*w * new_average / old_average).trunc();
            }
            for _ in 0..to_add {
                col_widths.push(new_average.trunc());
            }
            // write back
            if let Node::Elem(grid) = &mut tbl.children[tbl_grid_idx] {
                let mut i = 0usize;
                for c in grid.children.iter_mut() {
                    if let Node::Elem(e) = c {
                        if e.name == "w:gridCol" && i < col_widths.len() {
                            e.set_attr("w:w", &(col_widths[i] as i64).to_string());
                            i += 1;
                        }
                    }
                }
                for w in col_widths.iter().skip(i) {
                    let mut gc = Element::new("w:gridCol");
                    gc.set_attr("w:w", &(*w as i64).to_string());
                    grid.children.push(Node::Elem(gc));
                }
            }
        }
    }

    // refetch columns
    let grid = match &tbl.children[tbl_grid_idx] {
        Node::Elem(e) => e,
        _ => return,
    };
    let columns = grid.find_all("w:gridCol");
    let columns_len = columns.len();
    let col_widths2: Vec<f64> = columns
        .iter()
        .map(|c| {
            c.get_attr("w:w")
                .and_then(|w| w.parse::<f64>().ok())
                .unwrap_or(0.0)
        })
        .collect();

    let to_remove = columns_len.saturating_sub(cells_len_max);
    if to_remove > 0 && columns_len > 0 {
        let removed_width: f64 = col_widths2[columns_len - to_remove..].iter().sum();
        if let Node::Elem(grid) = &mut tbl.children[tbl_grid_idx] {
            // remove last to_remove gridCol children
            let mut idxs: Vec<usize> = grid
                .children
                .iter()
                .enumerate()
                .filter_map(|(i, c)| match c {
                    Node::Elem(e) if e.name == "w:gridCol" => Some(i),
                    _ => None,
                })
                .collect();
            idxs.reverse();
            for i in idxs.into_iter().take(to_remove) {
                grid.children.remove(i);
            }
            // redistribute removed width
            let n_left = grid.find_all("w:gridCol").len();
            let extra = if n_left > 0 {
                (removed_width / n_left as f64) as i64
            } else {
                0
            };
            for c in grid.children.iter_mut() {
                if let Node::Elem(e) = c {
                    if e.name == "w:gridCol" {
                        let w = e
                            .get_attr("w:w")
                            .and_then(|w| w.parse::<f64>().ok())
                            .unwrap_or(0.0);
                        e.set_attr("w:w", &((w as i64) + extra).to_string());
                    }
                }
            }
        }
    }
}
