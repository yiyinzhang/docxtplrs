//! docxcompose-style document concatenation: insert whole docx documents
//! into a master document's body.
//!
//! Reuses the subdoc merge pipeline ([`crate::subdoc::subdoc_xml_opts`]):
//! style conflicts are renamed (`X_1`) — or, with `preserve_styles = false`,
//! mapped to the master's same-named styles — numbering ids are offset and
//! the first list is restarted (skippable via `restart_numbering`), media is
//! sha1-deduped, relationship ids are remapped, footnotes and comments are
//! merged and bookmark ids are shifted. A page-break paragraph is inserted
//! before each appended document — unless the appended document has section
//! properties worth keeping, in which case they are preserved as its own
//! section (page setup and header/footer references included) and the
//! section break starts the new page.

use crate::template::{fix_tables_docpr_cnvpr, renumber_cnvpr, TplCore, DOCUMENT_PART};
use crate::xmldom::Node;

/// Page-break paragraph inserted between documents (docxcompose parity).
const PAGE_BREAK_P: &str = "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>";

/// Compose multiple docx documents into one by inserting their body content
/// into a master document (docxcompose Composer parity).
pub struct Composer {
    core: TplCore,
    /// whether at least one document was appended since the last fix pass
    appended: bool,
    /// restart numbering of each merged document's first list
    /// (docxcompose restart_numbering, default true)
    pub restart_numbering: bool,
    /// docxcompose preserve_styles: false (default) maps the merged
    /// document's styles to the master's same-named styles; true keeps them
    /// (conflicts renamed `X_1`) and retains the default paragraph style's
    /// formatting on unstyled runs/paragraphs
    pub preserve_styles: bool,
}

impl Composer {
    /// Create a composer from the master document's docx bytes.
    pub fn new(master_bytes: Vec<u8>) -> Result<Composer, String> {
        let mut core = TplCore::new(master_bytes);
        core.init_docx(false)?;
        Ok(Composer {
            core,
            appended: false,
            restart_numbering: true,
            preserve_styles: false,
        })
    }

    /// Body index at which [`Composer::append`] inserts: the master body's
    /// first `w:sectPr`, or the end of the body (docxcompose append_index).
    pub fn append_index(&mut self) -> Result<usize, String> {
        self.core.flush_doc()?;
        let dom = self.core.document_dom()?;
        let body = dom
            .root
            .find("w:body")
            .ok_or_else(|| "master document has no w:body".to_string())?;
        Ok(body
            .children
            .iter()
            .position(|c| matches!(c, Node::Elem(e) if e.name == "w:sectPr"))
            .unwrap_or(body.children.len()))
    }

    /// Append one whole docx document to the end of the master's body
    /// (before its trailing `w:sectPr`), preceded by a page break.
    pub fn append(&mut self, doc_bytes: &[u8]) -> Result<(), String> {
        self.append_opts(doc_bytes, true)
    }

    /// docxcompose append(): insert at [`Composer::append_index`].
    pub fn append_opts(&mut self, doc_bytes: &[u8], remove_property_fields: bool) -> Result<(), String> {
        let index = self.append_index()?;
        self.insert_at(index, doc_bytes, remove_property_fields)
    }

    /// Insert one whole docx document's body content at the given body
    /// index (docxcompose insert()), preceded by a page break. The inserted
    /// document's section properties (page setup, header/footer references)
    /// are preserved as its own section; when they are, the section break
    /// already starts a new page and no extra page-break paragraph is
    /// inserted. With `remove_property_fields`, DOCPROPERTY fields
    /// referencing custom document properties are dissolved into their
    /// literal values before merging.
    pub fn insert_at(
        &mut self,
        index: usize,
        doc_bytes: &[u8],
        remove_property_fields: bool,
    ) -> Result<(), String> {
        let owned;
        let doc_bytes = if remove_property_fields {
            owned = crate::customprops::dissolve_custom_property_fields(doc_bytes)?;
            &owned[..]
        } else {
            doc_bytes
        };
        // subdoc_xml_info_opts merges styles/numbering/media/rels/comments
        // into the package; make sure the package holds the latest
        // document.xml first
        self.core.flush_doc()?;
        let opts = crate::subdoc::SubdocOpts {
            keep_sections: true,
            restart_numbering: self.restart_numbering,
            // the composer dissolves property fields itself (value-aware);
            // with remove_property_fields=false they must be preserved
            dissolve_property_fields: false,
            map_styles_by_name: !self.preserve_styles,
            retain_default_formatting: self.preserve_styles,
        };
        let (fragment, kept_sections) =
            crate::subdoc::subdoc_xml_info_opts(&mut self.core, doc_bytes, opts)?;
        let frag_dom = crate::subdoc::parse_body_fragment(&fragment)?;

        let dom = self.core.document_dom()?;
        let body = dom
            .root
            .find_mut("w:body")
            .ok_or_else(|| "master document has no w:body".to_string())?;
        let pos = index.min(body.children.len());
        let mut inserted: Vec<Node> = Vec::new();
        if !kept_sections {
            let break_dom = crate::subdoc::parse_body_fragment(PAGE_BREAK_P)?;
            inserted.extend(break_dom.root.children);
        }
        inserted.extend(frag_dom.root.children);
        let tail = body.children.split_off(pos);
        body.children.extend(inserted);
        body.children.extend(tail);
        self.core.mark_doc_dirty();
        self.appended = true;
        Ok(())
    }

    /// Serialize the composed document. When documents were appended, runs
    /// the same post-merge fixups as the render pipeline: fix_tables + docPr
    /// renumber + pic:cNvPr renumber on the body, cNvPr renumber on
    /// headers/footers (docxcompose renumber_nvpicpr_ids parity).
    pub fn save_bytes(&mut self) -> Result<Vec<u8>, String> {
        self.core.flush_parts()?;
        if self.appended {
            let mut next_id: u32 = 1;
            {
                let pkg = self
                    .core
                    .package
                    .as_mut()
                    .ok_or_else(|| "package not loaded".to_string())?;
                let xml = pkg
                    .get_string(DOCUMENT_PART)
                    .ok_or_else(|| format!("missing {}", DOCUMENT_PART))?;
                let enc = pkg.encoding_of(DOCUMENT_PART);
                let fixed =
                    fix_tables_docpr_cnvpr(&xml, &mut self.core.docx_ids_index, Some(&mut next_id))?;
                pkg.set(
                    DOCUMENT_PART,
                    crate::package::encode_part_owned(fixed, &enc),
                );
            }
            self.core.invalidate_doc();
            let hf_parts: Vec<String> = [crate::package::rel_type::HEADER, crate::package::rel_type::FOOTER]
                .into_iter()
                .flat_map(|uri| self.core.header_footer_parts(uri))
                .collect();
            {
                let pkg = self.core.package.as_mut().unwrap();
                for part in hf_parts {
                    let Some(xml) = pkg.get_string(&part) else {
                        continue;
                    };
                    let enc = pkg.encoding_of(&part);
                    if let Some(fixed) = renumber_cnvpr(&xml, &mut next_id) {
                        pkg.set(&part, crate::package::encode_part_owned(fixed, &enc));
                    }
                }
            }
            self.core.invalidate_parts();
            self.appended = false;
        }
        self.core.save_bytes()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn para(text: &str) -> String {
        format!(
            "<w:p><w:r><w:t>{}</w:t></w:r></w:p>",
            crate::package::escape_xml_text(text)
        )
    }

    /// Build a minimal docx with the given body paragraphs and extra parts.
    fn make_docx_with(paras: &str, extra: &[(&str, String)]) -> Vec<u8> {
        use std::io::Write as _;
        let ct = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>"#;
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
        let doc = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}<w:sectPr/></w:body></w:document>",
            paras
        );
        let mut cursor = std::io::Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut cursor);
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("[Content_Types].xml", opts).unwrap();
            w.write_all(&ct[..]).unwrap();
            w.start_file("_rels/.rels", opts).unwrap();
            w.write_all(&rels[..]).unwrap();
            w.start_file(DOCUMENT_PART, opts).unwrap();
            w.write_all(doc.as_bytes()).unwrap();
            for (name, data) in extra {
                w.start_file(name, opts).unwrap();
                w.write_all(data.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        cursor.into_inner()
    }

    /// Build a minimal docx with the given body paragraphs.
    fn make_docx(paras: &str) -> Vec<u8> {
        make_docx_with(paras, &[])
    }

    fn body_xml(docx: &[u8]) -> String {
        part_xml(docx, DOCUMENT_PART)
    }

    fn part_xml(docx: &[u8], part: &str) -> String {
        let pkg = crate::package::Package::from_bytes(docx).unwrap();
        pkg.get_string(part).unwrap()
    }

    const W: &str = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    /// docProps/custom.xml with vt:lpwstr text properties.
    fn custom_props_xml(pairs: &[(&str, &str)]) -> String {
        let mut props = String::new();
        for (i, (name, value)) in pairs.iter().enumerate() {
            props.push_str(&format!(
                "<property fmtid=\"{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}\" pid=\"{}\" name=\"{}\"><vt:lpwstr>{}</vt:lpwstr></property>",
                i + 2,
                name,
                value
            ));
        }
        format!(
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">{}</Properties>",
            props
        )
    }

    #[test]
    fn test_composer_appends_with_page_breaks() {
        let master = make_docx(&para("master"));
        let sub1 = make_docx(&para("one"));
        let sub2 = make_docx(&para("two"));

        let mut c = Composer::new(master).unwrap();
        c.append(&sub1).unwrap();
        c.append(&sub2).unwrap();
        let out = c.save_bytes().unwrap();

        let xml = body_xml(&out);
        // note: search with text-tag delimiters ("one" is a substring of
        // "standalone" in the xml declaration)
        let i_master = xml.find(">master<").unwrap();
        let i_br1 = xml.find("<w:br").unwrap();
        let i_one = xml.find(">one<").unwrap();
        let i_br2 = xml.rfind("<w:br").unwrap();
        let i_two = xml.find(">two<").unwrap();
        let i_sectpr = xml.find("<w:sectPr").unwrap();
        assert!(i_master < i_br1 && i_br1 < i_one && i_one < i_br2 && i_br2 < i_two);
        assert!(i_two < i_sectpr, "sectPr must stay last");
    }

    #[test]
    fn test_insert_at_middle_of_body() {
        let master = make_docx(&format!("{}{}{}", para("a"), para("b"), para("c")));
        let sub = make_docx(&para("X"));

        let mut c = Composer::new(master).unwrap();
        assert_eq!(c.append_index().unwrap(), 3); // before the trailing sectPr
        c.insert_at(1, &sub, true).unwrap();
        let xml = body_xml(&c.save_bytes().unwrap());

        let ia = xml.find(">a<").unwrap();
        let ibr = xml.find("<w:br").unwrap();
        let ix = xml.find(">X<").unwrap();
        let ib = xml.find(">b<").unwrap();
        let ic = xml.find(">c<").unwrap();
        assert!(ia < ibr && ibr < ix && ix < ib && ib < ic);
    }

    /// remove_property_fields dissolves fldSimple fields into the custom
    /// property's literal value; fields naming unknown properties are kept.
    #[test]
    fn test_remove_property_fields_simple() {
        let body = concat!(
            "<w:p><w:fldSimple w:instr=\" DOCPROPERTY &quot;Firma&quot; \\* MERGEFORMAT \">",
            "<w:r><w:t>stale</w:t></w:r></w:fldSimple></w:p>",
            "<w:p><w:fldSimple w:instr=\" DOCPROPERTY &quot;Unknown&quot; \\* MERGEFORMAT \">",
            "<w:r><w:t>keep</w:t></w:r></w:fldSimple></w:p>"
        );
        let sub = make_docx_with(
            body,
            &[(
                "docProps/custom.xml",
                custom_props_xml(&[("Firma", "ACME")]),
            )],
        );
        let out = crate::customprops::dissolve_custom_property_fields(&sub).unwrap();
        let xml = body_xml(&out);
        assert!(xml.contains(">ACME<"), "value dissolved: {}", xml);
        assert!(!xml.contains("stale"));
        // unknown property: field untouched
        assert!(xml.contains("Unknown"));
        assert!(xml.contains(">keep<"));
        assert_eq!(xml.matches("fldSimple").count(), 2); // open + close tag
    }

    /// remove_property_fields dissolves complex fields (begin/instrText/
    /// separate/end), keeping the value run's formatting.
    #[test]
    fn test_remove_property_fields_complex() {
        let body = concat!(
            "<w:p>",
            "<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType=\"begin\"/></w:r>",
            "<w:r><w:instrText xml:space=\"preserve\"> DOCPROPERTY \"Ort\" \\* MERGEFORMAT </w:instrText></w:r>",
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>",
            "<w:r><w:rPr><w:i/></w:rPr><w:t>stale</w:t></w:r>",
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
            "</w:p>"
        );
        let sub = make_docx_with(body, &[("docProps/custom.xml", custom_props_xml(&[("Ort", "Bern")]))]);
        let out = crate::customprops::dissolve_custom_property_fields(&sub).unwrap();
        let xml = body_xml(&out);
        assert!(xml.contains(">Bern<"), "value dissolved: {}", xml);
        assert!(!xml.contains("stale"));
        assert!(!xml.contains("DOCPROPERTY"));
        assert!(!xml.contains("fldChar"));
        // the value run kept its formatting
        assert!(xml.contains("<w:i/>"));
    }

    /// With remove_property_fields=false the fields survive the merge.
    #[test]
    fn test_append_preserves_property_fields_when_disabled() {
        let body = "<w:p><w:fldSimple w:instr=\" DOCPROPERTY &quot;Firma&quot; \\* MERGEFORMAT \"><w:r><w:t>stale</w:t></w:r></w:fldSimple></w:p>";
        let sub = make_docx_with(body, &[("docProps/custom.xml", custom_props_xml(&[("Firma", "ACME")]))]);

        let mut c = Composer::new(make_docx(&para("m"))).unwrap();
        c.insert_at(1, &sub, false).unwrap();
        let xml = body_xml(&c.save_bytes().unwrap());
        assert!(xml.contains("DOCPROPERTY"));
        assert!(xml.contains("stale"));

        let mut c = Composer::new(make_docx(&para("m"))).unwrap();
        c.insert_at(1, &sub, true).unwrap();
        let xml = body_xml(&c.save_bytes().unwrap());
        assert!(xml.contains(">ACME<"));
        assert!(!xml.contains("DOCPROPERTY"));
    }

    fn numbered_docx(parts_styles: &str, body: &str) -> Vec<u8> {
        let numbering = format!(
            "<w:numbering {W}><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>"
        );
        make_docx_with(
            body,
            &[
                ("word/styles.xml", parts_styles.to_string()),
                ("word/numbering.xml", numbering),
            ],
        )
    }

    fn numbered_body(texts: &[&str]) -> String {
        texts
            .iter()
            .map(|t| {
                format!(
                    "<w:p><w:pPr><w:pStyle w:val=\"NumList\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>",
                    t
                )
            })
            .collect()
    }

    /// restart_numbering (default) gives the appended list a new w:num with
    /// startOverride=1; restart_numbering=false skips that.
    #[test]
    fn test_restart_numbering_toggle() {
        let styles = format!(
            "<w:styles {W}><w:style w:type=\"paragraph\" w:styleId=\"NumList\"><w:name w:val=\"NumList\"/></w:style></w:styles>"
        );
        let master = numbered_docx(&styles, &numbered_body(&["m1"]));
        let sub = numbered_docx(&styles, &numbered_body(&["s1", "s2"]));

        let mut c = Composer::new(master.clone()).unwrap();
        c.append(&sub).unwrap();
        let out = c.save_bytes().unwrap();
        let numbering = part_xml(&out, "word/numbering.xml");
        assert!(
            numbering.contains("startOverride"),
            "first list restarts: {}",
            numbering
        );

        let mut c = Composer::new(master).unwrap();
        c.restart_numbering = false;
        c.append(&sub).unwrap();
        let out = c.save_bytes().unwrap();
        let numbering = part_xml(&out, "word/numbering.xml");
        assert!(!numbering.contains("startOverride"));
        // the subdoc's list was still merged with an offset numId
        let xml = body_xml(&out);
        assert!(xml.contains("<w:numId w:val=\"3\"/>"), "merged numId: {}", xml);
    }

    /// preserve_styles=False (default): a style whose name exists in the
    /// master under a different id is mapped to the master's style id and
    /// not copied.
    #[test]
    fn test_preserve_styles_false_maps_by_name() {
        let master_styles = format!(
            "<w:styles {W}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style><w:style w:type=\"paragraph\" w:styleId=\"A\"><w:name w:val=\"Foo\"/><w:rPr><w:b/></w:rPr></w:style></w:styles>"
        );
        let sub_styles = format!(
            "<w:styles {W}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style><w:style w:type=\"paragraph\" w:styleId=\"B\"><w:name w:val=\"Foo\"/><w:rPr><w:i/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"C\"><w:name w:val=\"Bar\"/></w:style></w:styles>"
        );
        let master = make_docx_with(
            &para("m"),
            &[("word/styles.xml", master_styles)],
        );
        let sub = make_docx_with(
            "<w:p><w:pPr><w:pStyle w:val=\"B\"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>\
             <w:p><w:pPr><w:pStyle w:val=\"C\"/></w:pPr><w:r><w:t>other</w:t></w:r></w:p>",
            &[("word/styles.xml", sub_styles)],
        );

        let mut c = Composer::new(master.clone()).unwrap();
        c.append(&sub).unwrap();
        let out = c.save_bytes().unwrap();
        let xml = body_xml(&out);
        assert!(
            xml.contains("<w:pStyle w:val=\"A\"/>"),
            "reference mapped to master style id: {}",
            xml
        );
        let styles = part_xml(&out, "word/styles.xml");
        assert!(!styles.contains("w:styleId=\"B\""), "sub style not copied");
        assert!(styles.contains("w:styleId=\"C\""), "unknown style copied");

        // preserve_styles=True keeps the subdoc's style (no id conflict: B
        // is copied as-is) instead of mapping to the master's A
        let mut c = Composer::new(master).unwrap();
        c.preserve_styles = true;
        c.append(&sub).unwrap();
        let out = c.save_bytes().unwrap();
        let xml = body_xml(&out);
        assert!(xml.contains("<w:pStyle w:val=\"B\"/>"), "reference kept: {}", xml);
        let styles = part_xml(&out, "word/styles.xml");
        assert!(styles.contains("w:styleId=\"B\""), "sub style preserved");
    }

    /// preserve_styles=True retains the subdoc's default paragraph style
    /// formatting on unstyled runs and paragraphs.
    #[test]
    fn test_preserve_styles_true_retains_default_formatting() {
        let master_styles = format!(
            "<w:styles {W}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>"
        );
        let sub_styles = format!(
            "<w:styles {W}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:rPr><w:b/></w:rPr><w:pPr><w:spacing w:after=\"200\"/></w:pPr></w:style></w:styles>"
        );
        let master = make_docx_with(&para("m"), &[("word/styles.xml", master_styles)]);
        let sub = make_docx_with(
            "<w:p><w:r><w:t>plain</w:t></w:r></w:p>",
            &[("word/styles.xml", sub_styles)],
        );

        let mut c = Composer::new(master).unwrap();
        c.preserve_styles = true;
        c.append(&sub).unwrap();
        let out = c.save_bytes().unwrap();
        let xml = body_xml(&out);
        assert!(
            xml.contains("<w:p><w:pPr><w:spacing w:after=\"200\"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>plain</w:t></w:r></w:p>"),
            "default formatting retained: {}",
            xml
        );
    }
}
