//! docxcompose-style document concatenation: append whole docx documents
//! to the end of a master document's body.
//!
//! Reuses the subdoc merge pipeline ([`crate::subdoc::subdoc_xml_opts`]): style
//! conflicts are renamed (`X_1`), numbering ids are offset and the first list
//! is restarted, media is sha1-deduped, relationship ids are remapped,
//! footnotes and comments are merged and bookmark ids are shifted. A
//! page-break paragraph is inserted before each appended document — unless
//! the appended document has section properties worth keeping, in which case
//! they are preserved as its own section (page setup and header/footer
//! references included) and the section break starts the new page.

use crate::template::{fix_tables_docpr_cnvpr, renumber_cnvpr, TplCore, DOCUMENT_PART};
use crate::xmldom::Node;

/// Page-break paragraph inserted between documents (docxcompose parity).
const PAGE_BREAK_P: &str = "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>";

/// Compose multiple docx documents into one by appending their body content
/// to a master document.
pub struct Composer {
    core: TplCore,
    /// whether at least one document was appended since the last fix pass
    appended: bool,
}

impl Composer {
    /// Create a composer from the master document's docx bytes.
    pub fn new(master_bytes: Vec<u8>) -> Result<Composer, String> {
        let mut core = TplCore::new(master_bytes);
        core.init_docx(false)?;
        Ok(Composer {
            core,
            appended: false,
        })
    }

    /// Append one whole docx document to the end of the master's body
    /// (before its trailing `w:sectPr`), preceded by a page break. The
    /// appended document's section properties (page setup, header/footer
    /// references) are preserved as its own section; when they are, the
    /// section break already starts a new page and no extra page-break
    /// paragraph is inserted.
    pub fn append(&mut self, doc_bytes: &[u8]) -> Result<(), String> {
        // subdoc_xml_info merges styles/numbering/media/rels/comments into
        // the package; make sure the package holds the latest document.xml
        // first
        self.core.flush_doc()?;
        let (fragment, kept_sections) =
            crate::subdoc::subdoc_xml_info(&mut self.core, doc_bytes, true)?;
        let frag_dom = crate::subdoc::parse_body_fragment(&fragment)?;

        let dom = self.core.document_dom()?;
        let body = dom
            .root
            .find_mut("w:body")
            .ok_or_else(|| "master document has no w:body".to_string())?;
        // insert before the body-level sectPr (kept last), else at the end
        let pos = body
            .children
            .iter()
            .position(|c| matches!(c, Node::Elem(e) if e.name == "w:sectPr"))
            .unwrap_or(body.children.len());
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

    /// Build a minimal docx with the given body paragraphs.
    fn make_docx(paras: &str) -> Vec<u8> {
        use std::io::Write as _;
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
            paras
        );
        let mut cursor = std::io::Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut cursor);
            let opts = zip::write::SimpleFileOptions::default();
            for (name, data) in [
                ("[Content_Types].xml", &ct[..]),
                ("_rels/.rels", &rels[..]),
                (DOCUMENT_PART, doc.as_bytes()),
            ] {
                w.start_file(name, opts).unwrap();
                w.write_all(data).unwrap();
            }
            w.finish().unwrap();
        }
        cursor.into_inner()
    }

    fn body_xml(docx: &[u8]) -> String {
        let pkg = crate::package::Package::from_bytes(docx).unwrap();
        pkg.get_string(DOCUMENT_PART).unwrap()
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
}
