//! Minimal XML DOM used for structured transformations (fix_tables, docPr ids,
//! rels/content-types editing). Namespace prefixes are treated literally, which
//! is sufficient for OOXML manipulation.

use quick_xml::events::{BytesDecl, Event};
use quick_xml::Reader;

#[derive(Debug, Clone)]
pub struct Element {
    pub name: String,
    pub attrs: Vec<(String, String)>,
    pub children: Vec<Node>,
}

#[derive(Debug, Clone)]
pub enum Node {
    Elem(Element),
    Text(String),
}

#[derive(Debug, Clone)]
pub struct Document {
    /// Everything before the root element (xml declaration etc.)
    pub prolog: String,
    pub root: Element,
}

fn escape_text(s: &str, out: &mut String) {
    // bulk-copy runs without special chars; all escaped chars are ASCII, so
    // byte indexing is char-safe
    let b = s.as_bytes();
    let mut last = 0usize;
    let mut i = 0usize;
    while i < b.len() {
        let rep = match b[i] {
            b'&' => "&amp;",
            b'<' => "&lt;",
            b'>' => "&gt;",
            _ => {
                i += 1;
                continue;
            }
        };
        out.push_str(&s[last..i]);
        out.push_str(rep);
        i += 1;
        last = i;
    }
    out.push_str(&s[last..]);
}

fn escape_attr(s: &str, out: &mut String) {
    let b = s.as_bytes();
    let mut last = 0usize;
    let mut i = 0usize;
    while i < b.len() {
        let rep = match b[i] {
            b'&' => "&amp;",
            b'<' => "&lt;",
            b'"' => "&quot;",
            _ => {
                i += 1;
                continue;
            }
        };
        out.push_str(&s[last..i]);
        out.push_str(rep);
        i += 1;
        last = i;
    }
    out.push_str(&s[last..]);
}

impl Element {
    pub fn new(name: &str) -> Self {
        Element {
            name: name.to_string(),
            attrs: Vec::new(),
            children: Vec::new(),
        }
    }

    pub fn get_attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == name)
            .map(|(_, v)| v.as_str())
    }

    pub fn set_attr(&mut self, name: &str, value: &str) {
        if let Some(slot) = self.attrs.iter_mut().find(|(k, _)| k == name) {
            slot.1 = value.to_string();
        } else {
            self.attrs.push((name.to_string(), value.to_string()));
        }
    }

    /// First direct child element with given name
    pub fn find(&self, name: &str) -> Option<&Element> {
        self.children.iter().find_map(|c| match c {
            Node::Elem(e) if e.name == name => Some(e),
            _ => None,
        })
    }

    pub fn find_mut(&mut self, name: &str) -> Option<&mut Element> {
        self.children.iter_mut().find_map(|c| match c {
            Node::Elem(e) if e.name == name => Some(e),
            _ => None,
        })
    }

    /// All direct child elements with given name
    pub fn find_all(&self, name: &str) -> Vec<&Element> {
        self.children
            .iter()
            .filter_map(|c| match c {
                Node::Elem(e) if e.name == name => Some(e),
                _ => None,
            })
            .collect()
    }

    /// Deep iteration over all descendant elements (including self) with given name.
    pub fn iter_descendants<'a>(&'a self, name: &'a str, out: &mut Vec<&'a Element>) {
        for c in &self.children {
            if let Node::Elem(e) = c {
                if e.name == name {
                    out.push(e);
                }
                e.iter_descendants(name, out);
            }
        }
    }

    pub fn serialize(&self, out: &mut String) {
        out.push('<');
        out.push_str(&self.name);
        for (k, v) in &self.attrs {
            out.push(' ');
            out.push_str(k);
            out.push_str("=\"");
            escape_attr(v, out);
            out.push('"');
        }
        if self.children.is_empty() {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for c in &self.children {
            match c {
                Node::Elem(e) => e.serialize(out),
                Node::Text(t) => escape_text(t, out),
            }
        }
        out.push_str("</");
        out.push_str(&self.name);
        out.push('>');
    }

    /// Concatenated text content of all descendants
    pub fn text_content(&self) -> String {
        let mut s = String::new();
        for c in &self.children {
            match c {
                Node::Text(t) => s.push_str(t),
                Node::Elem(e) => s.push_str(&e.text_content()),
            }
        }
        s
    }
}

impl Document {
    pub fn parse(xml: &str) -> Result<Document, String> {
        let mut reader = Reader::from_str(xml);
        let mut prolog = String::new();
        let mut stack: Vec<Element> = Vec::new();
        let mut root: Option<Element> = None;

        loop {
            match reader.read_event().map_err(|e| e.to_string())? {
                Event::Start(e) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let mut attrs = Vec::new();
                    for a in e.attributes() {
                        let a = a.map_err(|e| e.to_string())?;
                        let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                        let val = a
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(|e| e.to_string())?
                            .into_owned();
                        attrs.push((key, val));
                    }
                    stack.push(Element {
                        name,
                        attrs,
                        children: Vec::new(),
                    });
                }
                Event::Empty(e) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let mut attrs = Vec::new();
                    for a in e.attributes() {
                        let a = a.map_err(|e| e.to_string())?;
                        let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                        let val = a
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(|e| e.to_string())?
                            .into_owned();
                        attrs.push((key, val));
                    }
                    let el = Element {
                        name,
                        attrs,
                        children: Vec::new(),
                    };
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(Node::Elem(el));
                    } else if root.is_none() {
                        root = Some(el);
                    }
                }
                Event::End(_) => {
                    if let Some(el) = stack.pop() {
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(Node::Elem(el));
                        } else {
                            root = Some(el);
                        }
                    }
                }
                Event::Text(e) => {
                    let raw = String::from_utf8_lossy(e.as_ref());
                    // fast path: no '&' => nothing to unescape, keep the single
                    // allocation instead of copy + re-copy
                    let t = if raw.contains('&') {
                        match quick_xml::escape::unescape(&raw) {
                            Ok(c) => c.into_owned(),
                            Err(_) => raw.into_owned(),
                        }
                    } else {
                        raw.into_owned()
                    };
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(Node::Text(t));
                    } else {
                        prolog.push_str(&t);
                    }
                }
                Event::CData(e) => {
                    let raw = String::from_utf8_lossy(e.as_ref()).to_string();
                    // keep CDATA content as raw text (will be escaped on serialize)
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(Node::Text(raw));
                    }
                }
                Event::Decl(e) => {
                    let s = String::from_utf8_lossy(e.as_ref()).to_string();
                    prolog.push_str("<?");
                    prolog.push_str(&s);
                    prolog.push_str("?>");
                }
                Event::Comment(e) => {
                    let s = String::from_utf8_lossy(e.as_ref()).to_string();
                    let comment = format!("<!--{}-->", s);
                    if let Some(parent) = stack.last_mut() {
                        // preserve comments as unescaped text via placeholder? simpler: drop into prolog only if outside root
                        parent.children.push(Node::Text(String::new()));
                    } else {
                        prolog.push_str(&comment);
                    }
                }
                Event::PI(e) => {
                    let s = String::from_utf8_lossy(e.as_ref()).to_string();
                    let pi = format!("<?{}?>", s);
                    if stack.is_empty() {
                        prolog.push_str(&pi);
                    }
                }
                Event::DocType(e) => {
                    let s = String::from_utf8_lossy(e.as_ref()).to_string();
                    prolog.push_str("<!DOCTYPE");
                    prolog.push_str(&s);
                    prolog.push('>');
                }
                Event::Eof => break,
            }
        }

        let root = root.ok_or_else(|| "no root element found".to_string())?;
        Ok(Document { prolog, root })
    }

    pub fn serialize(&self) -> String {
        self.serialize_with_capacity(0)
    }

    /// Serialize with a preallocated output capacity. Pass the source xml
    /// length when known: a parse→serialize round-trip stays close in size,
    /// so this avoids the doubling-growth reallocs on multi-MB documents.
    pub fn serialize_with_capacity(&self, cap: usize) -> String {
        let mut out = String::with_capacity(cap);
        out.push_str(&self.prolog);
        self.root.serialize(&mut out);
        out
    }
}

/// A new XML declaration helper (unused warnings guard)
#[allow(dead_code)]
pub fn decl(version: &str, encoding: &str, standalone: bool) -> String {
    let d = BytesDecl::new(version, Some(encoding), if standalone { Some("yes") } else { None });
    format!("<?{}?>", String::from_utf8_lossy(d.as_ref()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn el(name: &str) -> Element {
        Element::new(name)
    }

    // ---------- Element::new / get_attr / set_attr ----------

    #[test]
    fn test_new_element_has_no_attrs_or_children() {
        let e = el("w:p");
        assert_eq!(e.name, "w:p");
        assert!(e.attrs.is_empty());
        assert!(e.children.is_empty());
    }

    #[test]
    fn test_get_attr_returns_none_for_missing() {
        let e = el("a");
        assert_eq!(e.get_attr("nope"), None);
    }

    #[test]
    fn test_set_attr_appends_new_attribute() {
        let mut e = el("a");
        e.set_attr("x", "1");
        e.set_attr("y", "2");
        assert_eq!(e.get_attr("x"), Some("1"));
        assert_eq!(e.get_attr("y"), Some("2"));
        // insertion order is preserved
        assert_eq!(
            e.attrs,
            vec![("x".to_string(), "1".to_string()), ("y".to_string(), "2".to_string())]
        );
    }

    #[test]
    fn test_set_attr_replaces_existing_in_place() {
        let mut e = el("a");
        e.set_attr("x", "1");
        e.set_attr("y", "2");
        e.set_attr("x", "3");
        assert_eq!(e.get_attr("x"), Some("3"));
        // no duplicate key added, order unchanged
        assert_eq!(e.attrs.len(), 2);
        assert_eq!(e.attrs[0].0, "x");
        assert_eq!(e.attrs[1].0, "y");
    }

    // ---------- find / find_mut / find_all / iter_descendants ----------

    #[test]
    fn test_find_returns_first_direct_child_only() {
        let mut root = el("root");
        let mut first = el("item");
        first.set_attr("id", "1");
        let mut second = el("item");
        second.set_attr("id", "2");
        root.children.push(Node::Text("text".into()));
        root.children.push(Node::Elem(first));
        root.children.push(Node::Elem(second));

        let found = root.find("item").expect("should find first item");
        assert_eq!(found.get_attr("id"), Some("1"));
        // text nodes are skipped, missing name returns None
        assert!(root.find("text").is_none());
    }

    #[test]
    fn test_find_does_not_recurse_into_grandchildren() {
        let mut root = el("root");
        let mut child = el("child");
        child.children.push(Node::Elem(el("item")));
        root.children.push(Node::Elem(child));
        assert!(root.find("item").is_none());
    }

    #[test]
    fn test_find_mut_allows_in_place_mutation() {
        let mut root = el("root");
        root.children.push(Node::Elem(el("item")));
        root.find_mut("item").unwrap().set_attr("id", "9");
        assert_eq!(root.find("item").unwrap().get_attr("id"), Some("9"));
        assert!(root.find_mut("missing").is_none());
    }

    #[test]
    fn test_find_all_returns_all_direct_matches_in_order() {
        let mut root = el("root");
        for id in ["a", "b", "c"] {
            let mut item = el("item");
            item.set_attr("id", id);
            root.children.push(Node::Elem(item));
        }
        root.children.push(Node::Elem(el("other")));
        let ids: Vec<&str> = root
            .find_all("item")
            .iter()
            .map(|e| e.get_attr("id").unwrap())
            .collect();
        assert_eq!(ids, vec!["a", "b", "c"]);
        assert!(root.find_all("nope").is_empty());
    }

    #[test]
    fn test_iter_descendants_collects_nested_matches() {
        let doc = Document::parse("<root><a><b/><a><b/></a></a><b/></root>").unwrap();
        let mut out = Vec::new();
        doc.root.iter_descendants("b", &mut out);
        assert_eq!(out.len(), 3);
        // self is not included even if the name matches
        let mut out = Vec::new();
        doc.root.iter_descendants("root", &mut out);
        assert!(out.is_empty());
    }

    // ---------- Element::serialize escaping ----------

    #[test]
    fn test_serialize_escapes_text_special_chars() {
        let mut e = el("t");
        e.children.push(Node::Text(r#"a<b>c&d"e'f"#.into()));
        let mut out = String::new();
        e.serialize(&mut out);
        // text escapes & < >, but not quotes
        assert_eq!(out, r#"<t>a&lt;b&gt;c&amp;d"e'f</t>"#);
    }

    #[test]
    fn test_serialize_escapes_attr_special_chars() {
        let mut e = el("t");
        e.set_attr("v", r#"a<b&c"d>e'f"#);
        let mut out = String::new();
        e.serialize(&mut out);
        // attrs escape & < ", but not > or '
        assert_eq!(out, r#"<t v="a&lt;b&amp;c&quot;d>e'f"/>"#);
    }

    #[test]
    fn test_serialize_empty_element_is_self_closing() {
        let e = el("w:br");
        let mut out = String::new();
        e.serialize(&mut out);
        assert_eq!(out, "<w:br/>");
    }

    #[test]
    fn test_serialize_nested_structure_and_attr_order() {
        let mut root = el("w:tc");
        root.set_attr("w:val", "x");
        let mut p = el("w:p");
        p.children.push(Node::Text("hi".into()));
        root.children.push(Node::Elem(p));
        let mut out = String::new();
        root.serialize(&mut out);
        assert_eq!(out, r#"<w:tc w:val="x"><w:p>hi</w:p></w:tc>"#);
    }

    // ---------- Document::parse / serialize round-trip ----------

    #[test]
    fn test_parse_roundtrip_simple_document() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root a="1"><child>text</child></root>"#;
        let doc = Document::parse(xml).unwrap();
        assert_eq!(
            doc.prolog,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
        );
        assert_eq!(doc.root.name, "root");
        assert_eq!(doc.serialize(), xml);
    }

    #[test]
    fn test_parse_text_entities_are_decoded() {
        let doc = Document::parse("<t>a&lt;b&gt;c&amp;d&quot;e&apos;f</t>").unwrap();
        assert_eq!(doc.root.text_content(), "a<b>c&d\"e'f");
    }

    #[test]
    fn test_parse_numeric_char_references_are_decoded() {
        let doc = Document::parse("<t>&#65;&#x42;</t>").unwrap();
        assert_eq!(doc.root.text_content(), "AB");
    }

    #[test]
    fn test_parse_entity_roundtrip_in_text() {
        let xml = "<t>a&lt;b&amp;c</t>";
        let doc = Document::parse(xml).unwrap();
        assert_eq!(doc.serialize(), xml);
    }

    #[test]
    fn test_parse_attr_entities_are_decoded_and_reescaped() {
        let doc = Document::parse(r#"<t a="&lt;&amp;&quot;&gt;&apos;"/>"#).unwrap();
        assert_eq!(doc.root.get_attr("a"), Some("<&\">'"));
        // serialize re-escapes " as &quot;, but ' and > stay literal
        assert_eq!(doc.serialize(), r#"<t a="&lt;&amp;&quot;>'"/>"#);
    }

    #[test]
    fn test_parse_namespace_prefixes_treated_literally() {
        let xml = r#"<w:document xmlns:w="urn:x"><w:body w:rsidR="00AB"/></w:document>"#;
        let doc = Document::parse(xml).unwrap();
        assert_eq!(doc.root.name, "w:document");
        assert_eq!(doc.root.get_attr("xmlns:w"), Some("urn:x"));
        let body = doc.root.find("w:body").unwrap();
        assert_eq!(body.get_attr("w:rsidR"), Some("00AB"));
        assert_eq!(doc.serialize(), xml);
    }

    #[test]
    fn test_parse_self_closing_and_explicit_empty_tags() {
        let doc = Document::parse("<root><a/><b></b></root>").unwrap();
        assert!(doc.root.find("a").unwrap().children.is_empty());
        assert!(doc.root.find("b").unwrap().children.is_empty());
        // <b></b> has no children, so it serializes back self-closed
        assert_eq!(doc.serialize(), "<root><a/><b/></root>");
    }

    #[test]
    fn test_parse_deeply_nested_structure() {
        let xml = "<a><b><c><d>deep</d></c></b></a>";
        let doc = Document::parse(xml).unwrap();
        let d = doc
            .root
            .find("b")
            .and_then(|b| b.find("c"))
            .and_then(|c| c.find("d"))
            .unwrap();
        assert_eq!(d.text_content(), "deep");
        assert_eq!(doc.serialize(), xml);
    }

    #[test]
    fn test_parse_cdata_content_kept_as_raw_text() {
        let doc = Document::parse("<t><![CDATA[a<b&c]]></t>").unwrap();
        assert_eq!(doc.root.text_content(), "a<b&c");
        // CDATA markers are lost; content re-serializes as escaped text
        assert_eq!(doc.serialize(), "<t>a&lt;b&amp;c</t>");
    }

    #[test]
    fn test_parse_comment_before_root_goes_to_prolog() {
        let doc = Document::parse("<!--note--><root/>").unwrap();
        assert_eq!(doc.prolog, "<!--note-->");
        assert_eq!(doc.serialize(), "<!--note--><root/>");
    }

    #[test]
    fn test_parse_comment_inside_root_is_dropped() {
        let doc = Document::parse("<root><!--note--><a/></root>").unwrap();
        // comment inside root becomes an empty text node
        assert_eq!(doc.serialize(), "<root><a/></root>");
    }

    #[test]
    fn test_parse_pi_before_root_goes_to_prolog() {
        let doc = Document::parse("<?xml-stylesheet href=\"x\"?><root/>").unwrap();
        assert_eq!(doc.prolog, r#"<?xml-stylesheet href="x"?>"#);
        assert_eq!(doc.serialize(), r#"<?xml-stylesheet href="x"?><root/>"#);
    }

    #[test]
    fn test_parse_doctype_goes_to_prolog() {
        let doc = Document::parse("<!DOCTYPE html><root/>").unwrap();
        // NOTE: quick-xml strips the leading whitespace of the doctype content and
        // parse() rejoins without a space, so "<!DOCTYPE html>" does not round-trip
        assert_eq!(doc.prolog, "<!DOCTYPEhtml>");
        assert_eq!(doc.serialize(), "<!DOCTYPEhtml><root/>");
    }

    #[test]
    fn test_parse_whitespace_text_between_elements_preserved() {
        let xml = "<root> <a/> </root>";
        let doc = Document::parse(xml).unwrap();
        assert_eq!(doc.serialize(), xml);
    }

    #[test]
    fn test_text_content_concatenates_all_descendants() {
        let doc = Document::parse("<r>x<a>y</a>z<b><c>w</c></b></r>").unwrap();
        assert_eq!(doc.root.text_content(), "xyzw");
    }

    // ---------- error / edge paths ----------

    #[test]
    fn test_parse_mismatched_end_tag_errors() {
        let err = Document::parse("<a></b>").unwrap_err();
        assert!(!err.is_empty());
    }

    #[test]
    fn test_parse_unclosed_tag_errors_or_no_root() {
        // either quick-xml reports ill-formed, or the stack is never popped
        assert!(Document::parse("<a>").is_err());
    }

    #[test]
    fn test_parse_empty_input_errors_no_root() {
        let err = Document::parse("").unwrap_err();
        assert_eq!(err, "no root element found");
    }

    #[test]
    fn test_parse_prolog_only_errors_no_root() {
        let err = Document::parse(r#"<?xml version="1.0"?>"#).unwrap_err();
        assert_eq!(err, "no root element found");
    }

    #[test]
    fn test_parse_bare_text_errors_no_root() {
        assert!(Document::parse("not xml").is_err());
    }

    #[test]
    fn test_parse_unterminated_attr_value_errors() {
        assert!(Document::parse("<a x=\"1>").is_err());
    }

    #[test]
    fn test_parse_second_root_is_silently_dropped() {
        // documents current lenient behavior: only the first root is kept
        let doc = Document::parse("<a/><b/>").unwrap();
        assert_eq!(doc.root.name, "a");
        assert_eq!(doc.serialize(), "<a/>");
    }

    // ---------- decl helper ----------

    #[test]
    fn test_decl_standalone_yes() {
        assert_eq!(
            decl("1.0", "UTF-8", true),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
        );
    }

    #[test]
    fn test_decl_without_standalone() {
        assert_eq!(decl("1.0", "UTF-8", false), r#"<?xml version="1.0" encoding="UTF-8"?>"#);
    }
}
