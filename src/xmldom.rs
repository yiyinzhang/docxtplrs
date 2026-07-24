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
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(c),
        }
    }
}

fn escape_attr(s: &str, out: &mut String) {
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(c),
        }
    }
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
                    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    let mut attrs = Vec::new();
                    for a in e.attributes() {
                        let a = a.map_err(|e| e.to_string())?;
                        let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
                        let val = a
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(|e| e.to_string())?
                            .to_string();
                        attrs.push((key, val));
                    }
                    stack.push(Element {
                        name,
                        attrs,
                        children: Vec::new(),
                    });
                }
                Event::Empty(e) => {
                    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    let mut attrs = Vec::new();
                    for a in e.attributes() {
                        let a = a.map_err(|e| e.to_string())?;
                        let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
                        let val = a
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(|e| e.to_string())?
                            .to_string();
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
                    let raw = String::from_utf8_lossy(e.as_ref()).to_string();
                    let t = quick_xml::escape::unescape(&raw)
                        .map(|c| c.to_string())
                        .unwrap_or(raw);
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
        let mut out = self.prolog.clone();
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
