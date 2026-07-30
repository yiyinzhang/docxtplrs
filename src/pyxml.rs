//! Generic live XML element proxy (`XmlElement`), the docxtplrs equivalent of
//! the lxml `element` escape hatch in python-docx / docxtpl.
//!
//! Since docxtplrs keeps its DOM in Rust, an `XmlElement` is a *stateless live
//! proxy* addressed by (part path, child-element index path). Every operation
//! re-parses (or reuses the cached) part, navigates to the element and, for
//! mutations, writes the part back.

use crate::docmodel::with_core;
use crate::pyclasses::PyDocxTemplate;
use crate::template::TplCore;
use crate::xmldom::{Document, Element, Node};
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

fn py_err(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

/// An XML element inside a docx part (live proxy, lxml-inspired).
#[pyclass(name = "XmlElement", unsendable, skip_from_py_object)]
pub struct PyXmlElement {
    pub tpl: Py<PyDocxTemplate>,
    /// Part path inside the package, e.g. "word/settings.xml".
    pub part: String,
    /// Path from the part root: each entry is the index among *element*
    /// children (text nodes are skipped).
    pub path: Vec<usize>,
}

fn nav<'a>(root: &'a Element, path: &[usize]) -> Option<&'a Element> {
    let mut cur = root;
    for &i in path {
        cur = cur
            .children
            .iter()
            .filter_map(|c| match c {
                Node::Elem(e) => Some(e),
                _ => None,
            })
            .nth(i)?;
    }
    Some(cur)
}

fn nav_mut<'a>(root: &'a mut Element, path: &[usize]) -> Option<&'a mut Element> {
    let mut cur = root;
    for &i in path {
        cur = cur
            .children
            .iter_mut()
            .filter_map(|c| match c {
                Node::Elem(e) => Some(e),
                _ => None,
            })
            .nth(i)?;
    }
    Some(cur)
}

/// Read access to the addressed element.
fn read_el<R>(
    core: &mut TplCore,
    part: &str,
    path: &[usize],
    f: impl FnOnce(&Element) -> R,
) -> Result<R, String> {
    let dom = core.part_dom(part)?;
    let el = nav(&dom.root, path).ok_or_else(|| "element not found".to_string())?;
    Ok(f(el))
}

/// Mutating access to the addressed element; the cached DOM is written back
/// to the package on the next flush (render/save/etc).
fn edit_el<R>(
    core: &mut TplCore,
    part: &str,
    path: &[usize],
    f: impl FnOnce(&mut Element) -> R,
) -> Result<R, String> {
    let r = {
        let dom = core.part_dom(part)?;
        let el = nav_mut(&mut dom.root, path).ok_or_else(|| "element not found".to_string())?;
        f(el)
    };
    core.mark_part_dirty(part);
    Ok(r)
}

impl PyXmlElement {
    fn child(&self, py: Python<'_>, idx: usize) -> PyXmlElement {
        let mut path = self.path.clone();
        path.push(idx);
        PyXmlElement {
            tpl: self.tpl.clone_ref(py),
            part: self.part.clone(),
            path,
        }
    }

    fn read<R>(&self, py: Python<'_>, f: impl FnOnce(&Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            read_el(core, &self.part, &self.path, f).map_err(py_err)
        })
    }

    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut Element) -> R) -> PyResult<R> {
        with_core(&self.tpl, py, |core| {
            edit_el(core, &self.part, &self.path, f).map_err(py_err)
        })
    }
}

/// Parse an XML fragment (single element) or clone another XmlElement.
fn fragment_element(py: Python<'_>, obj: &Bound<'_, PyAny>) -> PyResult<Element> {
    if let Ok(b) = obj.cast::<PyXmlElement>() {
        // clone from its source of truth so cross-part/cross-template works
        let (tpl, part, path) = {
            let xel = b.borrow();
            (xel.tpl.clone_ref(py), xel.part.clone(), xel.path.clone())
        };
        return with_core(&tpl, py, |core| {
            read_el(core, &part, &path, |e| e.clone()).map_err(py_err)
        });
    }
    if let Ok(xml) = obj.extract::<String>() {
        let dom = Document::parse(&xml)
            .map_err(|e| PyValueError::new_err(format!("invalid XML fragment: {}", e)))?;
        return Ok(dom.root);
    }
    Err(PyValueError::new_err(
        "expected an XML string or an XmlElement",
    ))
}

#[pymethods]
impl PyXmlElement {
    /// The element tag, e.g. "w:settings".
    #[getter]
    fn tag(&self, py: Python<'_>) -> PyResult<String> {
        self.read(py, |e| e.name.to_string())
    }

    /// Serialized XML of this element (including children).
    #[getter]
    fn xml(&self, py: Python<'_>) -> PyResult<String> {
        self.read(py, |e| {
            let mut s = String::new();
            e.serialize(&mut s);
            s
        })
    }

    /// Concatenated text of all descendant text nodes.
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.read(py, |e| e.text_content())
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |e| {
            e.children.clear();
            e.children.push(Node::Text(v));
        })
    }

    /// Attributes as a dict {name: value}.
    #[getter]
    fn attrib(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let pairs = self.read(py, |e| e.attrs.clone())?;
        let d = PyDict::new(py);
        for (k, v) in pairs {
            d.set_item(k.to_string(), v)?;
        }
        Ok(d.unbind())
    }

    /// Get an attribute (None if absent).
    fn get(&self, py: Python<'_>, name: &str) -> PyResult<Option<String>> {
        self.read(py, |e| e.get_attr(name).map(|s| s.to_string()))
    }

    /// Set an attribute.
    fn set(&self, py: Python<'_>, name: &str, value: &str) -> PyResult<()> {
        self.edit(py, |e| e.set_attr(name, value))
    }

    /// Remove an attribute if present.
    fn remove_attr(&self, py: Python<'_>, name: &str) -> PyResult<()> {
        self.edit(py, |e| e.attrs.retain(|(k, _)| k != name))
    }

    /// First direct child element with this tag (None if absent).
    fn find(&self, py: Python<'_>, name: &str) -> PyResult<Option<PyXmlElement>> {
        let idx = self.read(py, |e| {
            e.children
                .iter()
                .filter_map(|c| match c {
                    Node::Elem(x) => Some(x),
                    _ => None,
                })
                .position(|x| x.name == name)
        })?;
        Ok(idx.map(|i| self.child(py, i)))
    }

    /// All direct child elements with this tag.
    fn findall(&self, py: Python<'_>, name: &str) -> PyResult<Vec<PyXmlElement>> {
        let idxs: Vec<usize> = self.read(py, |e| {
            e.children
                .iter()
                .filter_map(|c| match c {
                    Node::Elem(x) => Some(x),
                    _ => None,
                })
                .enumerate()
                .filter_map(|(i, x)| if x.name == name { Some(i) } else { None })
                .collect()
        })?;
        Ok(idxs.into_iter().map(|i| self.child(py, i)).collect())
    }

    /// All direct child elements.
    #[getter]
    fn children(&self, py: Python<'_>) -> PyResult<Vec<PyXmlElement>> {
        let n = self.read(py, |e| {
            e.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(_)))
                .count()
        })?;
        Ok((0..n).map(|i| self.child(py, i)).collect())
    }

    /// Append a child: an XML fragment string or another XmlElement.
    fn append(&self, py: Python<'_>, obj: &Bound<'_, PyAny>) -> PyResult<()> {
        let frag = fragment_element(py, obj)?;
        self.edit(py, |e| e.children.push(Node::Elem(frag)))
    }

    /// Insert a child at position `index` (counting element children only).
    fn insert(&self, py: Python<'_>, index: usize, obj: &Bound<'_, PyAny>) -> PyResult<()> {
        let frag = fragment_element(py, obj)?;
        self.edit(py, |e| {
            // position in `children` of the index-th element child
            let pos = e
                .children
                .iter()
                .enumerate()
                .filter(|(_, c)| matches!(c, Node::Elem(_)))
                .map(|(i, _)| i)
                .nth(index)
                .unwrap_or(e.children.len());
            e.children.insert(pos, Node::Elem(frag));
        })
    }

    /// Remove a direct child previously obtained from this element.
    fn remove(&self, py: Python<'_>, child: Bound<'_, PyXmlElement>) -> PyResult<()> {
        let (cpart, cpath) = {
            let c = child.borrow();
            (c.part.clone(), c.path.clone())
        };
        if cpart != self.part
            || cpath.len() != self.path.len() + 1
            || cpath[..self.path.len()] != self.path[..]
        {
            return Err(PyValueError::new_err(
                "not a direct child of this element",
            ));
        }
        let idx = cpath[self.path.len()];
        self.edit(py, |e| {
            // position in `children` of the idx-th element child
            if let Some(pos) = e
                .children
                .iter()
                .enumerate()
                .filter(|(_, c)| matches!(c, Node::Elem(_)))
                .map(|(i, _)| i)
                .nth(idx)
            {
                e.children.remove(pos);
            }
        })
    }

    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.read(py, |e| {
            e.children
                .iter()
                .filter(|c| matches!(c, Node::Elem(_)))
                .count()
        })
    }

    fn __str__(&self, py: Python<'_>) -> PyResult<String> {
        self.xml(py)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!("<XmlElement {} of {}>", self.tag(py)?, self.part))
    }
}
