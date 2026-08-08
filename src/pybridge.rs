//! Bridge between Python objects and minijinja values.

use crate::template::TplCore;
use std::cell::RefCell;

thread_local! {
    /// (core pointer, current part) while a template is being rendered.
    /// Used to route lazy conversions (InlineImage/Subdoc materialization)
    /// to the right TplCore even from deeply nested object access.
    static CURRENT_RENDER: RefCell<Option<(*mut TplCore, String)>> = const { RefCell::new(None) };
}

/// Set the current render context; returns a guard value to restore later.
pub fn set_current_render(core: *mut TplCore, part: &str) -> Option<(*mut TplCore, String)> {
    CURRENT_RENDER.with(|c| c.borrow_mut().replace((core, part.to_string())))
}

pub fn restore_current_render(prev: Option<(*mut TplCore, String)>) {
    CURRENT_RENDER.with(|c| *c.borrow_mut() = prev);
}

/// Name of the part currently being rendered on this thread
/// (docxtpl current_rendering_part), None outside a render.
pub fn current_rendering_part() -> Option<String> {
    CURRENT_RENDER.with(|c| c.borrow().as_ref().map(|(_, part)| part.clone()))
}

/// Access the current render core if set.
fn with_current_core<R>(f: impl FnOnce(&mut TplCore, &str) -> R) -> Option<R> {
    CURRENT_RENDER.with(|c| {
        let borrowed = c.borrow();
        borrowed.as_ref().map(|(ptr, part)| {
            // SAFETY: the pointer is valid for the duration of the render call
            // on this thread (set in render_part / render_properties).
            let core = unsafe { &mut **ptr };
            f(core, part)
        })
    })
}
use minijinja::value::{Enumerator, Object, ObjectRepr, Value};
use minijinja::{Error, ErrorKind};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple};
use std::fmt;
use std::sync::Arc;

use crate::pyclasses::{PyInlineImage, PyListing, PyRichText, PySubdoc};

/// Register a deferred value and return its placeholder token.
fn register_deferred(tpl: &mut TplCore, oid: usize, d: crate::template::Deferred) -> String {
    let idx = match tpl.deferred_by_oid.get(&oid) {
        Some(&i) => i,
        None => {
            let i = tpl.deferred.len();
            tpl.deferred.push(d);
            tpl.deferred_by_oid.insert(oid, i);
            i
        }
    };
    crate::template::deferred_token(idx)
}

/// Convert a Python object into a minijinja Value.
///
/// `tpl`/`part` are needed to materialize InlineImage / Subdoc values for the
/// part currently being rendered.
pub fn py_to_value(
    obj: &Bound<'_, PyAny>,
    tpl: &mut TplCore,
    part: &str,
    depth: usize,
) -> PyResult<Value> {
    if depth > 64 {
        return Ok(Value::from(()));
    }
    // fast path first: plain scalars are the bulk of all conversions and
    // can never be one of our custom pyclasses (checked below), so dispatch
    // them without the failed casts (hot: PyWrapper::get_value per lookup)
    if obj.is_none() {
        // jinja2 renders None as "None"; wrap so display matches while
        // keeping falsy semantics
        return Ok(Value::from_object(PyNoneObj));
    }
    if obj.is_instance_of::<PyBool>() {
        // jinja2 renders booleans as True/False
        let b = obj.extract::<bool>()?;
        return Ok(Value::from_object(PyBoolObj(b)));
    }
    if obj.is_instance_of::<PyInt>() {
        if let Ok(i) = obj.extract::<i64>() {
            return Ok(Value::from(i));
        }
        if let Ok(i) = obj.extract::<i128>() {
            return Ok(Value::from(i));
        }
        // very large ints: fall back to string
        let s = obj.str()?.to_string_lossy().to_string();
        return Ok(Value::from(s));
    }
    if obj.is_instance_of::<PyFloat>() {
        let f = obj.extract::<f64>()?;
        return Ok(Value::from(f));
    }
    if obj.is_instance_of::<PyString>() {
        let s = obj.extract::<String>()?;
        return Ok(Value::from(s));
    }
    if obj.is_instance_of::<PyBytes>() {
        let b = obj.cast::<PyBytes>()?;
        let s = String::from_utf8_lossy(b.as_bytes()).to_string();
        return Ok(Value::from(s));
    }
    // our own safe-xml types next
    if let Ok(rt) = obj.cast::<PyRichText>() {
        return Ok(Value::from_safe_string(rt.borrow().xml.borrow().clone()));
    }
    if let Ok(rt) = obj.cast::<crate::pyclasses::PyRichTextParagraph>() {
        return Ok(Value::from_safe_string(rt.borrow().xml.borrow().clone()));
    }
    if let Ok(l) = obj.cast::<PyListing>() {
        return Ok(Value::from_safe_string(l.borrow().xml.borrow().clone()));
    }
    if let Ok(img) = obj.cast::<PyInlineImage>() {
        let oid = obj.as_ptr() as usize;
        // cheap path: already registered for this render — don't rebuild
        // (and re-clone) the Deferred
        if let Some(&i) = tpl.deferred_by_oid.get(&oid) {
            return Ok(Value::from_safe_string(crate::template::deferred_token(i)));
        }
        let d = {
            let b = img.borrow();
            crate::template::Deferred::Image {
                blob: b.blob.clone(),
                filename: b.filename.clone(),
                width: b.width,
                height: b.height,
                anchor: b.anchor.clone(),
                title: b.title.clone(),
                descr: b.descr.clone(),
            }
        };
        return Ok(Value::from_safe_string(register_deferred(tpl, oid, d)));
    }
    if let Ok(sd) = obj.cast::<PySubdoc>() {
        let oid = obj.as_ptr() as usize;
        if let Some(&i) = tpl.deferred_by_oid.get(&oid) {
            return Ok(Value::from_safe_string(crate::template::deferred_token(i)));
        }
        let b = sd.borrow();
        let d = match &b.bytes {
            Some(bytes) => crate::template::Deferred::Subdoc {
                bytes: Some(std::sync::Arc::from(&bytes[..])),
                keep_sections: b.keep_sections,
            },
            None => crate::template::Deferred::SubdocBlocks {
                blocks: std::sync::Arc::new(b.blocks.borrow().clone()),
            },
        };
        return Ok(Value::from_safe_string(register_deferred(tpl, oid, d)));
    }
    if let Ok(dict) = obj.cast::<PyDict>() {
        if depth == 0 {
            // Top-level render context: jinja2 resolves names via the
            // mapping (namespace semantics), so use a native map.
            let mut map = std::collections::BTreeMap::new();
            for (k, v) in dict.iter() {
                let key = if let Ok(s) = k.extract::<String>() {
                    s
                } else {
                    k.str()?.to_string_lossy().to_string()
                };
                map.insert(key, py_to_value(&v, tpl, part, depth + 1)?);
            }
            return Ok(Value::from_serialize(&map));
        }
        // Nested dicts are wrapped lazily: preserves insertion order on
        // iteration (jinja2 semantics) and defers nested materialization.
        return Ok(Value::from_object(PyWrapper::new(obj.clone().unbind())));
    }
    if obj.is_instance_of::<PyList>() || obj.is_instance_of::<PyTuple>() {
        // sequences are wrapped lazily (jinja2 repr / iteration semantics)
        return Ok(Value::from_object(PyWrapper::new(obj.clone().unbind())));
    }
    // honor the __html__ protocol (jinja2 treats such objects as safe)
    if let Ok(html) = obj.call_method0("__html__") {
        if let Ok(s) = html.extract::<String>() {
            return Ok(Value::from_safe_string(s));
        }
    }
    // arbitrary python object: wrap for lazy attribute/method access
    Ok(Value::from_object(PyWrapper::new(obj.clone().unbind())))
}

/// Convert a minijinja value back to a Python object (for call args).
pub fn value_to_py<'py>(py: Python<'py>, v: &Value) -> PyResult<Bound<'py, PyAny>> {
    use minijinja::value::ValueKind;
    // unwrap our own wrappers back to the original Python object
    if let Some(object) = v.as_object() {
        if let Some(wrapper) = object.downcast_ref::<PyWrapper>() {
            return Ok(wrapper.obj.bind(py).clone());
        }
        if let Some(b) = object.downcast_ref::<PyBoolObj>() {
            return Ok(PyBool::new(py, b.0).to_owned().into_any());
        }
        if object.downcast_ref::<PyNoneObj>().is_some() {
            return Ok(py.None().into_bound(py));
        }
    }
    match v.kind() {
        ValueKind::Undefined | ValueKind::None => Ok(py.None().into_bound(py)),
        ValueKind::Bool => {
            let b = v.is_true();
            Ok(PyBool::new(py, b).to_owned().into_any())
        }
        ValueKind::Number => {
            if let Some(i) = v.as_i64() {
                Ok(i.into_pyobject(py)?.into_any())
            } else if let Ok(f) = f64::try_from(v.clone()) {
                Ok(f.into_pyobject(py)?.into_any())
            } else {
                Ok(v.to_string().into_pyobject(py)?.into_any())
            }
        }
        ValueKind::String | ValueKind::Bytes => {
            Ok(v.to_string().into_pyobject(py)?.into_any())
        }
        ValueKind::Seq => {
            let list = PyList::empty(py);
            if let Ok(iter) = v.try_iter() {
                for item in iter {
                    list.append(value_to_py(py, &item)?)?;
                }
            }
            Ok(list.into_any())
        }
        ValueKind::Map => {
            let dict = PyDict::new(py);
            if let Ok(iter) = v.try_iter() {
                for k in iter {
                    let kv = value_to_py(py, &k)?;
                    if let Ok(val) = v.get_item(&k) {
                        dict.set_item(kv, value_to_py(py, &val)?)?;
                    }
                }
            }
            Ok(dict.into_any())
        }
        _ => {
            // lazily-concatenated sequences (MergeSeq) and other iterables
            // must materialize as lists, not strings
            if let Ok(iter) = v.try_iter() {
                let list = PyList::empty(py);
                for item in iter {
                    list.append(value_to_py(py, &item)?)?;
                }
                Ok(list.into_any())
            } else {
                Ok(v.to_string().into_pyobject(py)?.into_any())
            }
        }
    }
}

/// Python bool with jinja2 display semantics (True/False).
#[derive(Debug, Clone, Copy)]
pub struct PyBoolObj(pub bool);

impl Object for PyBoolObj {
    fn repr(self: &Arc<Self>) -> ObjectRepr {
        ObjectRepr::Plain
    }

    fn is_true(self: &Arc<Self>) -> bool {
        self.0
    }

    fn custom_cmp(self: &Arc<Self>, other: &minijinja::value::DynObject) -> Option<std::cmp::Ordering> {
        let other = other.downcast_ref::<PyBoolObj>()?;
        Some(self.0.cmp(&other.0))
    }

    fn render(self: &Arc<Self>, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(if self.0 { "True" } else { "False" })
    }
}

/// Python None with jinja2 display semantics ("None").
#[derive(Debug, Clone, Copy)]
pub struct PyNoneObj;

impl Object for PyNoneObj {
    fn repr(self: &Arc<Self>) -> ObjectRepr {
        ObjectRepr::Plain
    }

    fn is_true(self: &Arc<Self>) -> bool {
        false
    }

    fn custom_cmp(self: &Arc<Self>, other: &minijinja::value::DynObject) -> Option<std::cmp::Ordering> {
        other.downcast_ref::<PyNoneObj>()?;
        Some(std::cmp::Ordering::Equal)
    }

    fn render(self: &Arc<Self>, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("None")
    }
}

/// Wrapper for arbitrary Python objects used from templates.
#[derive(Debug)]
pub struct PyWrapper {
    pub obj: Py<PyAny>,
    /// dict item cache: materialized on the first item lookup (one GIL attach
    /// for the whole dict instead of one attach per attribute access —
    /// loop-heavy templates look up thousands of attributes). Frozen once
    /// built: a dict mutated mid-render keeps its first-seen values. Values
    /// are Option to mirror convert_shallow failures (lookup -> undefined).
    dict_cache: std::sync::Mutex<Option<std::collections::HashMap<String, Option<Value>>>>,
}

impl PyWrapper {
    fn new(obj: Py<PyAny>) -> PyWrapper {
        PyWrapper {
            obj,
            dict_cache: std::sync::Mutex::new(None),
        }
    }
}

impl fmt::Display for PyWrapper {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let s = Python::attach(|py| {
            self.obj
                .bind(py)
                .str()
                .map(|s| s.to_string_lossy().to_string())
                .unwrap_or_default()
        });
        f.write_str(&s)
    }
}

/// Convert a global Python object to a Value (for env globals).
pub fn py_to_value_global(obj: &Py<PyAny>) -> Value {
    Python::attach(|py| {
        let b = obj.bind(py);
        if let Some(v) = with_current_core(|core, part| py_to_value(b, core, part, 1).ok()) {
            return v.unwrap_or_else(|| Value::from(()));
        }
        let mut core = TplCore::default();
        py_to_value(b, &mut core, "word/document.xml", 1).unwrap_or_else(|_| Value::from(()))
    })
}

/// Convert a Python call result back to a Value (filters/functions).
pub fn py_to_value_render(obj: &Bound<'_, PyAny>) -> PyResult<Value> {
    if let Some(v) = with_current_core(|core, part| py_to_value(obj, core, part, 1).ok()) {
        return v.ok_or_else(|| {
            pyo3::exceptions::PyRuntimeError::new_err("conversion failed")
        });
    }
    let mut core = TplCore::default();
    py_to_value(obj, &mut core, "word/document.xml", 1)
}

fn convert_shallow(obj: &Bound<'_, PyAny>) -> Option<Value> {
    // fast path: scalars and containers need no render core (only the
    // deferred-value types do), so skip the TLS lookup + full dispatch
    if obj.is_none() {
        return Some(Value::from_object(PyNoneObj));
    }
    if obj.is_instance_of::<PyBool>() {
        return Some(Value::from_object(PyBoolObj(obj.extract::<bool>().ok()?)));
    }
    if obj.is_instance_of::<PyInt>() {
        if let Ok(i) = obj.extract::<i64>() {
            return Some(Value::from(i));
        }
        if let Ok(i) = obj.extract::<i128>() {
            return Some(Value::from(i));
        }
        return Some(Value::from(obj.str().ok()?.to_string_lossy().to_string()));
    }
    if obj.is_instance_of::<PyFloat>() {
        return Some(Value::from(obj.extract::<f64>().ok()?));
    }
    if obj.is_instance_of::<PyString>() {
        return Some(Value::from(obj.extract::<String>().ok()?));
    }
    if obj.is_instance_of::<PyBytes>() {
        let b = obj.cast::<PyBytes>().ok()?;
        return Some(Value::from(String::from_utf8_lossy(b.as_bytes()).to_string()));
    }
    if obj.is_instance_of::<PyDict>() || obj.is_instance_of::<PyList>() || obj.is_instance_of::<PyTuple>() {
        return Some(Value::from_object(PyWrapper::new(obj.clone().unbind())));
    }
    // route through the active render core when available (so that deferred
    // values like InlineImage register correctly), otherwise use a scratch core
    if let Some(v) = with_current_core(|core, part| py_to_value(obj, core, part, 1).ok()) {
        return v;
    }
    let mut core = TplCore::default();
    py_to_value(obj, &mut core, "word/document.xml", 1).ok()
}

impl Object for PyWrapper {
    fn repr(self: &Arc<Self>) -> ObjectRepr {
        // shape the object kind so that containment (`in`), sorting and
        // iteration work like jinja2 expects for dicts / sequences
        Python::attach(|py| {
            let o = self.obj.bind(py);
            if o.cast::<PyDict>().is_ok() {
                ObjectRepr::Map
            } else if o.cast::<PyList>().is_ok() || o.cast::<PyTuple>().is_ok() {
                ObjectRepr::Seq
            } else if o.getattr("__iter__").is_ok() {
                ObjectRepr::Iterable
            } else {
                ObjectRepr::Plain
            }
        })
    }

    fn get_value(self: &Arc<Self>, key: &Value) -> Option<Value> {
        Python::attach(|py| {
            let obj = self.obj.bind(py);
            if let Some(name) = key.as_str() {
                if let Ok(dict) = obj.cast::<PyDict>() {
                    // dicts: item access first. The whole dict is materialized
                    // once (single attach) and later lookups are cache hits
                    // with no GIL round-trip at all.
                    {
                        let mut c = self.dict_cache.lock().unwrap();
                        if c.is_none() {
                            let mut m = std::collections::HashMap::with_capacity(dict.len());
                            for (k, v) in dict.iter() {
                                if let Ok(ks) = k.extract::<String>() {
                                    m.insert(ks, convert_shallow(&v));
                                }
                            }
                            *c = Some(m);
                        }
                    }
                    let hit = self
                        .dict_cache
                        .lock()
                        .unwrap()
                        .as_ref()
                        .unwrap()
                        .get(name)
                        .cloned();
                    match hit {
                        Some(Some(v)) => return Some(v),
                        // conversion failed at materialization: undefined
                        Some(None) => return None,
                        // no such item: fall back to attribute access (dict
                        // subclass attrs), like PyDict::get_item's miss
                        None => {
                            if let Ok(attr) = obj.getattr(name) {
                                return convert_shallow(&attr);
                            }
                        }
                    }
                } else {
                    // attribute access first (like jinja2 getattr)
                    if let Ok(attr) = obj.getattr(name) {
                        return convert_shallow(&attr);
                    }
                    if let Ok(item) = obj.get_item(name) {
                        return convert_shallow(&item);
                    }
                }
                None
            } else if let Some(i) = key.as_i64() {
                if let Ok(item) = obj.get_item(i) {
                    return convert_shallow(&item);
                }
                None
            } else {
                None
            }
        })
    }

    fn enumerate(self: &Arc<Self>) -> Enumerator {
        let items: Option<Vec<Value>> = Python::attach(|py| {
            let obj = self.obj.bind(py);
            let iter = obj.try_iter().ok()?;
            let mut out = Vec::new();
            for item in iter {
                out.push(convert_shallow(&item.ok()?)?);
            }
            Some(out)
        });
        match items {
            Some(v) => Enumerator::Iter(Box::new(v.into_iter())),
            None => Enumerator::Empty,
        }
    }

    fn enumerator_len(self: &Arc<Self>) -> Option<usize> {
        Python::attach(|py| self.obj.bind(py).len().ok())
    }

    fn is_true(self: &Arc<Self>) -> bool {
        Python::attach(|py| self.obj.bind(py).is_truthy().unwrap_or(true))
    }

    fn custom_cmp(self: &Arc<Self>, other: &minijinja::value::DynObject) -> Option<std::cmp::Ordering> {
        use std::cmp::Ordering;
        let other = other.downcast_ref::<PyWrapper>()?;
        Python::attach(|py| {
            let a = self.obj.bind(py);
            let b = other.obj.bind(py);
            use pyo3::basic::CompareOp;
            if let Ok(r) = a.rich_compare(b, CompareOp::Eq) {
                if r.is_truthy().unwrap_or(false) {
                    return Some(Ordering::Equal);
                }
            }
            if let Ok(r) = a.rich_compare(b, CompareOp::Lt) {
                if r.is_truthy().unwrap_or(false) {
                    return Some(Ordering::Less);
                }
            }
            if let Ok(r) = a.rich_compare(b, CompareOp::Gt) {
                if r.is_truthy().unwrap_or(false) {
                    return Some(Ordering::Greater);
                }
            }
            // total-order fallback (Python objects are not orderable here)
            Some((a.as_ptr() as usize).cmp(&(b.as_ptr() as usize)))
        })
    }

    fn call(self: &Arc<Self>, _state: &minijinja::State, args: &[Value]) -> Result<Value, Error> {
        Python::attach(|py| {
            let obj = self.obj.bind(py);
            let mut py_args = Vec::with_capacity(args.len());
            for a in args {
                py_args.push(
                    value_to_py(py, a)
                        .map_err(|e| Error::new(ErrorKind::InvalidOperation, e.to_string()))?,
                );
            }
            let tuple = PyTuple::new(py, py_args)
                .map_err(|e| Error::new(ErrorKind::InvalidOperation, e.to_string()))?;
            match obj.call1(tuple) {
                Ok(res) => convert_shallow(&res)
                    .ok_or_else(|| Error::new(ErrorKind::InvalidOperation, "conversion failed")),
                Err(e) => Err(Error::new(ErrorKind::InvalidOperation, e.to_string())),
            }
        })
    }

    fn call_method(
        self: &Arc<Self>,
        _state: &minijinja::State,
        name: &str,
        args: &[Value],
    ) -> Result<Value, Error> {
        Python::attach(|py| {
            let obj = self.obj.bind(py);
            let method = obj
                .getattr(name)
                .map_err(|e| Error::new(ErrorKind::UnknownMethod, e.to_string()))?;
            let mut py_args = Vec::with_capacity(args.len());
            for a in args {
                py_args.push(
                    value_to_py(py, a)
                        .map_err(|e| Error::new(ErrorKind::InvalidOperation, e.to_string()))?,
                );
            }
            let tuple = PyTuple::new(py, py_args)
                .map_err(|e| Error::new(ErrorKind::InvalidOperation, e.to_string()))?;
            match method.call1(tuple) {
                Ok(res) => convert_shallow(&res)
                    .ok_or_else(|| Error::new(ErrorKind::InvalidOperation, "conversion failed")),
                Err(e) => Err(Error::new(ErrorKind::InvalidOperation, e.to_string())),
            }
        })
    }

    fn render(self: &Arc<Self>, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self, f)
    }
}
