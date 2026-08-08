//! PyO3 classes mirroring the docxtpl Python API.

use crate::image::length as len;
use crate::package::crc32;
use crate::pybridge::py_to_value;
use crate::richtext::{self, TextProps};
use crate::template::TplCore;
use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PySet};
use std::cell::RefCell;
use std::collections::HashMap;

fn to_pyerr(e: String) -> PyErr {
    PyRuntimeError::new_err(e)
}

pyo3::create_exception!(docxtplrs, TemplateError, pyo3::exceptions::PyException);

/// Read bytes from a Python source: path str, bytes, or file-like object.
fn read_bytes_source(obj: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    if let Ok(b) = obj.cast::<PyBytes>() {
        return Ok(b.as_bytes().to_vec());
    }
    if let Ok(s) = obj.extract::<String>() {
        return std::fs::read(&s)
            .map_err(|e| PyValueError::new_err(format!("cannot read {}: {}", s, e)));
    }
    // os.PathLike
    if let Ok(fspath) = obj.call_method0("__fspath__") {
        if let Ok(s) = fspath.extract::<String>() {
            return std::fs::read(&s)
                .map_err(|e| PyValueError::new_err(format!("cannot read {}: {}", s, e)));
        }
    }
    // file-like
    if let Ok(data) = obj.call_method0("read") {
        if let Ok(b) = data.cast::<PyBytes>() {
            return Ok(b.as_bytes().to_vec());
        }
        if let Ok(s) = data.extract::<String>() {
            return Ok(s.into_bytes());
        }
    }
    Err(PyValueError::new_err(
        "expected a path, bytes, or file-like object",
    ))
}

fn py_truthy(obj: &Bound<'_, PyAny>) -> bool {
    obj.is_truthy().unwrap_or(false)
}

/// Duck-type a jinja2 Environment's options and customizations onto the core
/// (shared by render() and the exposed pipeline methods below).
/// Returns the env's autoescape flag.
fn import_jinja_env(core: &mut TplCore, env: &Bound<'_, PyAny>) -> PyResult<bool> {
    let mut autoescape = false;
    if let Ok(ae) = env.getattr("autoescape") {
        if ae.is_truthy().unwrap_or(false) {
            autoescape = true;
        }
    }
    // jinja2 environment options
    for (attr, slot) in [
        ("trim_blocks", 0u8),
        ("lstrip_blocks", 1u8),
        ("keep_trailing_newline", 2u8),
    ] {
        if let Ok(v) = env.getattr(attr) {
            let b = v.is_truthy().unwrap_or(false);
            match slot {
                0 => core.env_options.trim_blocks = Some(b),
                1 => core.env_options.lstrip_blocks = Some(b),
                _ => core.env_options.keep_trailing_newline = Some(b),
            }
        }
    }
    if let Ok(undefined_cls) = env.getattr("undefined") {
        if let Ok(name) = undefined_cls.getattr("__name__") {
            let behavior = match name.str()?.to_string_lossy().as_ref() {
                "ChainableUndefined" => "chainable",
                "StrictUndefined" => "strict",
                _ => "lenient",
            };
            core.env_options.undefined_behavior = Some(behavior.to_string());
        }
    }
    // jinja2 ships default globals (namespace, range, dict, ...) that
    // would shadow minijinja's native builtins. minijinja's own
    // `namespace` object is required for `{% set ns.attr = ... %}`,
    // so never let an env-provided global override these names.
    const SKIP_GLOBALS: &[&str] =
        &["namespace", "range", "dict", "cycler", "joiner", "lipsum"];
    // jinja2's own builtin filters/tests/globals are best handled by
    // minijinja's native implementations: importing them as plain
    // python callables would both shadow the faster native versions
    // and break undefined-value semantics (jinja2's builtins check
    // `isinstance(v, jinja2.Undefined)`, which never matches values
    // converted from minijinja). Detect builtins by object identity
    // against jinja2.defaults; only entries the user actually added
    // or overrode get imported. When jinja2 is not importable (e.g.
    // duck-typed fake environments) fall back to importing everything.
    let py = env.py();
    let jinja_defaults = PyModule::import(py, "jinja2.defaults").ok();
    let default_dict = |attr: &str| -> Option<Bound<'_, PyDict>> {
        jinja_defaults
            .as_ref()
            .and_then(|m| m.getattr(attr).ok())
            .and_then(|d| d.cast_into::<PyDict>().ok())
    };
    let default_filters = default_dict("DEFAULT_FILTERS");
    let default_globals = default_dict("DEFAULT_NAMESPACE");
    let default_tests = default_dict("DEFAULT_TESTS");
    for (attr, kind) in [("filters", 0u8), ("globals", 1u8), ("tests", 2u8)] {
        if let Ok(d) = env.getattr(attr) {
            if let Ok(d) = d.cast::<PyDict>() {
                for (k, v) in d.iter() {
                    let name = k.str()?.to_string_lossy().to_string();
                    if kind == 1 && SKIP_GLOBALS.contains(&name.as_str()) {
                        continue;
                    }
                    let defaults_dict = match kind {
                        0 => &default_filters,
                        1 => &default_globals,
                        _ => &default_tests,
                    };
                    if let Some(dd) = defaults_dict {
                        if let Ok(Some(dv)) = dd.get_item(&name) {
                            if dv.as_ptr() == v.as_ptr() {
                                // untouched jinja2 builtin -> native
                                continue;
                            }
                        }
                    }
                    if kind != 1 {
                        // jinja2's own builtin filters/tests are either
                        // @async_variant wrappers or @pass_environment /
                        // @pass_eval_context / @pass_context bound; both
                        // forms expect the jinja2 runtime as first arg
                        // and break when invoked as plain callables.
                        // minijinja provides native equivalents, so skip
                        // them and only import user-registered plain
                        // callables.
                        let async_variant = v
                            .getattr("jinja_async_variant")
                            .map(|x| x.is_truthy().unwrap_or(false))
                            .unwrap_or(false);
                        let pass_arg = v
                            .getattr("jinja_pass_arg")
                            .map(|x| !x.is_none())
                            .unwrap_or(false);
                        if async_variant || pass_arg {
                            continue;
                        }
                    }
                    let list = match kind {
                        0 => &mut core.custom_filters,
                        1 => &mut core.custom_globals,
                        _ => &mut core.custom_tests,
                    };
                    list.retain(|(n, _)| n != &name);
                    list.push((name, v.unbind()));
                }
            }
        }
    }
    Ok(autoescape)
}

// ---------------------------------------------------------------- DocxTemplate

/// Send/Ungil wrapper for values that are logically confined to one thread
/// but must cross a GIL-detached section (pyo3 requires Ungil). Sound here:
/// the detached closure still executes on the current thread; a concurrent
/// access to the same DocxTemplate from another thread fails loudly on the
/// RefCell borrow rather than racing.
struct AssertSend<T>(T);
unsafe impl<T> Send for AssertSend<T> {}

/// Class for managing docx files as they were jinja2 templates
#[pyclass(name = "DocxTemplate", unsendable)]
pub struct PyDocxTemplate {
    pub core: RefCell<TplCore>,
    /// cached document facade for __getattr__ delegation
    doc: RefCell<Option<Py<crate::docmodel::PyDocument>>>,
    /// template source path when constructed from a path (docxtpl
    /// template_file); None for bytes / file-like sources
    template_file: Option<String>,
}

#[pymethods]
impl PyDocxTemplate {
    /// Relationship type URI of header parts (docxtpl HEADER_URI).
    #[classattr]
    const HEADER_URI: &'static str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    /// Relationship type URI of footer parts (docxtpl FOOTER_URI).
    #[classattr]
    const FOOTER_URI: &'static str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

    #[new]
    fn new(template_file: &Bound<'_, PyAny>) -> PyResult<Self> {
        let bytes = read_bytes_source(template_file)?;
        let path = template_file.extract::<String>().ok().or_else(|| {
            template_file
                .call_method0("__fspath__")
                .ok()
                .and_then(|f| f.extract::<String>().ok())
        });
        Ok(PyDocxTemplate {
            core: RefCell::new(TplCore::new(bytes)),
            doc: RefCell::new(None),
            template_file: path,
        })
    }

    #[getter]
    fn is_rendered(&self) -> bool {
        self.core.borrow().is_rendered
    }

    #[getter]
    fn is_saved(&self) -> bool {
        self.core.borrow().is_saved
    }

    #[getter]
    fn allow_missing_pics(&self) -> bool {
        self.core.borrow().allow_missing_pics
    }

    #[setter]
    fn set_allow_missing_pics(&self, v: bool) {
        self.core.borrow_mut().allow_missing_pics = v;
    }

    /// Render the template with the given context dict.
    ///
    /// If jinja_env is provided, its `autoescape`, `filters`, `globals` and
    /// `tests` attributes are honored (duck-typed, works with real jinja2
    /// environments). jinja2's own builtins (identity-matched against
    /// jinja2.defaults) are left to minijinja's native implementations;
    /// only user-added/overridden entries are imported.
    #[pyo3(signature = (context, jinja_env=None, autoescape=false))]
    fn render(
        &self,
        context: &Bound<'_, PyAny>,
        jinja_env: Option<&Bound<'_, PyAny>>,
        autoescape: bool,
    ) -> PyResult<()> {
        let mut autoescape = autoescape;
        let mut core = self.core.borrow_mut();
        if let Some(env) = jinja_env {
            if import_jinja_env(&mut core, env)? {
                autoescape = true;
            }
        }
        // Render runs detached from the GIL: the engine re-acquires it on
        // demand (Python::attach) for context/filter callbacks, so other
        // Python threads keep running during a heavy render.
        let ctx_obj = context.clone().unbind();
        let core = AssertSend(core);
        let (result, ctx) = context.py().detach(move || {
            // bind the wrapper as a whole so the closure captures AssertSend
            // (a field access would capture the !Send RefMut directly)
            let mut core = core;
            let result = core.0.render(autoescape, &|core, part| {
                pyo3::Python::attach(|py| {
                    py_to_value(ctx_obj.bind(py), core, part, 0).map_err(|e| e.to_string())
                })
            });
            let ctx = core.0.last_error_context.clone();
            (result, ctx)
        });
        result.map_err(|e| {
            let err = PyErr::new::<crate::pyclasses::TemplateError, _>(e);
            if !ctx.is_empty() {
                Python::attach(|py| {
                    if let Ok(list) = PyList::new(py, &ctx) {
                        let _ = err.value(py).setattr("docx_context", list);
                    }
                });
            }
            err
        })
    }

    /// Unknown attributes are delegated to the document facade
    /// (like docxtpl's __getattr__ delegating to the python-docx Document).
    fn __getattr__(slf: Py<Self>, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        let cached = slf.bind(py).borrow().doc.borrow().as_ref().map(|d| d.clone_ref(py));
        let doc = match cached {
            Some(d) => d,
            None => {
                let d = Py::new(
                    py,
                    crate::docmodel::PyDocument {
                        tpl: slf.clone_ref(py),
                    },
                )?;
                slf.bind(py).borrow().doc.replace(Some(d.clone_ref(py)));
                d
            }
        };
        doc.bind(py).getattr(name).map(|o| o.unbind())
    }

    /// Register a custom jinja filter (python callable).
    fn register_filter(&self, name: &str, callable: Py<PyAny>) {
        let mut core = self.core.borrow_mut();
        core.custom_filters.retain(|(n, _)| n != name);
        core.custom_filters.push((name.to_string(), callable));
        core.env_gen += 1;
    }

    /// Register a custom jinja test (python callable).
    fn register_test(&self, name: &str, callable: Py<PyAny>) {
        let mut core = self.core.borrow_mut();
        core.custom_tests.retain(|(n, _)| n != name);
        core.custom_tests.push((name.to_string(), callable));
        core.env_gen += 1;
    }

    /// Register a custom jinja global function (python callable).
    fn register_function(&self, name: &str, callable: Py<PyAny>) {
        let mut core = self.core.borrow_mut();
        core.custom_functions.retain(|(n, _)| n != name);
        core.custom_functions.push((name.to_string(), callable));
        core.env_gen += 1;
    }

    /// Register a custom jinja global value.
    fn register_global(&self, name: &str, value: Py<PyAny>) {
        let mut core = self.core.borrow_mut();
        core.custom_globals.retain(|(n, _)| n != name);
        core.custom_globals.push((name.to_string(), value));
        core.env_gen += 1;
    }

    /// Set a template loader callable (name -> source or None),
    /// enabling {% include %} and {% import %}.
    fn set_template_loader(&self, loader: Py<PyAny>) {
        let mut core = self.core.borrow_mut();
        core.template_loader = Some(loader);
        core.env_gen += 1;
    }

    /// Install a gettext .mo catalog enabling {% trans %} translations.
    fn install_gettext(&self, mo_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let bytes = read_bytes_source(mo_file)?;
        let catalog = crate::gettext::Catalog::parse(&bytes).map_err(PyValueError::new_err)?;
        let mut core = self.core.borrow_mut();
        core.gettext_catalog = Some(std::sync::Arc::new(catalog));
        core.env_gen += 1;
        Ok(())
    }

    /// Save the rendered (or media-replaced) docx to a path or file-like object.
    pub fn save(&self, py: Python<'_>, filename: &Bound<'_, PyAny>) -> PyResult<()> {
        // zip compression is pure Rust: run it detached from the GIL
        let core = AssertSend(&self.core);
        let bytes = py
            .detach(move || {
                let core = core; // capture AssertSend as a whole
                core.0.borrow_mut().save_bytes()
            })
            .map_err(to_pyerr)?;
        if let Ok(path) = filename.extract::<String>() {
            std::fs::write(&path, &bytes)
                .map_err(|e| PyValueError::new_err(format!("cannot write {}: {}", path, e)))?;
            return Ok(());
        }
        if let Ok(fspath) = filename.call_method0("__fspath__") {
            if let Ok(path) = fspath.extract::<String>() {
                std::fs::write(&path, &bytes)
                    .map_err(|e| PyValueError::new_err(format!("cannot write {}: {}", path, e)))?;
                return Ok(());
            }
        }
        let b = PyBytes::new(py, &bytes);
        filename.call_method1("write", (b,))?;
        Ok(())
    }

    /// Return the current (rendered) document xml.
    fn get_xml(&self, py: Python<'_>) -> PyResult<String> {
        // flush + serialize is pure Rust: run it detached from the GIL
        let core = AssertSend(&self.core);
        py.detach(move || {
            let core = core; // capture AssertSend as a whole
            core.0.borrow_mut().get_xml()
        })
        .map_err(to_pyerr)
    }

    fn write_xml(&self, filename: &str) -> PyResult<()> {
        let xml = self.core.borrow_mut().get_xml().map_err(to_pyerr)?;
        std::fs::write(filename, xml)
            .map_err(|e| PyValueError::new_err(format!("cannot write {}: {}", filename, e)))
    }

    /// Return the docx as bytes (rendered state).
    fn get_docx_bytes(&self, py: Python<'_>) -> PyResult<Vec<u8>> {
        // zip compression is pure Rust: run it detached from the GIL
        let core = AssertSend(&self.core);
        py.detach(move || {
            let core = core; // capture AssertSend as a whole
            core.0.borrow_mut().save_bytes()
        })
        .map_err(to_pyerr)
    }

    /// Undeclared jinja variables used in the template.
    #[pyo3(signature = (jinja_env=None, context=None))]
    fn get_undeclared_template_variables(
        &self,
        py: Python<'_>,
        jinja_env: Option<&Bound<'_, PyAny>>,
        context: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<Py<PySet>> {
        let keys = match context {
            Some(c) => {
                let dict = c.cast::<PyDict>().map_err(|_| {
                    PyValueError::new_err("context must be a dict")
                })?;
                let mut set = std::collections::HashSet::new();
                for (k, _) in dict.iter() {
                    set.insert(k.str()?.to_string_lossy().to_string());
                }
                Some(set)
            }
            None => None,
        };
        let mut core = self.core.borrow_mut();
        // when a jinja_env is given, honor its parse-relevant options
        // (trim_blocks/lstrip_blocks/keep_trailing_newline) and the same
        // preprocessing as the render pipeline ({% trans %} etc.)
        let use_env = match jinja_env {
            Some(env) => {
                import_jinja_env(&mut core, env)?;
                true
            }
            None => false,
        };
        let vars = core.undeclared_variables(keys, use_env).map_err(to_pyerr)?;
        let out = PySet::new(py, vars.iter())?;
        Ok(out.unbind())
    }

    /// Replace one media by another one into the docx (matched by CRC32).
    fn replace_media(&self, src_file: &Bound<'_, PyAny>, dst_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let src = read_bytes_source(src_file)?;
        let dst = read_bytes_source(dst_file)?;
        self.core
            .borrow_mut()
            .crc_to_new_media
            .insert(crc32(&src), dst);
        Ok(())
    }

    /// Replace an embedded object by another one (matched by CRC32).
    fn replace_embedded(&self, src_file: &Bound<'_, PyAny>, dst_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let src = read_bytes_source(src_file)?;
        let dst = read_bytes_source(dst_file)?;
        self.core
            .borrow_mut()
            .crc_to_new_embedded
            .insert(crc32(&src), dst);
        Ok(())
    }

    /// Replace a picture in the docx by its original filename/title/description.
    fn replace_pic(&self, embedded_file: String, dst_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let dst = read_bytes_source(dst_file)?;
        self.core
            .borrow_mut()
            .pics_to_replace
            .insert(embedded_file, dst);
        Ok(())
    }

    /// Replace one file in the docx zip by its zip name.
    fn replace_zipname(&self, zipname: String, dst_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let dst = read_bytes_source(dst_file)?;
        self.core
            .borrow_mut()
            .zipname_to_replace
            .insert(zipname, dst);
        Ok(())
    }

    fn reset_replacements(&self) {
        self.core.borrow_mut().reset_replacements();
    }

    /// Create a new Subdoc from a docx file path (or empty if omitted).
    /// keep_sections=True preserves the subdoc's section properties (page
    /// size/orientation/margins, header/footer references) by making the
    /// subdoc content its own section; the default False matches docxtpl.
    #[pyo3(signature = (docpath=None, keep_sections=false))]
    fn new_subdoc(
        slf: Py<Self>,
        py: Python<'_>,
        docpath: Option<&Bound<'_, PyAny>>,
        keep_sections: bool,
    ) -> PyResult<Py<PySubdoc>> {
        slf.bind(py)
            .borrow()
            .core
            .borrow_mut()
            .init_docx(false)
            .map_err(to_pyerr)?;
        let bytes = match docpath {
            Some(p) => Some(read_bytes_source(p)?),
            None => None,
        };
        Py::new(py, PySubdoc { bytes, tpl: slf, blocks: std::cell::RefCell::new(Vec::new()), keep_sections })
    }

    /// Create an external hyperlink relationship, returns the rId.
    fn build_url_id(&self, url: &str) -> PyResult<String> {
        self.core.borrow_mut().build_url_id(url).map_err(to_pyerr)
    }

    /// A python-docx-inspired facade over the current document.
    fn get_docx(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<crate::docmodel::PyDocument>> {
        Py::new(py, crate::docmodel::PyDocument { tpl: slf })
    }

    #[getter]
    fn paragraphs(slf: Py<Self>, py: Python<'_>) -> Vec<crate::docmodel::PyParagraph> {
        crate::docmodel::PyDocument { tpl: slf }.paragraphs(py)
    }

    #[getter]
    fn tables(slf: Py<Self>, py: Python<'_>) -> Vec<crate::docmodel::PyTable> {
        crate::docmodel::PyDocument { tpl: slf }.tables(py)
    }

    #[getter]
    fn sections(slf: Py<Self>, py: Python<'_>) -> Vec<crate::docmodel::PySection> {
        crate::docmodel::PyDocument { tpl: slf }.sections(py)
    }

    #[getter]
    fn styles(slf: Py<Self>, py: Python<'_>) -> crate::docmodel::PyStyles {
        crate::docmodel::PyDocument { tpl: slf }.styles(py)
    }

    #[getter]
    fn settings(slf: Py<Self>, py: Python<'_>) -> crate::docmodel::PySettings {
        crate::docmodel::PyDocument { tpl: slf }.settings(py)
    }

    #[getter]
    fn comments(slf: Py<Self>, py: Python<'_>) -> crate::doccomments::PyComments {
        crate::docmodel::PyDocument { tpl: slf }.comments(py)
    }

    #[getter]
    fn core_properties(slf: Py<Self>, py: Python<'_>) -> crate::docmodel::PyCoreProperties {
        crate::docmodel::PyDocument { tpl: slf }.core_properties(py)
    }

    #[getter]
    fn inline_shapes(slf: Py<Self>, py: Python<'_>) -> Vec<crate::docmodel::PyInlineShape> {
        crate::docmodel::PyDocument { tpl: slf }.inline_shapes(py)
    }

    /// Map of pictures found while replacing: filename -> (target_ref, partname)
    fn get_pic_map(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let dict = PyDict::new(py);
        for (k, (tref, tpart)) in &self.core.borrow().pic_map {
            dict.set_item(k, (tref, tpart))?;
        }
        Ok(dict.unbind())
    }

    /// CRC32 of a file (path, bytes or file-like), as used by replace_pic
    /// to locate an embedded image (docxtpl get_file_crc).
    #[staticmethod]
    fn get_file_crc(file_obj: &Bound<'_, PyAny>) -> PyResult<u32> {
        let bytes = read_bytes_source(file_obj)?;
        Ok(crc32(&bytes))
    }

    /// Iterate (rel_id, element) of the header/footer parts of the document,
    /// filtered by relationship type URI (docxtpl get_headers_footers).
    /// `element` is a live XmlElement proxy of the part root.
    fn get_headers_footers(
        slf: Py<Self>,
        py: Python<'_>,
        uri: &str,
    ) -> PyResult<Vec<(String, crate::pyxml::PyXmlElement)>> {
        let tpl = slf.bind(py).borrow();
        let mut core = tpl.core.borrow_mut();
        core.init_docx(false).map_err(to_pyerr)?;
        let pkg = core
            .package
            .as_ref()
            .ok_or_else(|| to_pyerr("package not loaded".into()))?;
        let rels = pkg.rels(crate::template::DOCUMENT_PART);
        let mut out = Vec::new();
        for rel in rels.by_type(uri) {
            let part =
                crate::package::resolve_target(crate::template::DOCUMENT_PART, &rel.target);
            if !pkg.contains(&part) {
                continue;
            }
            out.push((
                rel.id.clone(),
                crate::pyxml::PyXmlElement {
                    tpl: slf.clone_ref(py),
                    part,
                    path: Vec::new(),
                },
            ));
        }
        Ok(out)
    }

    // ---------------- docxtpl pipeline methods (string-based) ----------------
    //
    // Unlike docxtpl (which passes lxml trees / Part objects around), these
    // take and return plain XML strings (and part names as zip paths),
    // matching the existing get_xml/write_xml style of this library.

    /// Template source path given at construction time (docxtpl
    /// template_file); None when constructed from bytes or a file-like.
    #[getter]
    fn template_file(&self) -> Option<String> {
        self.template_file.clone()
    }

    /// The document facade; equivalent to get_docx() (docxtpl `docx`,
    /// except that docxtpl's is None before the first init_docx()).
    #[getter(docx)]
    fn docx_prop(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<crate::docmodel::PyDocument>> {
        Py::new(py, crate::docmodel::PyDocument { tpl: slf })
    }

    /// Map of pictures found while replacing: filename -> (target_ref,
    /// partname). Equivalent to get_pic_map() (docxtpl `pic_map`).
    #[getter(pic_map)]
    fn pic_map_prop(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        self.get_pic_map(py)
    }

    /// Name of the part currently being rendered (docxtpl
    /// current_rendering_part); None outside a render.
    #[getter]
    fn current_rendering_part(&self) -> Option<String> {
        crate::pybridge::current_rendering_part()
    }

    #[setter]
    fn set_is_rendered(&self, v: bool) {
        self.core.borrow_mut().is_rendered = v;
    }

    #[setter]
    fn set_is_saved(&self, v: bool) {
        self.core.borrow_mut().is_saved = v;
    }

    /// Media replacement map {crc32: new_bytes} (docxtpl crc_to_new_media).
    /// The getter returns a snapshot copy: mutating it in place has no
    /// effect, reassign the attribute instead (unlike docxtpl's live dict).
    #[getter]
    fn crc_to_new_media(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let dict = PyDict::new(py);
        for (k, v) in &self.core.borrow().crc_to_new_media {
            dict.set_item(k, PyBytes::new(py, v))?;
        }
        Ok(dict.unbind())
    }

    #[setter]
    fn set_crc_to_new_media(&self, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = value.cast::<PyDict>()?;
        let mut map = HashMap::new();
        for (k, v) in dict.iter() {
            map.insert(k.extract::<u32>()?, v.extract::<Vec<u8>>()?);
        }
        self.core.borrow_mut().crc_to_new_media = map;
        Ok(())
    }

    /// Embedded-object replacement map {crc32: new_bytes} (docxtpl
    /// crc_to_new_embedded). Snapshot semantics: see crc_to_new_media.
    #[getter]
    fn crc_to_new_embedded(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let dict = PyDict::new(py);
        for (k, v) in &self.core.borrow().crc_to_new_embedded {
            dict.set_item(k, PyBytes::new(py, v))?;
        }
        Ok(dict.unbind())
    }

    #[setter]
    fn set_crc_to_new_embedded(&self, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = value.cast::<PyDict>()?;
        let mut map = HashMap::new();
        for (k, v) in dict.iter() {
            map.insert(k.extract::<u32>()?, v.extract::<Vec<u8>>()?);
        }
        self.core.borrow_mut().crc_to_new_embedded = map;
        Ok(())
    }

    /// Zip-entry replacement map {zipname: new_bytes} (docxtpl
    /// zipname_to_replace). Snapshot semantics: see crc_to_new_media.
    #[getter]
    fn zipname_to_replace(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let dict = PyDict::new(py);
        for (k, v) in &self.core.borrow().zipname_to_replace {
            dict.set_item(k, PyBytes::new(py, v))?;
        }
        Ok(dict.unbind())
    }

    #[setter]
    fn set_zipname_to_replace(&self, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = value.cast::<PyDict>()?;
        let mut map = HashMap::new();
        for (k, v) in dict.iter() {
            map.insert(k.extract::<String>()?, v.extract::<Vec<u8>>()?);
        }
        self.core.borrow_mut().zipname_to_replace = map;
        Ok(())
    }

    /// Picture replacement map {name/title/descr: new_bytes} (docxtpl
    /// pics_to_replace). Snapshot semantics: see crc_to_new_media.
    #[getter]
    fn pics_to_replace(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let dict = PyDict::new(py);
        for (k, v) in &self.core.borrow().pics_to_replace {
            dict.set_item(k, PyBytes::new(py, v))?;
        }
        Ok(dict.unbind())
    }

    #[setter]
    fn set_pics_to_replace(&self, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = value.cast::<PyDict>()?;
        let mut map = HashMap::new();
        for (k, v) in dict.iter() {
            map.insert(k.extract::<String>()?, v.extract::<Vec<u8>>()?);
        }
        self.core.borrow_mut().pics_to_replace = map;
        Ok(())
    }

    /// Run the internal patch pipeline on a raw xml string (docxtpl
    /// patch_xml): text-node entity decoding + tag merging/cleaning,
    /// returning the patched xml string.
    fn patch_xml(&self, src_xml: &str) -> String {
        crate::patch::patch_xml(&crate::patch::decode_text_entities(src_xml)).into_owned()
    }

    /// Expand newlines/tabs/page breaks inside w:t into w:br/w:tab/paragraph
    /// splits (docxtpl resolve_listing), returning the xml string.
    fn resolve_listing(&self, xml: &str) -> String {
        crate::patch::resolve_listing(xml).into_owned()
    }

    /// Fix table grids whose rows have more/fewer cells than w:gridCol
    /// declarations (docxtpl fix_tables), with the same three-level fallback
    /// (strict DOM parse -> recovery parse -> regex). Takes and returns an
    /// xml string instead of an lxml tree.
    fn fix_tables(&self, xml: &str) -> PyResult<String> {
        crate::template::fix_tables_only(xml).map_err(to_pyerr)
    }

    /// Renumber wp:docPr ids (from the internal docx_ids_index) and
    /// pic:cNvPr ids so they are unique (docxtpl fix_docpr_ids); no table
    /// fixing is done. Takes and returns an xml string instead of mutating
    /// an lxml tree.
    fn fix_docpr_ids(&self, xml: &str) -> String {
        let mut core = self.core.borrow_mut();
        let mut cnvpr_next = 1u32;
        crate::template::fix_docpr_cnvpr_ids(xml, &mut core.docx_ids_index, Some(&mut cnvpr_next))
    }

    /// docxtpl xml_to_string (etree.tostring). Here: a str is returned
    /// unchanged (no lxml normalization exists), bytes are decoded using
    /// `encoding` (any Python codec) and returned as str.
    #[pyo3(signature = (xml, encoding="utf-8"))]
    fn xml_to_string(&self, xml: &Bound<'_, PyAny>, encoding: &str) -> PyResult<String> {
        if let Ok(s) = xml.extract::<String>() {
            return Ok(s);
        }
        if let Ok(b) = xml.cast::<PyBytes>() {
            return Ok(b.call_method1("decode", (encoding,))?.extract::<String>()?);
        }
        Err(PyValueError::new_err("expected str or bytes"))
    }

    /// Render one already-patched xml string with the given context
    /// (docxtpl render_xml_part). `part` is the part's zip path (docxtpl
    /// passes a Part object) and is used for deferred InlineImage/Subdoc
    /// materialization and current_rendering_part. jinja_env is honored the
    /// same way as in render(). Returns the rendered xml string.
    #[pyo3(signature = (src_xml, part, context, jinja_env=None))]
    fn render_xml_part(
        &self,
        src_xml: &str,
        part: &str,
        context: &Bound<'_, PyAny>,
        jinja_env: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<String> {
        let mut autoescape = false;
        let mut core = self.core.borrow_mut();
        if let Some(env) = jinja_env {
            if import_jinja_env(&mut core, env)? {
                autoescape = true;
            }
        }
        core.init_docx(false).map_err(to_pyerr)?;
        let prev = crate::pybridge::set_current_render(&mut *core, part);
        let run = |core: &mut TplCore| -> Result<String, String> {
            let ctx = py_to_value(context, core, part, 0).map_err(|e| e.to_string())?;
            let rendered = crate::template::render_xml_str(src_xml, ctx, autoescape, core)?;
            let rendered = crate::patch::resolve_listing(&rendered).into_owned();
            core.materialize_deferred(part, rendered)
        };
        let result = run(&mut core);
        crate::pybridge::restore_current_render(prev);
        result.map_err(to_pyerr)
    }

    /// Render the jinja placeholders in docProps/core.xml with the given
    /// context, updating the part in the package (docxtpl
    /// render_properties). jinja_env is honored the same way as in
    /// render() (autoescape is always off here, like docxtpl).
    #[pyo3(signature = (context, jinja_env=None))]
    fn render_properties(
        &self,
        context: &Bound<'_, PyAny>,
        jinja_env: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let mut core = self.core.borrow_mut();
        if let Some(env) = jinja_env {
            import_jinja_env(&mut core, env)?;
        }
        core.init_docx(false).map_err(to_pyerr)?;
        core.render_properties(&|core, part| {
            py_to_value(context, core, part, 0).map_err(|e| e.to_string())
        })
        .map_err(to_pyerr)
    }

    /// Render the footnotes part(s) with the given context, updating them in
    /// the package (docxtpl render_footnotes). No-op when the document has
    /// no footnotes part.
    #[pyo3(signature = (context, jinja_env=None))]
    fn render_footnotes(
        &self,
        context: &Bound<'_, PyAny>,
        jinja_env: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let mut autoescape = false;
        let mut core = self.core.borrow_mut();
        if let Some(env) = jinja_env {
            if import_jinja_env(&mut core, env)? {
                autoescape = true;
            }
        }
        core.init_docx(false).map_err(to_pyerr)?;
        core.render_footnotes(autoescape, &|core, part| {
            py_to_value(context, core, part, 0).map_err(|e| e.to_string())
        })
        .map_err(to_pyerr)
    }

    /// Render the document body with the given context and return the
    /// rendered xml string (docxtpl build_xml). The package is not
    /// modified; table/docPr fixing is not applied (see map_tree()).
    #[pyo3(signature = (context, jinja_env=None))]
    fn build_xml(
        &self,
        context: &Bound<'_, PyAny>,
        jinja_env: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<String> {
        let mut autoescape = false;
        let mut core = self.core.borrow_mut();
        if let Some(env) = jinja_env {
            if import_jinja_env(&mut core, env)? {
                autoescape = true;
            }
        }
        let src = core.get_xml().map_err(to_pyerr)?;
        core.render_part(crate::template::DOCUMENT_PART, &src, autoescape, &|core, part| {
            py_to_value(context, core, part, 0).map_err(|e| e.to_string())
        })
        .map_err(to_pyerr)
    }

    /// Apply fix_tables + fix_docpr_ids to a rendered body xml string
    /// (docxtpl map_tree replaces the body with the fixed tree). Here the
    /// fixed xml string is returned instead; the package is not modified.
    fn map_tree(&self, xml: &str) -> PyResult<String> {
        let mut core = self.core.borrow_mut();
        crate::template::fix_tables_and_docpr(xml, &mut core.docx_ids_index).map_err(to_pyerr)
    }

    /// Current xml of a package part given by its zip path, e.g.
    /// "word/document.xml" (docxtpl get_part_xml, which takes a Part
    /// object and returns the lxml-serialized string; here the stored xml
    /// string is returned as-is).
    fn get_part_xml(&self, part: &str) -> PyResult<String> {
        self.core.borrow_mut().get_part_xml(part).map_err(to_pyerr)
    }

    /// Encoding declared in an xml declaration, "utf-8" when absent
    /// (docxtpl get_headers_footers_encoding). Accepts str or bytes.
    fn get_headers_footers_encoding(&self, xml: &Bound<'_, PyAny>) -> PyResult<String> {
        if let Ok(s) = xml.extract::<String>() {
            return Ok(crate::template::headers_footers_encoding(&s));
        }
        if let Ok(b) = xml.cast::<PyBytes>() {
            return Ok(crate::template::headers_footers_encoding(
                &String::from_utf8_lossy(b.as_bytes()),
            ));
        }
        Err(PyValueError::new_err("expected str or bytes"))
    }

    /// Render the header/footer parts matching `uri` (HEADER_URI /
    /// FOOTER_URI) with the given context (docxtpl build_headers_footers_xml).
    /// Returns a dict {relKey: rendered_xml} — docxtpl yields
    /// (relKey, bytes) pairs encoded with the part's declared encoding.
    #[pyo3(signature = (context, uri, jinja_env=None))]
    fn build_headers_footers_xml(
        &self,
        py: Python<'_>,
        context: &Bound<'_, PyAny>,
        uri: &str,
        jinja_env: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<Py<PyDict>> {
        let mut autoescape = false;
        let mut core = self.core.borrow_mut();
        if let Some(env) = jinja_env {
            if import_jinja_env(&mut core, env)? {
                autoescape = true;
            }
        }
        core.init_docx(false).map_err(to_pyerr)?;
        core.flush_parts().map_err(to_pyerr)?;
        let pairs: Vec<(String, String)> = match core.package.as_ref() {
            Some(pkg) => pkg
                .rels(crate::template::DOCUMENT_PART)
                .by_type(uri)
                .filter(|r| !r.is_external)
                .map(|r| {
                    (
                        r.id.clone(),
                        crate::package::resolve_target(crate::template::DOCUMENT_PART, &r.target),
                    )
                })
                .collect(),
            None => Vec::new(),
        };
        let srcs: Vec<(String, String, String)> = pairs
            .into_iter()
            .filter_map(|(rid, part)| {
                core.package
                    .as_ref()
                    .and_then(|p| p.get_string(&part))
                    .map(|s| (rid, part, s))
            })
            .collect();
        let out = PyDict::new(py);
        for (rid, part, src) in srcs {
            let rendered = core
                .render_part(&part, &src, autoescape, &|core, part| {
                    py_to_value(context, core, part, 0).map_err(|e| e.to_string())
                })
                .map_err(to_pyerr)?;
            out.set_item(rid, rendered)?;
        }
        Ok(out.unbind())
    }

    /// Apply the header/footer fixups to a rendered header/footer xml
    /// string: fix_tables, plus docPr id renumbering when relKey points to
    /// a header (docxtpl map_headers_footers_xml). Here the fixed xml
    /// string is returned instead of replacing the part in the package.
    fn map_headers_footers_xml(&self, rel_key: &str, xml: &str) -> PyResult<String> {
        let mut core = self.core.borrow_mut();
        core.init_docx(false).map_err(to_pyerr)?;
        let is_header = core
            .package
            .as_ref()
            .and_then(|p| {
                p.rels(crate::template::DOCUMENT_PART)
                    .get(rel_key)
                    .map(|r| r.rel_type == crate::package::rel_type::HEADER)
            })
            .unwrap_or(false);
        let fixed = crate::template::fix_tables_only(xml).map_err(to_pyerr)?;
        if is_header {
            Ok(crate::template::fix_docpr_cnvpr_ids(
                &fixed,
                &mut core.docx_ids_index,
                None,
            ))
        } else {
            Ok(fixed)
        }
    }

    /// (Re)initialize the internal package from the template bytes
    /// (docxtpl init_docx). With reload=False the package is only loaded if
    /// it has not been loaded yet (or was reloaded since the last render).
    #[pyo3(signature = (reload=true))]
    fn init_docx(&self, reload: bool) -> PyResult<()> {
        self.core.borrow_mut().init_docx(reload).map_err(to_pyerr)
    }

    /// Reset the per-render state: reloads the package, clears pic_map and
    /// deferred values, resets docx_ids_index and is_saved (docxtpl
    /// render_init).
    fn render_init(&self) -> PyResult<()> {
        self.core.borrow_mut().render_init().map_err(to_pyerr)
    }

    /// Apply the pending pics_to_replace replacements to the package and
    /// populate pic_map (docxtpl pre_processing). Raises for pictures not
    /// found in the template unless allow_missing_pics is set.
    fn pre_processing(&self) -> PyResult<()> {
        let mut core = self.core.borrow_mut();
        core.init_docx(false).map_err(to_pyerr)?;
        core.pre_processing().map_err(to_pyerr)
    }

    /// Apply the pending zip-level replacements (crc_to_new_media /
    /// crc_to_new_embedded / zipname_to_replace) to a docx file on disk,
    /// rewriting it in place (docxtpl post_processing). The live package is
    /// not touched. No-op when all replacement dicts are empty.
    fn post_processing(&self, docx_file: &str) -> PyResult<()> {
        let core = self.core.borrow();
        if core.crc_to_new_media.is_empty()
            && core.crc_to_new_embedded.is_empty()
            && core.zipname_to_replace.is_empty()
        {
            return Ok(());
        }
        let data = std::fs::read(docx_file)
            .map_err(|e| PyValueError::new_err(format!("cannot read {}: {}", docx_file, e)))?;
        let out = core.post_processing_bytes(&data).map_err(to_pyerr)?;
        std::fs::write(docx_file, &out)
            .map_err(|e| PyValueError::new_err(format!("cannot write {}: {}", docx_file, e)))
    }
}

// ---------------------------------------------------------------- RichText

fn parse_text_props(kwargs: Option<&Bound<'_, PyDict>>) -> PyResult<TextProps> {
    let mut p = TextProps::default();
    let Some(kw) = kwargs else { return Ok(p) };

    fn get_str(kw: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
        match kw.get_item(key)? {
            Some(v) if !v.is_none() => Ok(Some(v.extract::<String>()?)),
            _ => Ok(None),
        }
    }
    fn get_bool(kw: &Bound<'_, PyDict>, key: &str) -> PyResult<bool> {
        match kw.get_item(key)? {
            Some(v) if !v.is_none() => Ok(v.is_truthy()?),
            _ => Ok(false),
        }
    }

    p.style = get_str(kw, "style")?;
    p.color = get_str(kw, "color")?;
    p.highlight = get_str(kw, "highlight")?;
    if let Some(v) = kw.get_item("size")? {
        if !v.is_none() {
            p.size = Some(v.extract::<u32>()?);
        }
    }
    p.subscript = get_bool(kw, "subscript")?;
    p.superscript = get_bool(kw, "superscript")?;
    p.bold = get_bool(kw, "bold")?;
    p.italic = get_bool(kw, "italic")?;
    if let Some(v) = kw.get_item("underline")? {
        if !v.is_none() {
            // True -> "single"; a string value is used as-is
            if let Ok(s) = v.extract::<String>() {
                p.underline = Some(s);
            } else if v.is_truthy()? {
                p.underline = Some("single".to_string());
            }
        }
    }
    p.strike = get_bool(kw, "strike")?;
    p.font = get_str(kw, "font")?;
    if let Some(v) = kw.get_item("url_id")? {
        if !v.is_none() {
            // url_id may be a plain string rId or an object with rId-ish str
            p.url_id = Some(v.str()?.to_string_lossy().to_string());
        }
    }
    p.rtl = get_bool(kw, "rtl")?;
    p.lang = get_str(kw, "lang")?;
    Ok(p)
}

fn any_to_text(text: &Bound<'_, PyAny>) -> PyResult<String> {
    if let Ok(s) = text.extract::<String>() {
        return Ok(s);
    }
    if let Ok(b) = text.cast::<PyBytes>() {
        return Ok(String::from_utf8_lossy(b.as_bytes()).to_string());
    }
    Ok(text.str()?.to_string_lossy().to_string())
}

/// Generate rich text runs usable inside an existing paragraph via {{ r }}
#[pyclass(name = "RichText", unsendable)]
pub struct PyRichText {
    pub xml: RefCell<String>,
}

#[pymethods]
impl PyRichText {
    #[new]
    #[pyo3(signature = (text=None, **text_prop))]
    fn new(text: Option<&Bound<'_, PyAny>>, text_prop: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let rt = PyRichText {
            xml: RefCell::new(String::new()),
        };
        if let Some(t) = text {
            if py_truthy(t) {
                rt.add(t, text_prop)?;
            }
        }
        Ok(rt)
    }

    #[pyo3(signature = (text, **text_prop))]
    fn add(&self, text: &Bound<'_, PyAny>, text_prop: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        // If a RichText is added
        if let Ok(other) = text.cast::<PyRichText>() {
            let other_xml = other.borrow().xml.borrow().clone();
            self.xml.borrow_mut().push_str(&other_xml);
            return Ok(());
        }
        let props = parse_text_props(text_prop)?;
        let s = any_to_text(text)?;
        self.xml
            .borrow_mut()
            .push_str(&richtext::richtext_run(&s, &props));
        Ok(())
    }

    #[getter]
    fn xml(&self) -> String {
        self.xml.borrow().clone()
    }

    fn __str__(&self) -> String {
        self.xml.borrow().clone()
    }

    fn __html__(&self) -> String {
        self.xml.borrow().clone()
    }
}

/// Generate rich text paragraphs usable outside existing paragraphs via {{p rp }}
#[pyclass(name = "RichTextParagraph", unsendable)]
pub struct PyRichTextParagraph {
    pub xml: RefCell<String>,
}

#[pymethods]
impl PyRichTextParagraph {
    #[new]
    #[pyo3(signature = (text=None, **text_prop))]
    fn new(text: Option<&Bound<'_, PyAny>>, text_prop: Option<&Bound<'_, PyDict>>) -> PyResult<Self> {
        let rp = PyRichTextParagraph {
            xml: RefCell::new(String::new()),
        };
        if let Some(t) = text {
            if py_truthy(t) {
                rp.add(t, text_prop)?;
            }
        }
        Ok(rp)
    }

    #[pyo3(signature = (text, **text_prop))]
    fn add(&self, text: &Bound<'_, PyAny>, text_prop: Option<&Bound<'_, PyDict>>) -> PyResult<()> {
        let runs_xml = if let Ok(rt) = text.cast::<PyRichText>() {
            rt.borrow().xml.borrow().clone()
        } else {
            let s = any_to_text(text)?;
            richtext::richtext_run(&s, &TextProps::default())
        };
        let parastyle = match text_prop.and_then(|kw| kw.get_item("parastyle").ok().flatten()) {
            Some(v) if !v.is_none() => Some(v.extract::<String>()?),
            _ => None,
        };
        self.xml
            .borrow_mut()
            .push_str(&richtext::richtext_paragraph(&runs_xml, parastyle.as_deref()));
        Ok(())
    }

    #[getter]
    fn xml(&self) -> String {
        self.xml.borrow().clone()
    }

    fn __str__(&self) -> String {
        self.xml.borrow().clone()
    }

    fn __html__(&self) -> String {
        self.xml.borrow().clone()
    }
}

// ---------------------------------------------------------------- Listing

/// Keep \n, \a, \t, \f in text while keeping current template styling
#[pyclass(name = "Listing", unsendable)]
pub struct PyListing {
    pub xml: RefCell<String>,
}

#[pymethods]
impl PyListing {
    #[new]
    fn new(text: &Bound<'_, PyAny>) -> PyResult<Self> {
        let s = any_to_text(text)?;
        Ok(PyListing {
            xml: RefCell::new(richtext::listing_xml(&s)),
        })
    }

    #[getter]
    fn xml(&self) -> String {
        self.xml.borrow().clone()
    }

    fn __str__(&self) -> String {
        self.xml.borrow().clone()
    }

    fn __html__(&self) -> String {
        self.xml.borrow().clone()
    }
}

// ---------------------------------------------------------------- Length

/// Length value (in EMU) like docx.shared.Length
#[pyclass(name = "Length", skip_from_py_object)]
#[derive(Clone)]
pub struct PyLength {
    pub emu: i64,
}

#[pymethods]
impl PyLength {
    #[new]
    fn new(emu: i64) -> Self {
        PyLength { emu }
    }

    #[getter]
    fn emu(&self) -> i64 {
        self.emu
    }
    #[getter]
    fn inches(&self) -> f64 {
        self.emu as f64 / len::EMU_PER_INCH as f64
    }
    #[getter]
    fn cm(&self) -> f64 {
        self.emu as f64 / len::EMU_PER_CM as f64
    }
    #[getter]
    fn mm(&self) -> f64 {
        self.emu as f64 / len::EMU_PER_MM as f64
    }
    #[getter]
    fn pt(&self) -> f64 {
        self.emu as f64 / len::EMU_PER_PT as f64
    }
    #[getter]
    fn twips(&self) -> f64 {
        self.emu as f64 / len::EMU_PER_TWIP as f64
    }
    fn __int__(&self) -> i64 {
        self.emu
    }
    fn __repr__(&self) -> String {
        format!("Length({})", self.emu)
    }
}

#[pyfunction]
#[pyo3(name = "Emu")]
pub fn emu(v: i64) -> PyLength {
    PyLength { emu: v }
}

#[pyfunction]
#[pyo3(name = "Inches")]
pub fn inches(v: f64) -> PyLength {
    PyLength { emu: len::inches(v) }
}

#[pyfunction]
#[pyo3(name = "Cm")]
pub fn cm(v: f64) -> PyLength {
    PyLength { emu: len::cm(v) }
}

#[pyfunction]
#[pyo3(name = "Mm")]
pub fn mm(v: f64) -> PyLength {
    PyLength { emu: len::mm(v) }
}

#[pyfunction]
#[pyo3(name = "Pt")]
pub fn pt(v: f64) -> PyLength {
    PyLength { emu: len::pt(v) }
}

#[pyfunction]
#[pyo3(name = "Twips")]
pub fn twips(v: f64) -> PyLength {
    PyLength { emu: len::twips(v) }
}

/// public wrapper for docmodel
pub fn extract_length_pub(obj: &Bound<'_, PyAny>) -> PyResult<Option<i64>> {
    extract_length(obj)
}

fn extract_length(obj: &Bound<'_, PyAny>) -> PyResult<Option<i64>> {
    if obj.is_none() {
        return Ok(None);
    }
    if let Ok(l) = obj.cast::<PyLength>() {
        return Ok(Some(l.borrow().emu));
    }
    if let Ok(i) = obj.extract::<i64>() {
        return Ok(Some(i));
    }
    if let Ok(f) = obj.extract::<f64>() {
        return Ok(Some(f as i64));
    }
    Err(PyValueError::new_err(
        "width/height must be None, an int (EMU) or a Length",
    ))
}

// ---------------------------------------------------------------- InlineImage
/// Generate an inline image from a template variable
#[pyclass(name = "InlineImage", unsendable)]
pub struct PyInlineImage {
    /// Arc-shared so registering the Deferred per render is O(1)
    pub blob: std::sync::Arc<[u8]>,
    pub filename: Option<String>,
    pub width: Option<i64>,
    pub height: Option<i64>,
    pub anchor: Option<String>,
    pub title: Option<String>,
    pub descr: Option<String>,
    pub tpl: Py<PyDocxTemplate>,
}

#[pymethods]
impl PyInlineImage {
    #[new]
    #[pyo3(signature = (tpl, image_descriptor, width=None, height=None, anchor=None, title=None, descr=None))]
    fn new(
        tpl: Py<PyDocxTemplate>,
        image_descriptor: &Bound<'_, PyAny>,
        width: Option<&Bound<'_, PyAny>>,
        height: Option<&Bound<'_, PyAny>>,
        anchor: Option<String>,
        title: Option<String>,
        descr: Option<String>,
    ) -> PyResult<Self> {
        let filename = image_descriptor
            .extract::<String>()
            .ok()
            .and_then(|p| {
                std::path::Path::new(&p)
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
            });
        let blob = read_bytes_source(image_descriptor)?;
        Ok(PyInlineImage {
            blob: blob.into(),
            filename,
            width: width.map(|w| extract_length(w)).transpose()?.flatten(),
            height: height.map(|h| extract_length(h)).transpose()?.flatten(),
            anchor,
            title,
            descr,
            tpl,
        })
    }
}

// ---------------------------------------------------------------- Subdoc

/// Subdocument to insert into the master document via {{p subdoc }}.
///
/// With a docpath it merges the given document; without it, content can be
/// built programmatically (add_paragraph/add_heading/add_picture/add_table).
#[pyclass(name = "Subdoc", unsendable)]
pub struct PySubdoc {
    pub bytes: Option<Vec<u8>>,
    pub tpl: Py<PyDocxTemplate>,
    pub blocks: RefCell<Vec<crate::subdocbuilder::Block>>,
    /// preserve the subdoc's section properties (page setup, headers/footers)
    pub keep_sections: bool,
}

fn edit_block<R>(
    subdoc: &Py<PySubdoc>,
    py: Python<'_>,
    index: usize,
    f: impl FnOnce(&mut crate::subdocbuilder::Block) -> R,
) -> PyResult<R> {
    let sd = subdoc.bind(py).borrow();
    let mut blocks = sd.blocks.borrow_mut();
    let block = blocks
        .get_mut(index)
        .ok_or_else(|| PyValueError::new_err("invalid block index"))?;
    Ok(f(block))
}

/// Build a read-only PyDocument facade over the subdoc's current content
/// (the delegation target of docxtpl's Subdoc.__getattr__ and its `docx`
/// attribute). A subdoc created from a file/bytes is wrapped directly; a
/// bound (builder) subdoc serializes its accumulated blocks into a minimal
/// document (pictures are embedded into that document's package).
fn subdoc_docx(slf: &Py<PySubdoc>, py: Python<'_>) -> PyResult<Py<crate::docmodel::PyDocument>> {
    let (bytes, blocks) = {
        let sd = slf.bind(py).borrow();
        let blocks = sd.blocks.borrow().clone();
        (sd.bytes.clone(), blocks)
    };
    let core = match bytes {
        // file-based subdoc: facade over the source document
        Some(b) => TplCore::new(b),
        None => {
            let mut core = TplCore::new(crate::package::minimal_docx(""));
            if !blocks.is_empty() {
                core.init_docx(false).map_err(to_pyerr)?;
                let body = crate::subdocbuilder::serialize_blocks(
                    &mut core,
                    crate::template::DOCUMENT_PART,
                    &blocks,
                )
                .map_err(to_pyerr)?;
                let xml = format!(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}<w:sectPr/></w:body></w:document>",
                    body
                );
                core.package
                    .as_mut()
                    .unwrap()
                    .set(crate::template::DOCUMENT_PART, xml.into_bytes());
                core.invalidate_doc();
            }
            core
        }
    };
    let tpl = Py::new(
        py,
        PyDocxTemplate {
            core: RefCell::new(core),
            doc: RefCell::new(None),
            template_file: None,
        },
    )?;
    Py::new(py, crate::docmodel::PyDocument { tpl })
}

#[pymethods]
impl PySubdoc {
    /// A read-only document facade over the subdoc's current content
    /// (docxtpl Subdoc.docx). Rebuilt on each access, so programmatically
    /// added content is reflected.
    #[getter]
    fn docx(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<crate::docmodel::PyDocument>> {
        subdoc_docx(&slf, py)
    }

    /// Unknown attributes are delegated to a document facade over the
    /// subdoc's current content (docxtpl Subdoc.__getattr__ delegating to
    /// the inner python-docx Document).
    fn __getattr__(slf: Py<Self>, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        let doc = subdoc_docx(&slf, py)?;
        doc.bind(py).getattr(name).map(|o| o.unbind())
    }

    #[new]
    #[pyo3(signature = (tpl, docpath=None, keep_sections=false))]
    fn new(
        tpl: Py<PyDocxTemplate>,
        docpath: Option<&Bound<'_, PyAny>>,
        keep_sections: bool,
    ) -> PyResult<Self> {
        let bytes = match docpath {
            Some(p) => Some(read_bytes_source(p)?),
            None => None,
        };
        Ok(PySubdoc {
            bytes,
            tpl,
            blocks: RefCell::new(Vec::new()),
            keep_sections,
        })
    }

    /// Add a paragraph, returns a Paragraph proxy.
    #[pyo3(signature = (text="", style=None))]
    fn add_paragraph(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        style: Option<String>,
    ) -> PyResult<Py<PySubParagraph>> {
        let runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![crate::subdocbuilder::SubRun {
                text: text.to_string(),
                props: TextProps::default(),
            }]
        };
        let index = {
            let sd = slf.bind(py).borrow();
            let mut blocks = sd.blocks.borrow_mut();
            blocks.push(crate::subdocbuilder::Block::Paragraph { style, runs });
            blocks.len() - 1
        };
        Py::new(py, PySubParagraph { subdoc: slf, index })
    }

    /// Add a heading paragraph of the given level.
    #[pyo3(signature = (text="", level=1))]
    fn add_heading(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        level: u32,
    ) -> PyResult<Py<PySubParagraph>> {
        let style = Some(format!("Heading {}", level.max(1)));
        let runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![crate::subdocbuilder::SubRun {
                text: text.to_string(),
                props: TextProps::default(),
            }]
        };
        let index = {
            let sd = slf.bind(py).borrow();
            let mut blocks = sd.blocks.borrow_mut();
            blocks.push(crate::subdocbuilder::Block::Paragraph { style, runs });
            blocks.len() - 1
        };
        Py::new(py, PySubParagraph { subdoc: slf, index })
    }

    /// Add a picture paragraph.
    #[pyo3(signature = (image_descriptor, width=None, height=None))]
    fn add_picture(
        &self,
        image_descriptor: &Bound<'_, PyAny>,
        width: Option<&Bound<'_, PyAny>>,
        height: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let filename = image_descriptor
            .extract::<String>()
            .ok()
            .and_then(|p| {
                std::path::Path::new(&p)
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
            });
        let blob = read_bytes_source(image_descriptor)?;
        self.blocks.borrow_mut().push(crate::subdocbuilder::Block::Picture {
            blob,
            filename,
            width: width.map(|w| extract_length(w)).transpose()?.flatten(),
            height: height.map(|h| extract_length(h)).transpose()?.flatten(),
        });
        Ok(())
    }

    /// Add a table with the given dimensions, returns a Table proxy.
    fn add_table(slf: Py<Self>, py: Python<'_>, rows: usize, cols: usize) -> PyResult<Py<PySubTable>> {
        let index = {
            let sd = slf.bind(py).borrow();
            let mut blocks = sd.blocks.borrow_mut();
            blocks.push(crate::subdocbuilder::Block::Table {
                rows: vec![vec![String::new(); cols]; rows],
            });
            blocks.len() - 1
        };
        Py::new(py, PySubTable { subdoc: slf, index })
    }

    fn __str__(&self) -> String {
        // the final xml is produced lazily during template rendering
        String::new()
    }

    fn __html__(&self) -> String {
        String::new()
    }
}

/// Paragraph proxy of a bound Subdoc.
#[pyclass(name = "SubParagraph", unsendable)]
pub struct PySubParagraph {
    pub subdoc: Py<PySubdoc>,
    pub index: usize,
}

#[pymethods]
impl PySubParagraph {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Paragraph { runs, .. } => {
                runs.iter().map(|r| r.text.clone()).collect()
            }
            _ => String::new(),
        })
    }

    #[pyo3(signature = (text=""))]
    fn add_run(&self, py: Python<'_>, text: &str) -> PyResult<Py<PySubRun>> {
        let run_index = edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Paragraph { runs, .. } => {
                runs.push(crate::subdocbuilder::SubRun {
                    text: text.to_string(),
                    props: TextProps::default(),
                });
                runs.len() - 1
            }
            _ => 0,
        })?;
        Py::new(
            py,
            PySubRun {
                subdoc: self.subdoc.clone_ref(py),
                index: self.index,
                run_index,
            },
        )
    }
}

/// Run proxy of a bound Subdoc paragraph.
#[pyclass(name = "SubRun", unsendable)]
pub struct PySubRun {
    pub subdoc: Py<PySubdoc>,
    pub index: usize,
    pub run_index: usize,
}

impl PySubRun {
    fn edit<R>(&self, py: Python<'_>, f: impl FnOnce(&mut crate::subdocbuilder::SubRun) -> R) -> PyResult<R> {
        edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Paragraph { runs, .. } => {
                let r = runs.get_mut(self.run_index).expect("run index");
                f(r)
            }
            _ => panic!("not a paragraph block"),
        })
    }
}

#[pymethods]
impl PySubRun {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.edit(py, |r| r.text.clone())
    }

    #[setter]
    fn set_bold(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| r.props.bold = v)?;
        Ok(())
    }
    #[setter]
    fn set_italic(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| r.props.italic = v)?;
        Ok(())
    }
    #[setter]
    fn set_strike(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| r.props.strike = v)?;
        Ok(())
    }
    #[setter]
    fn set_subscript(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| r.props.subscript = v)?;
        Ok(())
    }
    #[setter]
    fn set_superscript(&self, py: Python<'_>, v: bool) -> PyResult<()> {
        self.edit(py, |r| r.props.superscript = v)?;
        Ok(())
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| r.text = v)?;
        Ok(())
    }


    #[setter]
    fn set_underline(&self, py: Python<'_>, v: &Bound<'_, PyAny>) -> PyResult<()> {
        let val = if let Ok(s) = v.extract::<String>() {
            Some(s)
        } else if v.is_truthy()? {
            Some("single".to_string())
        } else {
            None
        };
        self.edit(py, |r| r.props.underline = val)?;
        Ok(())
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| r.props.style = Some(v))?;
        Ok(())
    }

    #[setter]
    fn set_font(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| r.props.font = Some(v))?;
        Ok(())
    }

    #[setter]
    fn set_size(&self, py: Python<'_>, v: u32) -> PyResult<()> {
        self.edit(py, |r| r.props.size = Some(v))?;
        Ok(())
    }

    #[setter]
    fn set_color(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| r.props.color = Some(v))?;
        Ok(())
    }

    #[setter]
    fn set_highlight(&self, py: Python<'_>, v: String) -> PyResult<()> {
        self.edit(py, |r| r.props.highlight = Some(v))?;
        Ok(())
    }
}

/// Table proxy of a bound Subdoc.
#[pyclass(name = "SubTable", unsendable)]
pub struct PySubTable {
    pub subdoc: Py<PySubdoc>,
    pub index: usize,
}

#[pymethods]
impl PySubTable {
    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Vec<PySubTableRow>> {
        let n = edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Table { rows } => rows.len(),
            _ => 0,
        })?;
        Ok((0..n)
            .map(|row| PySubTableRow {
                subdoc: self.subdoc.clone_ref(py),
                index: self.index,
                row,
            })
            .collect())
    }

    fn add_row(&self, py: Python<'_>) -> PyResult<PySubTableRow> {
        let row = edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Table { rows } => {
                let cols = rows.first().map(|r| r.len()).unwrap_or(1);
                rows.push(vec![String::new(); cols]);
                rows.len() - 1
            }
            _ => 0,
        })?;
        Ok(PySubTableRow {
            subdoc: self.subdoc.clone_ref(py),
            index: self.index,
            row,
        })
    }

    fn cell(&self, py: Python<'_>, i: usize, j: usize) -> PySubTableCell {
        PySubTableCell {
            subdoc: self.subdoc.clone_ref(py),
            index: self.index,
            row: i,
            col: j,
        }
    }
}

/// Table row proxy of a bound Subdoc table.
#[pyclass(name = "SubTableRow", unsendable)]
pub struct PySubTableRow {
    pub subdoc: Py<PySubdoc>,
    pub index: usize,
    pub row: usize,
}

#[pymethods]
impl PySubTableRow {
    #[getter]
    fn cells(&self, py: Python<'_>) -> PyResult<Vec<PySubTableCell>> {
        let n = edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Table { rows } => {
                rows.get(self.row).map(|r| r.len()).unwrap_or(0)
            }
            _ => 0,
        })?;
        Ok((0..n)
            .map(|col| PySubTableCell {
                subdoc: self.subdoc.clone_ref(py),
                index: self.index,
                row: self.row,
                col,
            })
            .collect())
    }
}

/// Table cell proxy of a bound Subdoc table.
#[pyclass(name = "SubTableCell", unsendable)]
pub struct PySubTableCell {
    pub subdoc: Py<PySubdoc>,
    pub index: usize,
    pub row: usize,
    pub col: usize,
}

#[pymethods]
impl PySubTableCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Table { rows } => rows
                .get(self.row)
                .and_then(|r| r.get(self.col))
                .cloned()
                .unwrap_or_default(),
            _ => String::new(),
        })
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, v: String) -> PyResult<()> {
        edit_block(&self.subdoc, py, self.index, |b| match b {
            crate::subdocbuilder::Block::Table { rows } => {
                if let Some(cell) = rows.get_mut(self.row).and_then(|r| r.get_mut(self.col)) {
                    *cell = v;
                }
            }
            _ => {}
        })?;
        Ok(())
    }
}

// ---------------------------------------------------------------- Composer

/// docxcompose-style document concatenation: append whole docx documents
/// to the end of a master document.
///
/// ```python
/// composer = Composer("master.docx")
/// composer.append("chapter1.docx")
/// composer.append("chapter2.docx")
/// composer.save("out.docx")
/// ```
#[pyclass(name = "Composer", unsendable)]
pub struct PyComposer {
    inner: RefCell<crate::composer::Composer>,
}

#[pymethods]
impl PyComposer {
    /// master: path, bytes, or file-like object with the master docx.
    #[new]
    fn new(master: &Bound<'_, PyAny>) -> PyResult<Self> {
        let bytes = read_bytes_source(master)?;
        let inner = crate::composer::Composer::new(bytes).map_err(to_pyerr)?;
        Ok(PyComposer {
            inner: RefCell::new(inner),
        })
    }

    /// Append one whole docx (path, bytes, or file-like) to the master's
    /// body, preceded by a page break. Styles/numbering/media/footnotes are
    /// merged like docxcompose (style conflicts renamed `X_1`, the first
    /// list restarts numbering); the appended document's section properties
    /// (page setup, header/footer references) are dropped.
    fn append(&self, py: Python<'_>, doc: &Bound<'_, PyAny>) -> PyResult<()> {
        let bytes = read_bytes_source(doc)?;
        let inner = AssertSend(&self.inner);
        py.detach(move || {
            let inner = inner; // capture AssertSend as a whole
            inner.0.borrow_mut().append(&bytes)
        })
        .map_err(to_pyerr)
    }

    /// Save the composed docx to a path or file-like object.
    fn save(&self, py: Python<'_>, filename: &Bound<'_, PyAny>) -> PyResult<()> {
        let inner = AssertSend(&self.inner);
        let bytes = py
            .detach(move || {
                let inner = inner; // capture AssertSend as a whole
                inner.0.borrow_mut().save_bytes()
            })
            .map_err(to_pyerr)?;
        if let Ok(path) = filename.extract::<String>() {
            std::fs::write(&path, &bytes)
                .map_err(|e| PyValueError::new_err(format!("cannot write {}: {}", path, e)))?;
            return Ok(());
        }
        if let Ok(fspath) = filename.call_method0("__fspath__") {
            if let Ok(path) = fspath.extract::<String>() {
                std::fs::write(&path, &bytes)
                    .map_err(|e| PyValueError::new_err(format!("cannot write {}: {}", path, e)))?;
                return Ok(());
            }
        }
        let b = PyBytes::new(py, &bytes);
        filename.call_method1("write", (b,))?;
        Ok(())
    }
}
