//! Jinja2 Environment for custom filters

use pyo3::prelude::*;
use std::collections::HashMap;
use std::sync::Arc;

/// Jinja2 Environment for managing custom filters
///
/// This class allows you to define custom filters that can be used
/// in your templates. Filters are Python functions that transform values.
///
/// Example:
///     from docxtplrs import DocxTemplate, JinjaEnv
///
///     def format_currency(value):
///         return f"${value:,.2f}"
///
///     env = JinjaEnv()
///     env.add_filter("currency", format_currency)
///
///     doc = DocxTemplate("template.docx")
///     doc.render(context, jinja_env=env)
#[pyclass(name = "JinjaEnv")]
#[derive(Debug)]
pub struct JinjaEnv {
    /// Map of filter name to Python callable (wrapped in Arc for thread safety)
    filters: HashMap<String, Arc<PyObject>>,
}

#[pymethods]
impl JinjaEnv {
    /// Create a new Jinja2 environment
    #[new]
    fn new() -> Self {
        Self {
            filters: HashMap::new(),
        }
    }

    /// Add a custom filter
    ///
    /// Args:
    ///     name: The name of the filter (used in templates like {{ value|filter_name }})
    ///     func: A Python callable that takes the value as argument
    ///
    /// Example:
    ///     def uppercase(value):
    ///         return str(value).upper()
    ///
    ///     env.add_filter("upper", uppercase)
    ///     # In template: {{ name|upper }}
    fn add_filter(&mut self, name: String, func: PyObject) -> PyResult<()> {
        Python::with_gil(|py| {
            // Verify it's callable
            if !func.bind(py).is_callable() {
                return Err(pyo3::exceptions::PyTypeError::new_err(
                    "Filter must be callable",
                ));
            }
            self.filters.insert(name, Arc::new(func));
            Ok(())
        })
    }

    /// Remove a filter
    ///
    /// Args:
    ///     name: The name of the filter to remove
    fn remove_filter(&mut self, name: &str) -> PyResult<()> {
        if self.filters.remove(name).is_none() {
            return Err(pyo3::exceptions::PyKeyError::new_err(format!(
                "Filter '{}' not found",
                name
            )));
        }
        Ok(())
    }

    /// Get all filter names
    ///
    /// Returns:
    ///     List of filter names
    fn get_filter_names(&self) -> Vec<String> {
        self.filters.keys().cloned().collect()
    }

    /// Check if a filter exists
    ///
    /// Args:
    ///     name: The filter name to check
    ///
    /// Returns:
    ///     True if the filter exists
    fn has_filter(&self, name: &str) -> bool {
        self.filters.contains_key(name)
    }

    /// Clear all filters
    fn clear_filters(&mut self) {
        self.filters.clear();
    }

    fn __repr__(&self) -> String {
        format!("JinjaEnv(filters={})", self.filters.len())
    }
}

impl JinjaEnv {
    /// Get a filter function by name
    pub fn get_filter(&self, name: &str) -> Option<Arc<PyObject>> {
        self.filters.get(name).cloned()
    }

    /// Get all filters as a new Arc HashMap
    pub fn get_filters_arc(&self) -> Arc<HashMap<String, Arc<PyObject>>> {
        Arc::new(self.filters.clone())
    }

    /// Get filters count
    pub fn filter_count(&self) -> usize {
        self.filters.len()
    }
}

impl Default for JinjaEnv {
    fn default() -> Self {
        Self::new()
    }
}

impl Clone for JinjaEnv {
    fn clone(&self) -> Self {
        Self {
            filters: self.filters.clone(),
        }
    }
}
