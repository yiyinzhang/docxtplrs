//! jinja2.utils equivalents: Cycler, Joiner, generate_lorem_ipsum.

use pyo3::prelude::*;
use std::cell::RefCell;

/// jinja2.utils.Cycler
#[pyclass(name = "Cycler", unsendable)]
pub struct PyCycler {
    items: Vec<Py<PyAny>>,
    pos: RefCell<usize>,
}

#[pymethods]
impl PyCycler {
    #[new]
    #[pyo3(signature = (*items))]
    fn new(items: &Bound<'_, pyo3::types::PyTuple>) -> Self {
        PyCycler {
            items: items.iter().map(|i| i.unbind()).collect(),
            pos: RefCell::new(0),
        }
    }

    fn next(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let item = {
            let this = slf.bind(py).borrow();
            if this.items.is_empty() {
                return Err(pyo3::exceptions::PyStopIteration::new_err("no items"));
            }
            let mut pos = this.pos.borrow_mut();
            let item = this.items[*pos % this.items.len()].clone_ref(py);
            *pos += 1;
            item
        };
        Ok(item)
    }

    fn __next__(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyAny>> {
        Self::next(slf, py)
    }

    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn reset(&self) {
        *self.pos.borrow_mut() = 0;
    }

    #[getter]
    fn current(&self, py: Python<'_>) -> Option<Py<PyAny>> {
        if self.items.is_empty() {
            return None;
        }
        let pos = *self.pos.borrow();
        let idx = if pos == 0 { self.items.len() - 1 } else { (pos - 1) % self.items.len() };
        Some(self.items[idx].clone_ref(py))
    }
}

/// jinja2.utils.Joiner
#[pyclass(name = "Joiner", unsendable)]
pub struct PyJoiner {
    sep: String,
    used: RefCell<bool>,
}

#[pymethods]
impl PyJoiner {
    #[new]
    #[pyo3(signature = (sep=", "))]
    fn new(sep: &str) -> Self {
        PyJoiner {
            sep: sep.to_string(),
            used: RefCell::new(false),
        }
    }

    #[pyo3(signature = (*args))]
    fn __call__(&self, args: &Bound<'_, pyo3::types::PyTuple>) -> PyResult<String> {
        let parts: Vec<String> = args
            .iter()
            .map(|a| a.str().map(|s| s.to_string_lossy().to_string()))
            .collect::<PyResult<_>>()?;
        let mut used = self.used.borrow_mut();
        if *used {
            Ok(format!("{}{}", self.sep, parts.join(&self.sep)))
        } else {
            *used = true;
            Ok(parts.join(&self.sep))
        }
    }
}

const LOREM_WORDS: &[&str] = &[
    "lorem", "ipsum", "dolor", "sit", "amet", "consectetur", "adipiscing", "elit",
    "sed", "do", "eiusmod", "tempor", "incididunt", "ut", "labore", "et", "dolore",
    "magna", "aliqua", "enim", "ad", "minim", "veniam", "quis", "nostrud",
    "exercitation", "ullamco", "laboris", "nisi", "aliquip", "ex", "ea", "commodo",
    "consequat", "duis", "aute", "irure", "in", "reprehenderit", "voluptate",
    "velit", "esse", "cillum", "fugiat", "nulla", "pariatur", "excepteur", "sint",
    "occaecat", "cupidatat", "non", "proident", "sunt", "culpa", "qui", "officia",
    "deserunt", "mollit", "anim", "id", "est", "laborum",
];

/// jinja2.utils.generate_lorem_ipsum(n=5, html=True, min=20, max=100)
#[pyfunction]
#[pyo3(signature = (n=5, html=true, min=20, max=100))]
pub fn generate_lorem_ipsum(n: usize, html: bool, min: usize, max: usize) -> String {
    let mut seed = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.subsec_nanos() as u64 ^ d.as_secs())
        .unwrap_or(0x9E3779B9);
    let mut rand = move || {
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
        (seed >> 33) as usize
    };
    let mut out = String::new();
    for _ in 0..n.max(1) {
        let count = min + rand() % (max.saturating_sub(min).max(1) + 1);
        let words: Vec<&str> = (0..count)
            .map(|_| LOREM_WORDS[rand() % LOREM_WORDS.len()])
            .collect();
        let mut sentence = words.join(" ");
        if let Some(c) = sentence.chars().next() {
            sentence = c.to_uppercase().collect::<String>() + &sentence[1..];
        }
        sentence.push('.');
        if html {
            out.push_str(&format!("<p>{}</p>\n", sentence));
        } else {
            out.push_str(&sentence);
            out.push_str("\n\n");
        }
    }
    out.trim_end().to_string()
}
