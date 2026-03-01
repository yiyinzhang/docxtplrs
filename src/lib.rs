//! docxtplrs - A Rust implementation of python-docx-template with Python bindings
//!
//! This crate provides functionality for generating Word documents from templates
//! using Jinja2-like syntax, similar to the Python docxtpl library.
//!
//! # Example
//!
//! ```python
//! from docxtplrs import DocxTemplate
//!
//! doc = DocxTemplate("template.docx")
//! context = {"name": "World", "company": "Example Corp"}
//! doc.render(context)
//! doc.save("output.docx")
//! ```

mod image;
mod jinja_env;
mod richtext;
mod subdoc;
mod template;
mod types;
mod xml_utils;

use pyo3::prelude::*;

// Re-export Python classes
pub use crate::image::{Cm, ImageFormat, Inches, InlineImage, Mm, Pt};
pub use crate::jinja_env::JinjaEnv;
pub use crate::richtext::{r_shortcut, rp_shortcut, Listing, RichText, RichTextParagraph};
pub use crate::subdoc::{
    CellColor, ColSpan, DocumentBuilder, Subdoc, VerticalMerge,
};
pub use crate::template::DocxTemplate;

/// Escape XML special characters
///
/// Args:
///     text: The text to escape
///
/// Returns:
///     Escaped text safe for XML
#[pyfunction]
fn escape_xml(text: String) -> String {
    xml_utils::escape_xml(&text)
}

/// Unescape XML entities
///
/// Args:
///     text: The text to unescape
///
/// Returns:
///     Unescaped text
#[pyfunction]
fn unescape_xml(text: String) -> String {
    xml_utils::unescape_xml(&text)
}

/// Get the version of the docxtplrs library
#[pyfunction]
fn version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

/// A Rust implementation of python-docx-template with Python bindings
///
/// This module provides functionality for generating Word documents from templates
/// using Jinja2-like syntax.
///
/// Main Classes:
///     DocxTemplate: Load and render Word document templates
///     RichText: Create styled text content
///     InlineImage: Insert images into documents
///     Subdoc: Embed sub-documents
///     Listing: Insert escaped text with formatting
///
/// Measurement Classes:
///     Mm: Millimeters
///     Cm: Centimeters
///     Inches: Inches
///     Pt: Points
///
/// Helper Functions:
///     R: Shortcut for RichText
///     RP: Shortcut for RichTextParagraph
///     escape_xml: Escape XML special characters
///     unescape_xml: Unescape XML entities
#[pymodule]
fn docxtplrs(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Core template class
    m.add_class::<DocxTemplate>()?;

    // Jinja environment class
    m.add_class::<JinjaEnv>()?;

    // Rich text classes
    m.add_class::<RichText>()?;
    m.add_class::<RichTextParagraph>()?;
    m.add_class::<Listing>()?;

    // Image classes
    m.add_class::<InlineImage>()?;
    m.add_class::<Mm>()?;
    m.add_class::<Cm>()?;
    m.add_class::<Inches>()?;
    m.add_class::<Pt>()?;

    // Subdoc classes
    m.add_class::<Subdoc>()?;
    m.add_class::<DocumentBuilder>()?;
    m.add_class::<CellColor>()?;
    m.add_class::<ColSpan>()?;
    m.add_class::<VerticalMerge>()?;

    // Helper functions
    m.add_function(wrap_pyfunction!(r_shortcut, m)?)?;
    m.add_function(wrap_pyfunction!(rp_shortcut, m)?)?;
    m.add_function(wrap_pyfunction!(escape_xml, m)?)?;
    m.add_function(wrap_pyfunction!(unescape_xml, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;

    // Version
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;

    Ok(())
}
