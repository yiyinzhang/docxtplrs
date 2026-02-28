//! Type definitions and error handling for docxtplrs

use pyo3::exceptions::{PyIOError, PyRuntimeError, PyValueError};
use pyo3::PyErr;
use std::fmt;
use thiserror::Error;

/// Main error type for docxtplrs
#[derive(Error, Debug)]
pub enum DocxTplError {
    /// IO errors (file not found, permission denied, etc.)
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// ZIP archive errors
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// XML parsing errors
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// Template rendering errors
    #[error("Template error: {0}")]
    Template(String),

    /// JSON serialization errors
    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),

    /// Regex errors
    #[error("Regex error: {0}")]
    Regex(#[from] regex::Error),

    /// MiniJinja errors
    #[error("Template engine error: {0}")]
    MiniJinja(String),

    /// Invalid argument
    #[error("Invalid argument: {0}")]
    InvalidArgument(String),

    /// Template variable not found
    #[error("Missing variable: {0}")]
    MissingVariable(String),

    /// Other errors
    #[error("{0}")]
    Other(String),
}

impl From<minijinja::Error> for DocxTplError {
    fn from(err: minijinja::Error) -> Self {
        DocxTplError::MiniJinja(err.to_string())
    }
}

impl From<DocxTplError> for PyErr {
    fn from(err: DocxTplError) -> PyErr {
        match err {
            DocxTplError::Io(e) => PyIOError::new_err(e.to_string()),
            DocxTplError::InvalidArgument(msg) | DocxTplError::MissingVariable(msg) => {
                PyValueError::new_err(msg)
            }
            _ => PyRuntimeError::new_err(err.to_string()),
        }
    }
}

/// Result type alias for docxtplrs
pub type Result<T> = std::result::Result<T, DocxTplError>;

/// Document part types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocxPart {
    Document,
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
}

impl fmt::Display for DocxPart {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DocxPart::Document => write!(f, "document"),
            DocxPart::Header => write!(f, "header"),
            DocxPart::Footer => write!(f, "footer"),
            DocxPart::Footnotes => write!(f, "footnotes"),
            DocxPart::Endnotes => write!(f, "endnotes"),
            DocxPart::Comments => write!(f, "comments"),
        }
    }
}

/// Template tag types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TagType {
    /// Standard Jinja2 tag
    Normal,
    /// Paragraph tag ({%p %})
    Paragraph,
    /// Table row tag ({%tr %})
    TableRow,
    /// Table cell tag ({%tc %})
    TableCell,
    /// Run tag ({%r %})
    Run,
}

impl fmt::Display for TagType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TagType::Normal => write!(f, ""),
            TagType::Paragraph => write!(f, "p"),
            TagType::TableRow => write!(f, "tr"),
            TagType::TableCell => write!(f, "tc"),
            TagType::Run => write!(f, "r"),
        }
    }
}

impl TagType {
    /// Parse a tag prefix (e.g., "p" from "{%p if ... %}")
    pub fn from_prefix(prefix: &str) -> Self {
        match prefix {
            "p" => TagType::Paragraph,
            "tr" => TagType::TableRow,
            "tc" => TagType::TableCell,
            "r" => TagType::Run,
            _ => TagType::Normal,
        }
    }

    /// Check if this is a special tag type (not Normal)
    pub fn is_special(&self) -> bool {
        !matches!(self, TagType::Normal)
    }
}

/// Measurement units for images
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Measurement {
    Millimeters(f64),
    Inches(f64),
    Points(f64),
    Emus(i64),
}

impl Measurement {
    /// Convert to EMUs (English Metric Units used by Office Open XML)
    pub fn to_emus(&self) -> i64 {
        match self {
            Measurement::Millimeters(mm) => (*mm * 36000.0) as i64,
            Measurement::Inches(inches) => (*inches * 914400.0) as i64,
            Measurement::Points(points) => (*points * 12700.0) as i64,
            Measurement::Emus(emus) => *emus,
        }
    }

    /// Convert to points
    pub fn to_points(&self) -> f64 {
        match self {
            Measurement::Millimeters(mm) => *mm * 2.83465,
            Measurement::Inches(inches) => *inches * 72.0,
            Measurement::Points(points) => *points,
            Measurement::Emus(emus) => *emus as f64 / 12700.0,
        }
    }
}

impl fmt::Display for Measurement {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Measurement::Millimeters(v) => write!(f, "{:.2}mm", v),
            Measurement::Inches(v) => write!(f, "{:.2}in", v),
            Measurement::Points(v) => write!(f, "{:.2}pt", v),
            Measurement::Emus(v) => write!(f, "{}emu", v),
        }
    }
}

/// WordprocessingML namespace constants
pub mod namespaces {
    pub const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    pub const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    pub const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    pub const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    pub const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
}
