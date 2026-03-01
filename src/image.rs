//! Inline image handling for Word documents

use crate::types::{DocxTplError, Measurement, Result};
use pyo3::prelude::*;
use std::fs;
use std::path::Path;

/// Image format detection
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Wmf,
    Emf,
    Svg,
    Unknown,
}

impl ImageFormat {
    /// Detect image format from file extension
    pub fn from_path(path: &Path) -> Self {
        match path.extension() {
            Some(ext) => match ext.to_str().unwrap_or("").to_lowercase().as_str() {
                "png" => ImageFormat::Png,
                "jpg" | "jpeg" => ImageFormat::Jpeg,
                "gif" => ImageFormat::Gif,
                "bmp" => ImageFormat::Bmp,
                "tiff" | "tif" => ImageFormat::Tiff,
                "wmf" => ImageFormat::Wmf,
                "emf" => ImageFormat::Emf,
                "svg" => ImageFormat::Svg,
                _ => ImageFormat::Unknown,
            },
            None => ImageFormat::Unknown,
        }
    }

    /// Get MIME type for the image format
    pub fn mime_type(&self) -> &'static str {
        match self {
            ImageFormat::Png => "image/png",
            ImageFormat::Jpeg => "image/jpeg",
            ImageFormat::Gif => "image/gif",
            ImageFormat::Bmp => "image/bmp",
            ImageFormat::Tiff => "image/tiff",
            ImageFormat::Wmf => "image/x-wmf",
            ImageFormat::Emf => "image/x-emf",
            ImageFormat::Svg => "image/svg+xml",
            ImageFormat::Unknown => "application/octet-stream",
        }
    }

    /// Get file extension
    pub fn extension(&self) -> &'static str {
        match self {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpg",
            ImageFormat::Gif => "gif",
            ImageFormat::Bmp => "bmp",
            ImageFormat::Tiff => "tiff",
            ImageFormat::Wmf => "wmf",
            ImageFormat::Emf => "emf",
            ImageFormat::Svg => "svg",
            ImageFormat::Unknown => "bin",
        }
    }

    /// Get content type for [Content_Types].xml
    pub fn content_type(&self) -> &'static str {
        self.mime_type()
    }
}

/// Image dimensions
#[derive(Debug, Clone, Copy)]
pub struct ImageDimensions {
    pub width: i64,  // In EMUs
    pub height: i64, // In EMUs
    pub original_width_px: u32,
    pub original_height_px: u32,
}

impl ImageDimensions {
    /// Create from width and height measurements
    pub fn from_measurements(
        width: Option<Measurement>,
        height: Option<Measurement>,
        original_width_px: u32,
        original_height_px: u32,
    ) -> Self {
        let orig_width_emu = (original_width_px as f64 * 9525.0) as i64; // 1 px = 9525 EMUs at 96 DPI
        let orig_height_emu = (original_height_px as f64 * 9525.0) as i64;

        let (width_emu, height_emu) = match (width, height) {
            (Some(w), Some(h)) => (w.to_emus(), h.to_emus()),
            (Some(w), None) => {
                let w_emu = w.to_emus();
                let h_emu = (w_emu as f64 * orig_height_emu as f64 / orig_width_emu as f64) as i64;
                (w_emu, h_emu)
            }
            (None, Some(h)) => {
                let h_emu = h.to_emus();
                let w_emu = (h_emu as f64 * orig_width_emu as f64 / orig_height_emu as f64) as i64;
                (w_emu, h_emu)
            }
            (None, None) => (orig_width_emu, orig_height_emu),
        };

        Self {
            width: width_emu,
            height: height_emu,
            original_width_px,
            original_height_px,
        }
    }
}

/// Read image dimensions from file
pub fn read_image_dimensions(path: &Path) -> Result<(u32, u32)> {
    let data = fs::read(path)?;

    // Try PNG
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        if let Some(dimensions) = read_png_dimensions(&data) {
            return Ok(dimensions);
        }
    }

    // Try JPEG
    if data.starts_with(b"\xff\xd8") {
        if let Some(dimensions) = read_jpeg_dimensions(&data) {
            return Ok(dimensions);
        }
    }

    // Try GIF
    if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        if data.len() >= 10 {
            let width = u16::from_le_bytes([data[6], data[7]]) as u32;
            let height = u16::from_le_bytes([data[8], data[9]]) as u32;
            return Ok((width, height));
        }
    }

    // Try BMP
    if data.starts_with(b"BM") && data.len() >= 26 {
        let width = u32::from_le_bytes([data[18], data[19], data[20], data[21]]);
        let height = u32::from_le_bytes([data[22], data[23], data[24], data[25]]);
        return Ok((width, height));
    }

    Err(DocxTplError::InvalidArgument(format!(
        "Could not read dimensions for image: {}",
        path.display()
    )))
}

fn read_png_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    if data.len() < 24 {
        return None;
    }
    // PNG dimensions are in the IHDR chunk, starting at byte 16
    let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
    let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
    Some((width, height))
}

fn read_jpeg_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    let mut i = 2;
    while i < data.len() {
        if i + 4 > data.len() {
            break;
        }
        let marker = data[i];
        if marker == 0xFF {
            let segment_type = data[i + 1];
            // Skip padding
            if segment_type == 0xFF {
                i += 1;
                continue;
            }
            // SOF markers (Start of Frame)
            if (0xC0..=0xCF).contains(&segment_type) && segment_type != 0xC4 && segment_type != 0xC8
            {
                if i + 9 < data.len() {
                    let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
                    let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
                    return Some((width, height));
                }
            }
            // Skip this segment
            if i + 4 < data.len() {
                let len = u16::from_be_bytes([data[i + 2], data[i + 3]]) as usize;
                i += 2 + len;
            } else {
                break;
            }
        } else {
            break;
        }
    }
    None
}

/// InlineImage for inserting images into Word documents
///
/// This struct represents an image that can be inserted into a Word document
/// template using the {{ variable }} syntax.
#[pyclass(name = "InlineImage")]
#[derive(Debug, Clone)]
pub struct InlineImage {
    pub image_path: String,
    pub width: Option<i64>, // In EMUs
    pub height: Option<i64>, // In EMUs
    pub image_data: Vec<u8>,
    pub format: ImageFormat,
    pub dimensions: ImageDimensions,
    pub relationship_id: Option<String>,
}

#[pymethods]
impl InlineImage {
    /// Create a new InlineImage object
    ///
    /// Args:
    ///     template: The DocxTemplate object (for relationship management)
    ///     image_descriptor: Path to the image file
    ///     width: Optional width (use Mm, Inches, or Pt classes)
    ///     height: Optional height (use Mm, Inches, or Pt classes)
    #[new]
    #[pyo3(signature = (template, image_descriptor, width=None, height=None))]
    fn new(
        template: &crate::template::DocxTemplate,
        image_descriptor: String,
        width: Option<PyObject>,
        height: Option<PyObject>,
    ) -> PyResult<Self> {
        let path = Path::new(&image_descriptor);

        if !path.exists() {
            return Err(DocxTplError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("Image not found: {}", image_descriptor),
            ))
            .into());
        }

        let format = ImageFormat::from_path(path);
        let (orig_width, orig_height) = read_image_dimensions(path)
            .map_err(|e| DocxTplError::Other(e.to_string()))?;

        // Convert Python measurement objects to Measurement
        let width_meas = width
            .map(|w| python_to_measurement(&w))
            .transpose()
            .map_err(|e| DocxTplError::InvalidArgument(e.to_string()))?;

        let height_meas = height
            .map(|h| python_to_measurement(&h))
            .transpose()
            .map_err(|e| DocxTplError::InvalidArgument(e.to_string()))?;

        let dimensions =
            ImageDimensions::from_measurements(width_meas, height_meas, orig_width, orig_height);

        let image_data = fs::read(path)
            .map_err(|e| DocxTplError::Io(e))?;

        Ok(Self {
            image_path: image_descriptor,
            width: width_meas.map(|m| m.to_emus()),
            height: height_meas.map(|m| m.to_emus()),
            image_data,
            format,
            dimensions,
            relationship_id: None,
        })
    }

    fn __repr__(&self) -> String {
        format!(
            "InlineImage(path='{}', format={:?}, {}x{} EMUs)",
            self.image_path,
            self.format,
            self.dimensions.width,
            self.dimensions.height
        )
    }
}

/// Convert a Python measurement object to Measurement
fn python_to_measurement(obj: &PyObject) -> PyResult<Measurement> {
    Python::with_gil(|py| {
        // Check if it's a raw number (treat as millimeters)
        if let Ok(mm) = obj.extract::<f64>(py) {
            return Ok(Measurement::Millimeters(mm));
        }

        // Try to get the value from measurement objects
        // Cm, Mm, Inches, Pt classes from docxtplrs
        if let Ok(bound) = obj.bind(py).getattr("cm") {
            if let Ok(v) = bound.extract::<f64>() {
                return Ok(Measurement::Centimeters(v));
            }
        }

        if let Ok(bound) = obj.bind(py).getattr("mm") {
            if let Ok(v) = bound.extract::<f64>() {
                return Ok(Measurement::Millimeters(v));
            }
        }

        if let Ok(bound) = obj.bind(py).getattr("inches") {
            if let Ok(v) = bound.extract::<f64>() {
                return Ok(Measurement::Inches(v));
            }
        }

        if let Ok(bound) = obj.bind(py).getattr("pt") {
            if let Ok(v) = bound.extract::<f64>() {
                return Ok(Measurement::Points(v));
            }
        }

        // Try EMUs directly
        if let Ok(bound) = obj.bind(py).getattr("emu") {
            if let Ok(v) = bound.extract::<i64>() {
                return Ok(Measurement::Emus(v));
            }
        }

        // Last resort - try extracting as f64
        if let Ok(v) = obj.extract::<f64>(py) {
            return Ok(Measurement::Points(v));
        }

        Err(pyo3::exceptions::PyTypeError::new_err(
            "Invalid measurement type. Use Cm, Mm, Inches, or Pt from docxtplrs.",
        ))
    })
}

impl InlineImage {
    /// Generate the WordprocessingML for this image
    pub fn to_xml(&self, rel_id: &str, doc_pr_id: u32) -> String {
        let cx = self.dimensions.width;
        let cy = self.dimensions.height;

        format!(
            r#"<w:drawing>
                <wp:inline distT="0" distB="0" distL="0" distR="0">
                    <wp:extent cx="{}" cy="{}"/>
                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
                    <wp:docPr id="{}" name="Picture {}"/>
                    <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                    </wp:cNvGraphicFramePr>
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:nvPicPr>
                                    <pic:cNvPr id="0" name="{}"/>
                                    <pic:cNvPicPr/>
                                </pic:nvPicPr>
                                <pic:blipFill>
                                    <a:blip r:embed="{}"/>
                                    <a:stretch>
                                        <a:fillRect/>
                                    </a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                    <a:xfrm>
                                        <a:off x="0" y="0"/>
                                        <a:ext cx="{}" cy="{}"/>
                                    </a:xfrm>
                                    <a:prstGeom prst="rect">
                                        <a:avLst/>
                                    </a:prstGeom>
                                </pic:spPr>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing>"#,
            cx,
            cy,
            doc_pr_id,
            doc_pr_id,
            self.image_path,
            rel_id,
            cx,
            cy
        )
    }

    /// Get the file name for embedding
    pub fn file_name(&self) -> String {
        Path::new(&self.image_path)
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| format!("image.{}", self.format.extension()))
    }
}

/// Measurement helper classes for Python
#[pyclass(name = "Mm")]
#[derive(Debug, Clone, Copy)]
pub struct Mm {
    value: f64,
}

#[pymethods]
impl Mm {
    #[new]
    fn new(value: f64) -> Self {
        Self { value }
    }

    fn __float__(&self) -> f64 {
        self.value
    }

    fn __repr__(&self) -> String {
        format!("Mm({})", self.value)
    }
}

#[pyclass(name = "Cm")]
#[derive(Debug, Clone, Copy)]
pub struct Cm {
    value: f64,
}

#[pymethods]
impl Cm {
    #[new]
    fn new(value: f64) -> Self {
        Self { value }
    }

    fn __float__(&self) -> f64 {
        self.value
    }

    fn __repr__(&self) -> String {
        format!("Cm({})", self.value)
    }
}

#[pyclass(name = "Inches")]
#[derive(Debug, Clone, Copy)]
pub struct Inches {
    value: f64,
}

#[pymethods]
impl Inches {
    #[new]
    fn new(value: f64) -> Self {
        Self { value }
    }

    fn __float__(&self) -> f64 {
        self.value
    }

    fn __repr__(&self) -> String {
        format!("Inches({})", self.value)
    }
}

#[pyclass(name = "Pt")]
#[derive(Debug, Clone, Copy)]
pub struct Pt {
    value: f64,
}

#[pymethods]
impl Pt {
    #[new]
    fn new(value: f64) -> Self {
        Self { value }
    }

    fn __float__(&self) -> f64 {
        self.value
    }

    fn __repr__(&self) -> String {
        format!("Pt({})", self.value)
    }
}
