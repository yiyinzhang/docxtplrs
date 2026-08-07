//! docxtplrs: Rust implementation of python-docx-template (docxtpl)
//! with Python bindings via PyO3.

#[cfg(feature = "python")]
mod doccomments;
#[cfg(feature = "python")]
mod docmodel;
#[cfg(feature = "python")]
mod docmodel_add;
#[cfg(feature = "python")]
mod docmodel_fmt;
pub mod composer;
pub mod gettext;
pub mod image;
mod inline_image;
#[cfg(feature = "python")]
mod jutils;
pub mod package;
pub mod patch;
#[cfg(feature = "python")]
mod pybridge;
#[cfg(feature = "python")]
mod pyclasses;
#[cfg(feature = "python")]
mod pyxml;
pub mod richtext;
mod subdoc;
mod subdocbuilder;
pub mod template;
pub mod xmldom;

#[cfg(feature = "python")]
use pyo3::prelude::*;

#[cfg(feature = "python")]
#[pymodule]
fn docxtplrs(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", "0.2.1")?;

    m.add_class::<pyclasses::PyDocxTemplate>()?;
    m.add_class::<pyclasses::PyRichText>()?;
    m.add_class::<pyclasses::PyRichTextParagraph>()?;
    m.add_class::<pyclasses::PyListing>()?;
    m.add_class::<pyclasses::PyInlineImage>()?;
    m.add_class::<pyclasses::PySubdoc>()?;
    m.add_class::<pyclasses::PyComposer>()?;
    m.add_class::<pyclasses::PyLength>()?;
    m.add_class::<pyclasses::PySubParagraph>()?;
    m.add_class::<pyclasses::PySubRun>()?;
    m.add_class::<pyclasses::PySubTable>()?;
    m.add_class::<pyclasses::PySubTableRow>()?;
    m.add_class::<pyclasses::PySubTableCell>()?;
    m.add_class::<docmodel::PyDocument>()?;
    m.add_class::<docmodel::PyParagraph>()?;
    m.add_class::<docmodel::PyRun>()?;
    m.add_class::<docmodel::PyTable>()?;
    m.add_class::<docmodel::PyTableRow>()?;
    m.add_class::<docmodel::PyCell>()?;
    m.add_class::<docmodel::PySection>()?;
    m.add_class::<docmodel::PyStyle>()?;
    m.add_class::<docmodel::PyInlineShape>()?;
    m.add_class::<docmodel::PyCoreProperties>()?;
    m.add_class::<docmodel::PySectionHdrFtr>()?;
    m.add_class::<docmodel::PyStyles>()?;
    m.add_class::<docmodel::PyStyleFont>()?;
    m.add_class::<docmodel_fmt::PyFont>()?;
    m.add_class::<docmodel_fmt::PyColorFormat>()?;
    m.add_class::<docmodel_fmt::PyParagraphFormat>()?;
    m.add_class::<docmodel_fmt::PyTabStops>()?;
    m.add_class::<docmodel_fmt::PyTabStop>()?;
    m.add_class::<docmodel_fmt::PyTableColumn>()?;
    m.add_class::<docmodel_fmt::PyCellParagraph>()?;
    m.add_class::<docmodel_fmt::PyHyperlink>()?;
    m.add_class::<docmodel_fmt::PyCellTable>()?;
    m.add_class::<docmodel_fmt::PyNestedRow>()?;
    m.add_class::<docmodel_fmt::PyNestedCell>()?;
    m.add_class::<docmodel_fmt::PyRenderedPageBreak>()?;
    m.add_class::<docmodel_fmt::PyPart>()?;
    m.add_class::<docmodel_fmt::PyField>()?;
    m.add_class::<docmodel::PySettings>()?;
    m.add_class::<doccomments::PyComments>()?;
    m.add_class::<doccomments::PyComment>()?;
    m.add_class::<pyxml::PyXmlElement>()?;

    // aliases like docxtpl: R = RichText, RP = RichTextParagraph
    m.add("R", m.getattr("RichText")?)?;
    m.add("RP", m.getattr("RichTextParagraph")?)?;

    m.add(
        "TemplateError",
        m.py().get_type::<pyclasses::TemplateError>(),
    )?;

    m.add_function(wrap_pyfunction!(pyclasses::emu, m)?)?;
    m.add_function(wrap_pyfunction!(pyclasses::inches, m)?)?;
    m.add_function(wrap_pyfunction!(pyclasses::cm, m)?)?;
    m.add_function(wrap_pyfunction!(pyclasses::mm, m)?)?;
    m.add_function(wrap_pyfunction!(pyclasses::pt, m)?)?;
    m.add_function(wrap_pyfunction!(pyclasses::twips, m)?)?;

    // jinja2.utils equivalents
    m.add_class::<jutils::PyCycler>()?;
    m.add_class::<jutils::PyJoiner>()?;
    m.add_function(wrap_pyfunction!(jutils::generate_lorem_ipsum, m)?)?;

    Ok(())
}
