"""
docxtplrs - A Rust implementation of python-docx-template with Python bindings

This package provides functionality for generating Word documents from templates
using Jinja2-like syntax.

Example:
    >>> from docxtplrs import DocxTemplate
    >>> doc = DocxTemplate("template.docx")
    >>> context = {"name": "World", "company": "Example Corp"}
    >>> doc.render(context)
    >>> doc.save("output.docx")
"""

from docxtplrs.docxtplrs import (
    DocxTemplate,
    RichText,
    RichTextParagraph,
    InlineImage,
    Subdoc,
    DocumentBuilder,
    Listing,
    CellColor,
    ColSpan,
    VerticalMerge,
    Mm,
    Inches,
    Pt,
    R,
    RP,
    escape_xml,
    unescape_xml,
    version,
)

__version__ = version()

__all__ = [
    "DocxTemplate",
    "RichText",
    "RichTextParagraph",
    "InlineImage",
    "Subdoc",
    "DocumentBuilder",
    "Listing",
    "CellColor",
    "ColSpan",
    "VerticalMerge",
    "Mm",
    "Inches",
    "Pt",
    "R",
    "RP",
    "escape_xml",
    "unescape_xml",
    "version",
]
