"""Subdoc endnotes merging (mirrors the footnotes merge tests)."""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, docx_names, tp, XML_DECL, NSDECL
from test_subdoc_merge import build_subdoc, render_with_sub
from test_subdoc_sections_comments import build_doc

from docxtplrs import DocxTemplate

ENDNOTES_CT = '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
ENDNOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"


def endnotes_xml(*notes):
    """notes: list of (id, text); separator entries are always included."""
    inner = (
        '<w:endnote w:type="separator" w:id="0"/>'
        '<w:endnote w:type="continuationSeparator" w:id="1"/>'
        + "".join(
            f'<w:endnote w:id="{i}"><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:endnote>'
            for i, t in notes
        )
    )
    return XML_DECL + f"<w:endnotes {NSDECL}>{inner}</w:endnotes>"


def endnote_body(eid):
    return (
        "<w:p><w:r><w:t>text</w:t></w:r>"
        '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
        f'<w:endnoteReference w:id="{eid}"/></w:r></w:p>'
    )


def test_endnotes_merged_with_offset():
    sub = build_subdoc(
        endnote_body(2),
        files={"word/endnotes.xml": endnotes_xml((2, "sub endnote text"))},
        doc_rels=[("rId11", ENDNOTES_RT, "endnotes.xml")],
        content_types_extra=ENDNOTES_CT,
    )
    # master has its own endnote with id 2
    master = build_doc(
        tp("{{p sub }}"),
        files={"word/endnotes.xml": endnotes_xml((2, "master note"))},
        doc_rels=[("rId3", ENDNOTES_RT, "endnotes.xml")],
        content_types_extra=ENDNOTES_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    en = read_docx_part(data, "word/endnotes.xml")
    assert "master note" in en and "sub endnote text" in en
    doc = read_docx_part(data, "word/document.xml")
    # reference must be remapped (master max id is 2 -> sub note 2 becomes 5)
    assert '<w:endnoteReference w:id="5"/>' in doc


def test_endnotes_copied_when_master_has_none():
    sub = build_subdoc(
        endnote_body(2),
        files={"word/endnotes.xml": endnotes_xml((2, "sub endnote text"))},
        doc_rels=[("rId11", ENDNOTES_RT, "endnotes.xml")],
        content_types_extra=ENDNOTES_CT,
    )
    data = render_with_sub(sub)  # master without an endnotes part

    assert "word/endnotes.xml" in docx_names(data)
    en = read_docx_part(data, "word/endnotes.xml")
    assert "sub endnote text" in en and 'w:id="2"' in en
    doc = read_docx_part(data, "word/document.xml")
    # copied whole: the body reference keeps its id
    assert '<w:endnoteReference w:id="2"/>' in doc
    # relationship + content type registered
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    assert "endnotes.xml" in rels and ENDNOTES_RT in rels
    ct = read_docx_part(data, "[Content_Types].xml")
    assert "endnotes+xml" in ct
