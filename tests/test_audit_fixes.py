"""Tests for the audit round: printf operator, debug, striptags,
Document facade mutation, __getattr__ delegation, docx_context."""

import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, tp, make_png

from docxtplrs import DocxTemplate, TemplateError, Mm


def render_xml(body, context, **kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body, **kw)))
    tpl.render(context)
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


def test_printf_operator():
    body = tp("{{ '%s-%d' % ('a', 2) }}|{{ '%05.1f' % (3.14159,) }}|{{ '%s' % name }}")
    t = text_of(render_xml(body, {"name": "n"}))
    assert "a-2|003.1|n" in t


def test_modulo_still_works():
    assert "1" in text_of(render_xml(tp("{{ 7 % 3 }}"), {}))


def test_printf_ignores_strings_with_percent():
    t = text_of(render_xml(tp("{{ '100% sure' }}"), {}))
    assert "100% sure" in t


def test_debug_statement():
    xml = render_xml(tp("{% set x = 1 %}{% debug %}"), {})
    assert "x" in xml


def test_striptags_filter():
    t = text_of(render_xml(tp("{{ '<b>x</b> <i>y</i>'|striptags }}"), {}))
    assert "x y" in t


def test_doc_add_paragraph_and_run():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("first"))))
    doc = tpl.get_docx()
    p = doc.add_paragraph("hello ")
    assert p.text == "hello "
    r = p.add_run("world")
    r.bold = True
    r.italic = True
    r.color = "#FF0000"
    r.size = 24
    r.font = "Arial"
    r.underline = True
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "hello world" in text_of(xml)
    for frag in ["<w:b/>", "<w:i/>", 'w:val="FF0000"', 'w:val="24"', 'w:ascii="Arial"', '<w:u w:val="single"/>']:
        assert frag in xml


def test_doc_add_heading_pagebreak():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    doc = tpl.get_docx()
    doc.add_heading("Title", level=1)
    doc.add_page_break()
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "Title" in text_of(xml)
    assert '<w:br w:type="page"/>' in xml
    assert '<w:pStyle w:val="Heading 1"/>' in xml


def test_doc_add_picture(tmp_path):
    png = tmp_path / "p.png"
    png.write_bytes(make_png(10, 10))
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.get_docx().add_picture(str(png), width=Mm(10))
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "<w:drawing>" in xml and 'cx="360000"' in xml


def test_doc_add_table_and_cells():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tbl = tpl.get_docx().add_table(rows=1, cols=3)
    hdr = tbl.rows[0].cells
    hdr[0].text = "Qty"
    hdr[2].text = "Desc"
    row = tbl.add_row().cells
    row[1].text = "42"
    assert tbl.cell(1, 1).text == "42"
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    t = text_of(xml)
    assert "Qty" in t and "42" in t and "Desc" in t
    assert xml.count("<w:tr>") == 2


def test_getattr_delegation():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("one") + tp("two"))))
    # paragraphs is a real property; anything else delegates to the facade
    assert [p.text for p in tpl.paragraphs] == ["one", "two"]
    assert len(tpl.tables) == 0
    with pytest.raises(AttributeError):
        tpl.definitely_not_a_document_attribute


def test_template_error_docx_context():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("start {% bad %} end"))))
    with pytest.raises(TemplateError) as exc_info:
        tpl.render({})
    ctx = exc_info.value.docx_context
    assert isinstance(ctx, list)
    assert any("bad" in line for line in ctx)


def test_subdoc_ole_relid_remapped():
    # OLE object references (o:relid) must be remapped like r:id
    import zipfile
    from helpers import document_xml, XML_DECL, NSDECL

    sub_body = (
        '<w:p><w:r><w:object><o:OLEObject Type="Embed" o:relid="rId9" ProgID="Excel.Sheet.12"/>'
        "</w:object></w:r></w:p>"
    )
    buf = io.BytesIO()
    z = zipfile.ZipFile(buf, "w")
    z.writestr(
        "[Content_Types].xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>",
    )
    z.writestr("word/document.xml", document_xml(sub_body))
    z.writestr(
        "word/_rels/document.xml.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/s1.xlsx"/>'
        "</Relationships>",
    )
    z.writestr("word/embeddings/s1.xlsx", b"FAKE-XLSX")
    z.close()

    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(buf.getvalue()))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    doc = read_docx_part(data, "word/document.xml")
    assert 'o:relid="rId9"' not in doc
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    assert "oleObject" in rels
    with zipfile.ZipFile(io.BytesIO(data)) as z2:
        assert z2.read("word/embeddings/s1.xlsx") == b"FAKE-XLSX"
