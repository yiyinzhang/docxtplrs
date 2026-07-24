"""Tests for the document object model facade and bound Subdoc building."""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, docx_names, text_of, tp, cell, tr, tbl, make_png

from docxtplrs import DocxTemplate, Inches, Mm


def test_paragraphs_facade():
    body = tp("first") + tp("second")
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    paras = tpl.paragraphs
    assert [p.text for p in paras] == ["first", "second"]
    # same through get_docx()
    doc = tpl.get_docx()
    assert [p.text for p in doc.paragraphs] == ["first", "second"]


def test_paragraph_runs():
    body = tp("hello")
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    p = tpl.paragraphs[0]
    assert len(p.runs) == 1
    assert p.runs[0].text == "hello"
    assert not p.runs[0].bold


def test_tables_facade():
    body = tbl([tr(cell(tp("a1")), cell(tp("a2"))), tr(cell(tp("b1")), cell(tp("b2")))])
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tables = tpl.tables
    assert len(tables) == 1
    t = tables[0]
    assert len(t.rows) == 2
    assert t.rows[0].cells[0].text == "a1"
    assert t.rows[1].cells[1].text == "b2"


def test_core_properties_facade():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    cp = tpl.core_properties
    # no core.xml initially -> empty, then set (creates the part)
    assert cp.author == ""
    cp.author = "Bob"
    cp.title = "My Doc"
    assert cp.author == "Bob"
    assert cp.title == "My Doc"
    tpl.render({})  # properties edits made before render survive
    out = io.BytesIO()
    tpl.save(out)
    core_xml = read_docx_part(out.getvalue(), "docProps/core.xml")
    assert "<dc:creator>Bob</dc:creator>" in core_xml
    assert "<dc:title>My Doc</dc:title>" in core_xml


def test_styles_facade():
    styles = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>'
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"), styles=styles)))
    styles = tpl.styles
    assert any(s.style_id == "Heading1" and s.name == "heading 1" for s in styles)


def test_bound_subdoc_paragraphs():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    sd = tpl.new_subdoc()
    sd.add_paragraph("This is a sub-document")
    p = sd.add_paragraph("It has been ")
    p.add_run("dynamically").bold = True
    p.add_run(" generated")
    sd.add_heading("Heading, level 1", level=1)
    tpl.render({"sub": sd})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    t = text_of(xml)
    assert "This is a sub-document" in t
    assert "It has been dynamically generated" in t
    assert "Heading, level 1" in t
    assert "<w:b/>" in xml


def test_bound_subdoc_picture(tmp_path):
    png = tmp_path / "p.png"
    png.write_bytes(make_png(10, 10))
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    sd = tpl.new_subdoc()
    sd.add_picture(str(png), width=Inches(1.0))
    tpl.render({"sub": sd})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    xml = read_docx_part(data, "word/document.xml")
    assert "<w:drawing>" in xml
    assert 'cx="914400"' in xml
    assert any(n.startswith("word/media/") for n in docx_names(data))


def test_bound_subdoc_table():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    sd = tpl.new_subdoc()
    table = sd.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "Qty"
    hdr[1].text = "Id"
    hdr[2].text = "Desc"
    row = table.add_row().cells
    row[0].text = "1"
    row[1].text = "101"
    row[2].text = "Spam"
    tpl.render({"sub": sd})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    t = text_of(xml)
    assert "Qty" in t and "Spam" in t
    assert xml.count("<w:tr>") == 2
    assert table.cell(1, 2).text == "Spam"


def test_bound_subdoc_not_printed():
    # building a bound subdoc without printing it must not change the output
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("static"))))
    sd = tpl.new_subdoc()
    sd.add_paragraph("invisible")
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    t = text_of(read_docx_part(out.getvalue(), "word/document.xml"))
    assert "invisible" not in t


def test_run_setters():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    sd = tpl.new_subdoc()
    p = sd.add_paragraph()
    r = p.add_run("styled")
    r.italic = True
    r.underline = True
    r.color = "#FF0000"
    r.size = 24
    r.font = "Arial"
    r.highlight = "#00FF00"
    r.subscript = True
    r.text = "styled2"
    assert r.text == "styled2"
    tpl.render({"sub": sd})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "<w:i/>" in xml
    assert '<w:u w:val="single"/>' in xml
    assert '<w:color w:val="FF0000"/>' in xml
    assert '<w:sz w:val="24"/>' in xml
    assert 'w:ascii="Arial"' in xml
    assert '<w:shd w:fill="00FF00"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert "styled2" in text_of(xml)


def test_sections_facade():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sections = tpl.sections
    assert len(sections) == 1  # our fixture has one empty sectPr


def test_inline_shapes_facade(tmp_path):
    png = tmp_path / "p.png"
    png.write_bytes(make_png(10, 10))
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    sd = tpl.new_subdoc()
    sd.add_picture(str(png), width=Mm(10))
    tpl.render({"sub": sd})
    shapes = tpl.inline_shapes
    assert len(shapes) == 1
    assert shapes[0].width.emu == 360000
    assert shapes[0].type == "picture"
