"""Tests for the writable document model: facade write-back, sections,
styles, settings, comments, jinja2.utils."""

import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, docx_names, text_of, tp, cell, tr, tbl, make_png

from docxtplrs import DocxTemplate, Cycler, Joiner, generate_lorem_ipsum, Pt


def saved_xml(tpl):
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


# ---------------- facade write-back ----------------

def test_paragraph_writeback():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hello") + tp("world"))))
    p = tpl.paragraphs[1]
    assert p.text == "world"
    p.text = "changed"
    assert tpl.paragraphs[1].text == "changed"
    tpl.render({})
    assert "changed" in text_of(saved_xml(tpl))
    assert "world" not in text_of(saved_xml(tpl))


def test_paragraph_style_writeback():
    styles = '<w:style w:type="paragraph" w:styleId="MyStyle"><w:name w:val="my style"/></w:style>'
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"), styles=styles)))
    p = tpl.paragraphs[0]
    p.style = "my style"  # by name
    assert p.style == "MyStyle"


def test_run_writeback():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("plain"))))
    r = tpl.paragraphs[0].runs[0]
    assert not r.bold
    r.bold = True
    assert r.bold is True
    r.text = "strong"
    assert r.text == "strong"
    r.color = "#00FF00"
    tpl.render({})
    xml = saved_xml(tpl)
    assert "<w:b/>" in xml and "strong" in text_of(xml) and 'w:val="00FF00"' in xml


def test_paragraph_add_run_writeback():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("a"))))
    p = tpl.paragraphs[0]
    r = p.add_run(" tail")
    r.italic = True
    tpl.render({})
    xml = saved_xml(tpl)
    assert "a tail" in text_of(xml)
    assert "<w:i/>" in xml


def test_table_writeback():
    body = tbl([tr(cell(tp("a1")), cell(tp("a2")))])
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    t = tpl.tables[0]
    t.rows[0].cells[1].text = "B2"
    assert t.cell(0, 1).text == "B2"
    row = t.add_row().cells
    row[0].text = "c1"
    tpl.render({})
    xml = saved_xml(tpl)
    assert "B2" in text_of(xml) and "c1" in text_of(xml)
    assert xml.count("<w:tr>") == 2


# ---------------- sections ----------------

def test_section_margins():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    s = tpl.sections[0]
    from docxtplrs import Inches

    s.page_width = Inches(8.5)
    s.page_height = Inches(11)
    s.left_margin = Inches(1)
    assert s.page_width.emu == Inches(8.5).emu
    tpl.render({})
    xml = saved_xml(tpl)
    assert 'w:w="12240"' in xml  # 8.5in = 12240 twips


def test_section_orientation():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    s = tpl.sections[0]
    from docxtplrs import Inches

    s.page_width = Inches(8.5)
    s.page_height = Inches(11)
    s.orientation = "landscape"
    assert s.orientation == "landscape"
    tpl.render({})
    xml = saved_xml(tpl)
    assert 'w:orient="landscape"' in xml
    # dimensions swapped
    assert 'w:w="15840"' in xml and 'w:h="12240"' in xml


def test_add_section():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    doc = tpl.get_docx()
    assert len(doc.sections) == 1
    doc.add_section(2)  # NEW_PAGE
    assert len(doc.sections) == 2
    tpl.render({})
    xml = saved_xml(tpl)
    assert xml.count("<w:sectPr") == 2
    assert 'w:type w:val="nextPage"' in xml


def test_section_header_footer():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("body"))))
    s = tpl.sections[0]
    assert s.header.is_linked_to_previous
    s.header.is_linked_to_previous = False
    assert not s.header.is_linked_to_previous
    s.header.add_paragraph("My Header")
    assert "My Header" in s.header.paragraphs
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    import zipfile

    with zipfile.ZipFile(out) as z:
        hdr = z.read("word/header1.xml").decode()
    assert "My Header" in hdr


def test_different_first_page():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    s = tpl.sections[0]
    assert not s.different_first_page_header_footer
    s.different_first_page_header_footer = True
    tpl.render({})
    assert "<w:titlePg/>" in saved_xml(tpl)


# ---------------- styles ----------------

def test_add_style_and_font():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    styles = tpl.styles
    st = styles.add_style("My Char Style", 2)
    assert st.name == "My Char Style"
    assert st.style_type == "character"
    st.font.bold = True
    st.font.size = Pt(14)  # 28 half-points
    st.font.color.rgb = "#FF0000"
    st.font.name = "Courier New"
    assert st.font.bold is True
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    sx = read_docx_part(out.getvalue(), "word/styles.xml")
    assert 'w:val="My Char Style"' in sx
    assert "<w:b/>" in sx and 'w:val="28"' in sx and 'w:val="FF0000"' in sx
    assert 'w:ascii="Courier New"' in sx


def test_styles_getitem_and_delete():
    styles_xml = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>'
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"), styles=styles_xml)))
    st = tpl.styles["heading 1"]
    assert st.style_id == "Heading1"
    st.delete()
    with pytest.raises(Exception):
        tpl.styles["heading 1"]


def test_styles_iteration():
    styles_xml = '<w:style w:type="paragraph" w:styleId="S1"><w:name w:val="s1"/></w:style>'
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"), styles=styles_xml)))
    assert [s.style_id for s in tpl.styles] == ["S1"]
    assert len(tpl.styles) == 1


# ---------------- settings ----------------

def test_settings_odd_even():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    settings = tpl.settings
    assert not settings.odd_and_even_pages_header_footer
    settings.odd_and_even_pages_header_footer = True
    assert settings.odd_and_even_pages_header_footer
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    assert "<w:evenAndOddHeaders/>" in read_docx_part(out.getvalue(), "word/settings.xml")


# ---------------- comments ----------------

def test_add_comment_anchored():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("important text"))))
    doc = tpl.get_docx()
    run = doc.paragraphs[0].runs[0]
    c = doc.add_comment(run, text="a note", author="Bob", initials="BB")
    assert c.text == "a note"
    assert c.author == "Bob"
    assert c.initials == "BB"
    assert len(doc.comments) == 1
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    docxml = read_docx_part(data, "word/document.xml")
    assert "<w:commentRangeStart" in docxml
    assert "<w:commentReference" in docxml
    comments = read_docx_part(data, "word/comments.xml")
    assert "a note" in comments and 'w:author="Bob"' in comments


def test_comments_collection_add():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    c = tpl.comments.add_comment(text="standalone", author="A")
    assert c.comment_id == 0
    c2 = tpl.comments.add_comment(text="second")
    assert c2.comment_id == 1
    assert len(tpl.comments) == 2
    texts = [cm.text for cm in tpl.comments]
    assert texts == ["standalone", "second"]


# ---------------- jinja2.utils ----------------

def test_cycler():
    c = Cycler(1, 2, 3)
    assert c.next() == 1
    assert c.next() == 2
    assert c.current == 2
    assert c.next() == 3
    assert c.next() == 1
    c.reset()
    assert c.next() == 1


def test_joiner():
    j = Joiner()
    assert j("a") == "a"
    assert j("b") == ", b"
    assert j("c", "d") == ", c, d"


def test_lipsum():
    text = generate_lorem_ipsum(2, html=False)
    assert len(text) > 50
    html = generate_lorem_ipsum(1)
    assert "<p>" in html
