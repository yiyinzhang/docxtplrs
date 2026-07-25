"""Full-matrix round-trip tests for the docmodel facade (docmodel.rs coverage):
sections, styles + fonts, core properties, paragraphs/runs, tables, settings.
"""
import io
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp  # noqa: E402

from docxtplrs import DocxTemplate, Inches, Cm  # noqa: E402


def new_tpl(body=None, **kw):
    return DocxTemplate(io.BytesIO(make_docx(body or tp("x"), **kw)))


def saved_xml(tpl, part="word/document.xml"):
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), part)


# ---------------- sections ----------------

def test_section_all_margins_roundtrip():
    tpl = new_tpl()
    s = tpl.sections[0]
    s.top_margin = Inches(1.25)
    s.bottom_margin = Inches(0.75)
    s.left_margin = Inches(1.5)
    s.right_margin = Cm(2)
    assert s.top_margin.emu == Inches(1.25).emu
    assert s.bottom_margin.emu == Inches(0.75).emu
    assert s.left_margin.emu == Inches(1.5).emu
    # twips round-trip is lossy: 2cm -> 1134 twips -> 720090 emu
    assert abs(s.right_margin.emu - Cm(2).emu) < 1000
    tpl.render({})
    xml = saved_xml(tpl)
    assert 'w:top="1800"' in xml      # 1.25in
    assert 'w:bottom="1080"' in xml   # 0.75in
    assert 'w:left="2160"' in xml     # 1.5in
    assert 'w:right="1133"' in xml    # 2cm (720000emu/635 rounded down)


def test_section_page_size_roundtrip():
    tpl = new_tpl()
    s = tpl.sections[0]
    s.page_width = Inches(8.5)
    s.page_height = Inches(14)
    assert s.page_width.emu == Inches(8.5).emu
    assert s.page_height.emu == Inches(14).emu
    tpl.render({})
    xml = saved_xml(tpl)
    assert 'w:w="12240"' in xml and 'w:h="20160"' in xml


def test_section_orientation_landscape_swaps():
    tpl = new_tpl()
    s = tpl.sections[0]
    s.page_width = Inches(8.5)
    s.page_height = Inches(11)
    s.orientation = "landscape"
    assert s.orientation == "landscape"
    tpl.render({})
    xml = saved_xml(tpl)
    assert 'w:orient="landscape"' in xml
    # python-docx swaps w/h on orientation change
    assert 'w:w="15840"' in xml and 'w:h="12240"' in xml


def test_section_different_first_page_roundtrip():
    tpl = new_tpl()
    s = tpl.sections[0]
    assert s.different_first_page_header_footer is False
    s.different_first_page_header_footer = True
    assert s.different_first_page_header_footer is True
    tpl.render({})
    assert "w:titlePg" in saved_xml(tpl)


def test_header_footer_link_to_previous():
    tpl = new_tpl()
    s = tpl.sections[0]
    assert s.header.is_linked_to_previous is True
    s.header.is_linked_to_previous = False
    assert s.header.is_linked_to_previous is False
    s.footer.is_linked_to_previous = False
    assert s.footer.is_linked_to_previous is False
    s.header.is_linked_to_previous = True
    assert s.header.is_linked_to_previous is True


# ---------------- styles & fonts ----------------

def test_style_name_and_base_style():
    styles_xml = (
        '<w:style w:type="paragraph" w:styleId="Base">'
        '<w:name w:val="Base"/></w:style>'
    )
    tpl = new_tpl(styles=styles_xml)
    st = tpl.styles.add_style("Derived", 1)
    assert st.name == "Derived"
    st.name = "Derived2"
    assert st.name == "Derived2"
    st.base_style = "Base"
    assert st.base_style == "Base"
    tpl.render({})
    sx = saved_xml(tpl, "word/styles.xml")
    assert 'w:val="Derived2"' in sx and 'w:basedOn w:val="Base"' in sx


def test_style_font_full_matrix():
    tpl = new_tpl()
    st = tpl.styles.add_style("F", 2)
    f = st.font
    f.bold = True
    f.italic = True
    f.underline = True
    f.strike = True
    f.small_caps = True
    f.all_caps = True
    f.color = "#123456"
    f.size = 22
    f.name = "Consolas"
    assert f.bold is True and f.italic is True
    assert f.strike is True and f.small_caps is True and f.all_caps is True
    assert f.color == "123456"  # getter strips the leading '#'
    assert f.size == 22
    assert f.name == "Consolas"
    tpl.render({})
    sx = saved_xml(tpl, "word/styles.xml")
    for frag in (
        "<w:b/>", "<w:i/>", "<w:strike/>",
        "<w:smallCaps/>", "<w:caps/>",
        'w:val="123456"', 'w:val="22"', 'w:ascii="Consolas"',
    ):
        assert frag in sx, frag


# ---------------- core properties ----------------

def test_core_properties_all_fields():
    tpl = new_tpl()
    cp = tpl.core_properties
    vals = {
        "author": "Bob",
        "category": "reports",
        "comments": "note",
        "content_status": "draft",
        "created": "2024-01-02T03:04:05Z",
        "identifier": "id-42",
        "keywords": "a;b",
        "language": "zh-CN",
        "last_modified_by": "Alice",
        "modified": "2024-02-03T04:05:06Z",
        "revision": "7",
        "subject": "topic",
        "title": "Title",
    }
    for k, v in vals.items():
        setattr(cp, k, v)
    for k, v in vals.items():
        assert getattr(cp, k) == v, k
    tpl.render({})
    core_xml = saved_xml(tpl, "docProps/core.xml")
    assert "<dc:creator>Bob</dc:creator>" in core_xml
    assert "<dc:title>Title</dc:title>" in core_xml
    assert "id-42" in core_xml


# ---------------- paragraphs & runs ----------------

def test_paragraph_style_and_add_run():
    styles_xml = (
        '<w:style w:type="paragraph" w:styleId="MyP"><w:name w:val="MyP"/></w:style>'
    )
    tpl = new_tpl(styles=styles_xml)
    p = tpl.add_paragraph("hello")
    assert p.text == "hello"
    p.style = "MyP"
    assert p.style == "MyP"
    r = p.add_run(" world")
    r.bold = True
    assert p.text == "hello world"
    assert len(p.runs) == 2
    tpl.render({})
    xml = saved_xml(tpl)
    assert 'w:val="MyP"' in xml and "<w:b/>" in xml


def test_run_rpr_matrix():
    tpl = new_tpl()
    p = tpl.add_paragraph()
    r = p.add_run("x")
    r.bold = True
    r.italic = True
    r.underline = True
    r.strike = True
    r.subscript = True
    r.color = "#A0B0C0"
    r.highlight = "#00FF00"
    r.font = "Courier New"
    r.size = 18
    assert r.bold is True and r.italic is True
    assert r.underline == "single"
    assert r.strike is True
    assert r.subscript is True
    assert r.superscript is False  # vertAlign present, value is subscript
    assert r.color == "A0B0C0"  # getter strips the leading '#'
    assert r.font_name == "Courier New"
    assert r.size == 18
    tpl.render({})
    xml = saved_xml(tpl)
    for frag in ("<w:b/>", "<w:i/>", "<w:strike/>", "<w:vertAlign"):
        assert frag in xml, frag


# ---------------- tables ----------------

def test_table_rows_cells_add_row():
    tpl = new_tpl()
    t = tpl.add_table(2, 2)
    assert len(t.rows) == 2
    t.cell(0, 0).text = "a"
    t.cell(0, 1).text = "b"
    row = t.add_row()
    assert len(t.rows) == 3
    row.cells[0].text = "c"
    assert t.cell(0, 0).text == "a"
    assert t.cell(1, 0).text == ""
    assert len(t.rows[0].cells) == 2
    tpl.render({})
    xml = saved_xml(tpl)
    assert xml.count("<w:tr>") == 3
    assert ">a</w:t>" in xml and ">c</w:t>" in xml


# ---------------- settings ----------------

def test_settings_odd_even_roundtrip():
    tpl = new_tpl()
    st = tpl.settings
    assert st.odd_and_even_pages_header_footer is False
    st.odd_and_even_pages_header_footer = True
    assert st.odd_and_even_pages_header_footer is True
    tpl.render({})
    sx = saved_xml(tpl, "word/settings.xml")
    assert "w:evenAndOddHeaders" in sx


# ---------------- inline shapes ----------------

def test_inline_shapes_lists_picture():
    tpl = new_tpl()
    tpl.add_picture(io.BytesIO(_png()))
    shapes = tpl.inline_shapes
    assert len(shapes) == 1
    assert shapes[0].width.emu > 0 and shapes[0].height.emu > 0
    assert shapes[0].type == "picture"


def _png():
    import struct
    import zlib

    def chunk(t, d):
        return (
            struct.pack(">I", len(d)) + t + d
            + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", 4, 3, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * 4 for _ in range(3))
    return (
        b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    )
