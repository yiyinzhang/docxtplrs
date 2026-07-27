"""Tests for the API-parity additions: XmlElement escape hatch
(settings/document/styles.element), get_file_crc, HEADER_URI/FOOTER_URI,
get_headers_footers, Table.style, Cell.merge, Section header/footer variants."""

import binascii
import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, docx_names, text_of, tp, cell, tr, tbl

from docxtplrs import DocxTemplate

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def saved_part(tpl, name):
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), name)


# ---------------- XmlElement: settings.element ----------------

def test_settings_element_read_write():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    el = tpl.settings.element
    assert el.tag == "w:settings"
    el.append(f'<w:zoom {WNS} w:percent="120"/>')
    zoom = el.find("w:zoom")
    assert zoom is not None
    assert zoom.get("w:percent") == "120"
    zoom.set("w:percent", "150")
    assert el.find("w:zoom").get("w:percent") == "150"
    tpl.render({})
    xml = saved_part(tpl, "word/settings.xml")
    assert 'w:zoom' in xml and 'w:percent="150"' in xml


def test_settings_element_creates_part():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    _ = tpl.settings.element  # part created on demand
    out = io.BytesIO()
    tpl.save(out)
    assert "word/settings.xml" in docx_names(out.getvalue())


def test_settings_element_remove_and_attrib():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    el = tpl.settings.element
    el.append(f'<w:zoom {WNS} w:percent="100"/>')
    zoom = el.find("w:zoom")
    assert zoom.attrib["w:percent"] == "100"
    el.remove(zoom)
    assert el.find("w:zoom") is None
    el.set("w:custom", "1")
    assert el.get("w:custom") == "1"
    el.remove_attr("w:custom")
    assert el.get("w:custom") is None


def test_settings_element_odd_even_still_works():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.settings.odd_and_even_pages_header_footer = True
    el = tpl.settings.element
    assert el.find("w:evenAndOddHeaders") is not None
    tpl.settings.odd_and_even_pages_header_footer = False
    assert tpl.settings.element.find("w:evenAndOddHeaders") is None


# ---------------- XmlElement: document.element & styles.element ----------------

def test_document_element():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hello"))))
    el = tpl.element  # delegated to Document.element
    assert el.tag == "w:document"
    body = el.find("w:body")
    assert body is not None
    tags = [c.tag for c in body.children]
    assert "w:p" in tags and "w:sectPr" in tags
    p = body.find("w:p")
    assert p.text == "hello"
    assert len(body) >= 2


def test_document_element_mutation():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hello"))))
    body = tpl.get_docx().element.find("w:body")
    sectpr = body.find("w:sectPr")
    body.insert(0, f'<w:p {WNS}><w:r><w:t>injected</w:t></w:r></w:p>')
    assert body.find("w:p").text == "injected"
    tpl.render({})
    xml = saved_part(tpl, "word/document.xml")
    assert "injected" in text_of(xml)


def test_document_element_append_xmlelement():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("a") + tp("b"))))
    body = tpl.element.find("w:body")
    first_p = body.find("w:p")
    body.append(first_p)  # deep-copies the element
    ps = body.findall("w:p")
    assert len(ps) == 3
    assert ps[-1].text == "a"


def test_styles_element():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    el = tpl.styles.element
    assert el.tag == "w:styles"
    ids = [s.get("w:styleId") for s in el.findall("w:style")]
    assert "Normal" in ids
    el.append(
        f'<w:style {WNS} w:type="paragraph" w:styleId="ZZ"><w:name w:val="zz"/></w:style>'
    )
    assert "ZZ" in [s.style_id for s in tpl.styles]
    tpl.render({})
    assert 'w:styleId="ZZ"' in saved_part(tpl, "word/styles.xml")


def test_xmlelement_str_and_errors():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    el = tpl.settings.element
    el.append(f'<w:zoom {WNS} w:percent="90"/>')
    assert 'w:percent="90"' in str(el)
    with pytest.raises(ValueError):
        el.append("<not-closed>")
    with pytest.raises(ValueError):
        # not a direct child (different part)
        el.remove(tpl.element.find("w:body"))


# ---------------- get_file_crc / URIs / get_headers_footers ----------------

def test_header_footer_uri_constants():
    assert DocxTemplate.HEADER_URI.endswith("relationships/header")
    assert DocxTemplate.FOOTER_URI.endswith("relationships/footer")


def test_get_file_crc():
    data = b"\x89PNG\r\n\x1a\n" + bytes(range(256))
    expected = binascii.crc32(data) & 0xFFFFFFFF
    assert DocxTemplate.get_file_crc(data) == expected
    assert DocxTemplate.get_file_crc(io.BytesIO(data)) == expected


def test_get_file_crc_path(tmp_path):
    p = tmp_path / "blob.bin"
    p.write_bytes(b"hello world")
    assert DocxTemplate.get_file_crc(str(p)) == binascii.crc32(b"hello world") & 0xFFFFFFFF


def test_get_headers_footers():
    tpl = DocxTemplate(
        io.BytesIO(
            make_docx(
                tp("x"),
                headers={"header1.xml": tp("hdr text")},
                footers={"footer1.xml": tp("ftr text")},
            )
        )
    )
    hdrs = tpl.get_headers_footers(DocxTemplate.HEADER_URI)
    ftrs = tpl.get_headers_footers(DocxTemplate.FOOTER_URI)
    assert len(hdrs) == 1 and len(ftrs) == 1
    rid, el = hdrs[0]
    assert isinstance(rid, str) and rid
    assert el.tag == "w:hdr"
    assert el.find("w:p").text == "hdr text"
    assert ftrs[0][1].find("w:p").text == "ftr text"
    # the element proxy is live
    el.find("w:p").append(f'<w:r {WNS}><w:t>+added</w:t></w:r>')
    tpl.render({})
    assert "added" in text_of(saved_part(tpl, "word/header1.xml"))


def test_get_headers_footers_empty():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    assert tpl.get_headers_footers(DocxTemplate.HEADER_URI) == []


# ---------------- Table.style ----------------

def test_table_style():
    styles = '<w:style w:type="table" w:styleId="MyTbl"><w:name w:val="my table"/></w:style>'
    tpl = DocxTemplate(io.BytesIO(make_docx(tbl([tr(cell(tp("a")))], widths=(2000,)), styles=styles)))
    t = tpl.tables[0]
    assert t.style is None
    t.style = "my table"  # by name
    assert t.style == "MyTbl"
    tpl.render({})
    xml = saved_part(tpl, "word/document.xml")
    assert '<w:tblStyle w:val="MyTbl"/>' in xml


# ---------------- Cell.merge ----------------

def test_cell_merge_horizontal():
    tpl = DocxTemplate(io.BytesIO(make_docx(
        tbl([tr(cell(tp("L")), cell(tp("M")), cell(tp("R")))], widths=(1000, 1000, 1000))
    )))
    t = tpl.tables[0]
    m = t.cell(0, 0).merge(t.cell(0, 2))
    assert m.text == "LMR"
    tpl.render({})
    xml = saved_part(tpl, "word/document.xml")
    assert 'w:gridSpan w:val="3"' in xml
    assert xml.count("<w:tc>") == 1


def test_cell_merge_vertical():
    tpl = DocxTemplate(io.BytesIO(make_docx(
        tbl([tr(cell(tp("top"))), tr(cell(tp("mid"))), tr(cell(tp("bot")))], widths=(1000,))
    )))
    t = tpl.tables[0]
    m = t.cell(0, 0).merge(t.cell(1, 0))
    assert m.text == "topmid"
    assert t.cell(2, 0).text == "bot"  # untouched
    tpl.render({})
    xml = saved_part(tpl, "word/document.xml")
    assert '<w:vMerge w:val="restart"/>' in xml
    assert "<w:vMerge/>" in xml
    assert xml.count("<w:tr>") == 3  # rows are kept for vertical merges


def test_cell_merge_rectangular():
    tpl = DocxTemplate(io.BytesIO(make_docx(
        tbl([
            tr(cell(tp("A")), cell(tp("B"))),
            tr(cell(tp("C")), cell(tp("D"))),
        ], widths=(1000, 1000))
    )))
    t = tpl.tables[0]
    m = t.cell(0, 0).merge(t.cell(1, 1))
    assert m.text == "ABCD"
    tpl.render({})
    xml = saved_part(tpl, "word/document.xml")
    assert 'w:gridSpan w:val="2"' in xml
    assert '<w:vMerge w:val="restart"/>' in xml
    assert "<w:vMerge/>" in xml


def test_cell_merge_different_tables_rejected():
    tpl = DocxTemplate(io.BytesIO(make_docx(
        tbl([tr(cell(tp("a")))], widths=(1000,)) + tbl([tr(cell(tp("b")))], widths=(1000,))
    )))
    with pytest.raises(ValueError):
        tpl.tables[0].cell(0, 0).merge(tpl.tables[1].cell(0, 0))


# ---------------- Section header/footer variants ----------------

def test_section_first_page_header_footer():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sec = tpl.sections[0]
    assert sec.first_page_header.is_linked_to_previous
    sec.different_first_page_header_footer = True
    sec.first_page_header.add_paragraph("first hdr")
    sec.first_page_footer.add_paragraph("first ftr")
    assert sec.first_page_header.paragraphs == ["first hdr"]
    assert sec.first_page_footer.paragraphs == ["first ftr"]
    # default header unaffected
    assert sec.header.is_linked_to_previous
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    names = docx_names(out.getvalue())
    assert any("header" in n for n in names) and any("footer" in n for n in names)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert '<w:headerReference w:type="first"' in xml
    assert '<w:footerReference w:type="first"' in xml


def test_section_even_page_header_footer():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sec = tpl.sections[0]
    tpl.settings.odd_and_even_pages_header_footer = True
    sec.even_page_header.add_paragraph("even hdr")
    sec.even_page_footer.add_paragraph("even ftr")
    assert sec.even_page_header.paragraphs == ["even hdr"]
    assert sec.even_page_footer.paragraphs == ["even ftr"]
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert '<w:headerReference w:type="even"' in xml
    assert '<w:footerReference w:type="even"' in xml


def test_section_variants_unlink():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sec = tpl.sections[0]
    sec.first_page_header.add_paragraph("first hdr")
    assert not sec.first_page_header.is_linked_to_previous
    sec.first_page_header.is_linked_to_previous = True
    assert sec.first_page_header.is_linked_to_previous
    xml = tpl.get_xml()
    assert '<w:headerReference w:type="first"' not in xml
