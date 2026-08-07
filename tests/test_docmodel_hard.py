"""Tests for the so-called 'hard boundary' items that turned out feasible:
rendered_page_breaks, Section.iter_inner_content, part facade, and
drawing/rendered-break items in Run.iter_inner_content.
"""
import io
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp, p, run, cell, tr, tbl, make_png  # noqa: E402

from docxtplrs import DocxTemplate, Inches  # noqa: E402


def new_tpl(body=None, **kw):
    return DocxTemplate(io.BytesIO(make_docx(body or tp("x"), **kw)))


def saved_xml(tpl, part="word/document.xml"):
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), part)


# ---------------- rendered_page_breaks ----------------

def test_rendered_page_breaks():
    body = (
        tp("plain")
        + p('<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t>'
            '<w:lastRenderedPageBreak/></w:r>')
    )
    tpl = new_tpl(body)
    p0, p1 = tpl.paragraphs
    assert len(p0.rendered_page_breaks) == 0
    assert len(p1.rendered_page_breaks) == 2
    assert p1.contains_page_break is True


# ---------------- Section.iter_inner_content ----------------

def test_section_iter_inner_content():
    sect1 = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>'
    body = tp("s1p0") + tp("s1p1") + sect1 + tp("s2p0") + tbl([tr(cell(tp("t")), cell(tp("u")))])
    tpl = new_tpl(body)
    doc = tpl.get_docx()
    assert len(doc.sections) == 2
    s1 = doc.sections[0].iter_inner_content()
    assert [type(i).__name__ for i in s1] == ["Paragraph", "Paragraph", "Paragraph"]
    assert [i.text for i in s1] == ["s1p0", "s1p1", ""]
    s2 = doc.sections[1].iter_inner_content()
    assert [type(i).__name__ for i in s2] == ["Paragraph", "Table"]
    assert s2[0].text == "s2p0"
    assert s2[1].cell(0, 0).text == "t"


# ---------------- part facade ----------------

def test_part_facade():
    tpl = new_tpl()
    part = tpl.get_docx().part
    assert part.partname == "/word/document.xml"
    assert b"<w:document" in part.blob
    # proxies on other objects point at the same part
    assert tpl.paragraphs[0].part.partname == "/word/document.xml"
    assert tpl.paragraphs[0].runs[0].part.partname == "/word/document.xml"
    styles = '<w:style w:type="paragraph" w:styleId="N"><w:name w:val="N"/></w:style>'
    tpl2 = new_tpl(styles=styles)
    assert tpl2.styles["N"].part.partname == "/word/styles.xml"


def test_part_rels_and_content_type():
    footnotes = '<w:footnote w:id="2"><w:p><w:r><w:t>fn</w:t></w:r></w:p></w:footnote>'
    tpl = new_tpl(footnotes=footnotes)
    part = tpl.get_docx().part
    by_id = {r["rId"]: r for r in part.rels}
    assert by_id["rId1"]["type"].endswith("/relationships/footnotes")
    assert by_id["rId1"]["target"] == "footnotes.xml"
    assert by_id["rId1"]["is_external"] is False
    # Override-based content type
    assert part.content_type.endswith("wordprocessingml.document.main+xml")

    styles = '<w:style w:type="paragraph" w:styleId="N"><w:name w:val="N"/></w:style>'
    tpl2 = new_tpl(styles=styles)
    spart = tpl2.styles["N"].part
    assert spart.content_type.endswith("wordprocessingml.styles+xml")
    # no word/_rels/styles.xml.rels -> empty list
    assert spart.rels == []


# ---------------- Run.iter_inner_content with drawings ----------------

def test_run_iter_inner_content_with_drawing_and_break():
    tpl = new_tpl()
    r = tpl.paragraphs[0].add_run("before")
    r.add_picture(io.BytesIO(make_png(4, 4)), width=Inches(1).emu)
    r.add_text("after")
    items = r.iter_inner_content()
    kinds = [type(i).__name__ for i in items]
    assert kinds == ["str", "XmlElement", "str"]
    assert items[0] == "before" and items[2] == "after"
    # the drawing is a live xml proxy
    assert items[1].tag == "w:drawing"
    assert "<w:drawing>" in items[1].xml


def test_run_iter_inner_content_with_rendered_break():
    tpl = new_tpl(p('<w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r>'))
    items = tpl.paragraphs[0].runs[0].iter_inner_content()
    kinds = [type(i).__name__ for i in items]
    assert kinds == ["str", "RenderedPageBreak", "str"]
    assert items[0] == "a" and items[2] == "b"
