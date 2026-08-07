"""Tests for the python-docx P1 parity round: hyperlinks, iter_inner_content,
contains_page_break, run.add_picture, mark_comment_range, table_direction,
grid_span/grid_cols, nested tables in cells.
"""
import io
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp, p, run, cell, tr, tbl, make_png  # noqa: E402
from test_subdoc_sections_comments import build_doc  # noqa: E402

from docxtplrs import DocxTemplate, Inches  # noqa: E402

HYPERLINK_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def new_tpl(body=None, **kw):
    return DocxTemplate(io.BytesIO(make_docx(body or tp("x"), **kw)))


def saved_xml(tpl, part="word/document.xml"):
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), part)


# ---------------- hyperlinks ----------------

def _hyperlinked_doc():
    body = (
        tp("before")
        + '<w:p><w:r><w:t>see </w:t></w:r>'
        '<w:hyperlink r:id="rId5"><w:r><w:t>example</w:t></w:r></w:hyperlink>'
        '<w:hyperlink w:anchor="_Toc123"><w:r><w:t>internal</w:t></w:r></w:hyperlink>'
        "<w:r><w:t> after</w:t></w:r></w:p>"
    )
    return build_doc(
        body,
        doc_rels=[("rId5", HYPERLINK_RT, "https://example.com")],
    )


def test_paragraph_hyperlinks():
    tpl = DocxTemplate(io.BytesIO(_hyperlinked_doc()))
    para = tpl.paragraphs[1]
    links = para.hyperlinks
    assert len(links) == 2
    assert links[0].text == "example"
    assert links[0].address == "https://example.com"
    assert links[0].fragment == ""
    assert links[1].text == "internal"
    assert links[1].address == ""  # internal jump
    assert links[1].fragment == "_Toc123"
    assert links[0].contains_page_break is False


def test_contains_page_break():
    body = tp("plain") + p(run("a") + '<w:br w:type="page"/>' + run("b")) + p('<w:r><w:t>c</w:t><w:lastRenderedPageBreak/></w:r>')
    tpl = new_tpl(body)
    paras = tpl.paragraphs
    assert paras[0].contains_page_break is False
    # hard break does not count (python-docx semantics)
    assert paras[1].contains_page_break is False
    # rendered break counts
    assert paras[2].contains_page_break is True
    assert paras[2].runs[0].contains_page_break is True


# ---------------- iter_inner_content ----------------

def test_document_iter_inner_content():
    tpl = new_tpl(tp("p0") + tbl([tr(cell(tp("x")), cell(tp("y")))]) + tp("p1"))
    items = tpl.get_docx().iter_inner_content()
    assert [type(i).__name__ for i in items] == ["Paragraph", "Table", "Paragraph"]
    assert items[0].text == "p0" and items[2].text == "p1"
    assert items[1].cell(0, 0).text == "x"


def test_paragraph_iter_inner_content():
    tpl = DocxTemplate(io.BytesIO(_hyperlinked_doc()))
    items = tpl.paragraphs[1].iter_inner_content()
    assert [type(i).__name__ for i in items] == ["Run", "Hyperlink", "Hyperlink", "Run"]
    assert items[1].text == "example"


def test_run_iter_inner_content():
    tpl = new_tpl()
    r = tpl.paragraphs[0].add_run("a\tb\nc")
    items = r.iter_inner_content()
    assert items == ["a\tb\nc"]


def test_cell_iter_inner_content():
    tpl = new_tpl(tbl([tr(cell(tp("c0")), cell(tp("c1")))]))
    c = tpl.tables[0].cell(0, 0)
    t = c.add_table(1, 1)
    t.cell(0, 0).text = "nested"
    items = c.iter_inner_content()
    assert [type(i).__name__ for i in items] == ["CellParagraph", "CellTable", "CellParagraph"]
    assert items[0].text == "c0"
    assert items[1].cell(0, 0).text == "nested"


# ---------------- run.add_picture / mark_comment_range ----------------

def test_run_add_picture():
    tpl = new_tpl()
    r = tpl.paragraphs[0].add_run("x")
    r.add_picture(io.BytesIO(make_png(5, 5)), width=Inches(1).emu)
    xml = saved_xml(tpl)
    assert "<w:drawing>" in xml
    names = read_names(tpl)
    assert any(n.startswith("word/media/") for n in names)


def read_names(tpl):
    import zipfile
    out = io.BytesIO()
    tpl.save(out)
    out.seek(0)
    with zipfile.ZipFile(out) as z:
        return z.namelist()


def test_mark_comment_range():
    body = p(run("first") + run("second"))
    tpl = new_tpl(body)
    runs = tpl.paragraphs[0].runs
    runs[0].mark_comment_range(runs[1], 3)
    xml = saved_xml(tpl)
    assert '<w:commentRangeStart w:id="3"/>' in xml
    assert '<w:commentRangeEnd w:id="3"/>' in xml
    assert '<w:commentReference w:id="3"/>' in xml
    # start before "first", end after "second"
    assert xml.index("commentRangeStart") < xml.index("first")
    assert xml.index("second") < xml.index("commentRangeEnd")


# ---------------- table_direction / grid_span / grid_cols ----------------

def test_table_direction():
    tpl = new_tpl(tbl([tr(cell(tp("a")), cell(tp("b")))]))
    t = tpl.tables[0]
    assert t.table_direction == 0
    t.table_direction = 1
    assert t.table_direction == 1
    assert "<w:bidiVisual/>" in saved_xml(tpl)


def test_cell_grid_span_and_row_grid_cols():
    body = tbl([tr(cell(tp("a"), tcpr='<w:gridSpan w:val="2"/>'), cell(tp("b")))], widths=(2000, 2000))
    tpl = new_tpl(body)
    assert tpl.tables[0].cell(0, 0).grid_span == 2
    row = tpl.tables[0].rows[0]
    assert row.grid_cols_before == 0 and row.grid_cols_after == 0


def test_cell_nested_table():
    tpl = new_tpl(tbl([tr(cell(tp("outer")), cell(tp("x")))]))
    c = tpl.tables[0].cell(0, 0)
    assert len(c.tables) == 0
    t = c.add_table(2, 2)
    assert len(c.tables) == 1
    t.cell(1, 1).text = "deep"
    assert t.cell(1, 1).text == "deep"
    assert len(t.rows) == 2
    assert [cc.text for cc in t.rows[0].cells] == ["", ""]
    xml = saved_xml(tpl)
    assert xml.count("<w:tbl>") == 2
    # cell still ends with a paragraph after the nested table
    assert "<w:tbl>" in xml and "</w:tbl><w:p/>" in xml
