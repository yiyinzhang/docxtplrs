"""Logical grid (gridSpan/vMerge) table access, python-docx semantics:
table.cell(i,j) / row.cells / column_cells / row_cells resolve merged
coordinates to the merged origin cell."""
import io
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, tp, p, run, cell, tr, tbl  # noqa: E402

from docxtplrs import DocxTemplate  # noqa: E402


def new_tpl(body, **kw):
    return DocxTemplate(io.BytesIO(make_docx(body, **kw)))


# 3-col table; first row merged across all three columns
HMERGED = tbl(
    [tr(cell(tp("merged"), tcpr='<w:gridSpan w:val="3"/>')),
     tr(cell(tp("a")), cell(tp("b")), cell(tp("c")))],
    widths=(1000, 1000, 1000),
)

# 2 rows; first column vertically merged
VMERGED = tbl(
    [tr(cell(tp("top"), tcpr='<w:vMerge w:val="restart"/>'), cell(tp("a"))),
     tr(cell(tp(""), tcpr='<w:vMerge/>'), cell(tp("b")))],
    widths=(1000, 1000),
)


def test_gridspan_cell_access():
    t = new_tpl(HMERGED).tables[0]
    # all three logical columns of row 0 resolve to the merged origin
    assert t.cell(0, 0).text == "merged"
    assert t.cell(0, 1).text == "merged"
    assert t.cell(0, 2).text == "merged"
    # row 1 is normal
    assert [t.cell(1, j).text for j in range(3)] == ["a", "b", "c"]


def test_gridspan_row_cells_expansion():
    t = new_tpl(HMERGED).tables[0]
    # python-docx row.cells repeats the merged cell per covered column
    assert [c.text for c in t.rows[0].cells] == ["merged", "merged", "merged"]
    assert [c.text for c in t.rows[1].cells] == ["a", "b", "c"]


def test_vmerge_cell_access():
    t = new_tpl(VMERGED).tables[0]
    # continuation cell maps to the restart origin
    assert t.cell(0, 0).text == "top"
    assert t.cell(1, 0).text == "top"
    assert t.cell(1, 1).text == "b"


def test_column_cells_logical():
    t = new_tpl(HMERGED).tables[0]
    assert [c.text for c in t.column_cells(0)] == ["merged", "a"]
    assert [c.text for c in t.column_cells(2)] == ["merged", "c"]
    t2 = new_tpl(VMERGED).tables[0]
    assert [c.text for c in t2.column_cells(0)] == ["top", "top"]


def test_row_cells_logical():
    t = new_tpl(VMERGED).tables[0]
    assert [c.text for c in t.row_cells(1)] == ["top", "b"]


def test_cell_out_of_range_raises():
    t = new_tpl(HMERGED).tables[0]
    with pytest.raises(Exception):
        t.cell(5, 0)
    with pytest.raises(Exception):
        t.cell(0, 5)


def test_unmerged_table_unchanged():
    t = new_tpl(tbl([tr(cell(tp("a")), cell(tp("b"))), tr(cell(tp("c")), cell(tp("d")))])).tables[0]
    assert [t.cell(i, j).text for i in range(2) for j in range(2)] == ["a", "b", "c", "d"]
    assert [c.text for c in t.rows[1].cells] == ["c", "d"]


def test_python_docx_agrees():
    """Same document: our logical access matches python-docx's exactly."""
    docx = pytest.importorskip("docx")
    for body in (HMERGED, VMERGED):
        ours = DocxTemplate(io.BytesIO(make_docx(body)))
        theirs = docx.Document(io.BytesIO(make_docx(body)))
        t_ours = ours.tables[0]
        t_theirs = theirs.tables[0]
        rows = len(t_theirs.rows)
        cols = len(t_theirs.columns)
        for i in range(rows):
            assert [c.text for c in t_ours.rows[i].cells] == [
                c.text for c in t_theirs.rows[i].cells
            ]
            for j in range(cols):
                assert t_ours.cell(i, j).text == t_theirs.cell(i, j).text
