"""Coverage for patch.rs secondary branches: hm with existing gridSpan,
colspan cleanup, clean_tags (smart quotes/entities), comment element tags,
richtext statement form, brace escaping, dash merges, split tags.
"""
import io
import re
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, run, p, tp, cell, tr, tbl  # noqa: E402

from docxtplrs import DocxTemplate, R  # noqa: E402


def render(body, ctx=None):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {})
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


# ---------------- hm with existing gridSpan ----------------

def test_hm_multiplies_existing_gridspan():
    tcpr = '<w:gridSpan w:val="2"/>'
    body = tbl(
        [
            tr(cell(tp("{%tr for x in items %}"))),
            tr(cell(tp("{% hm %}{{ x }}"), tcpr=tcpr)),
            tr(cell(tp("{%tr endfor %}"))),
        ],
        widths=(1000, 1000),
    )
    xml = render(body, {"items": [1, 2]})
    # gridSpan 2 * loop.length(2) = 4 on the first (kept) cell
    assert 'w:gridSpan w:val="4"' in xml


# ---------------- colspan cleanup branches ----------------

def test_colspan_replaces_existing_gridspan():
    tcpr = '<w:gridSpan w:val="3"/>'
    body = tbl(
        [
            tr(cell(tp("{% colspan span %}"), tcpr=tcpr), cell(tp("{{ v }}"))),
        ],
        widths=(1000, 1000),
    )
    xml = render(body, {"span": 2, "v": "x"})
    assert 'w:gridSpan w:val="2"' in xml
    assert 'w:gridSpan w:val="3"' not in xml


def test_cellbg_replaces_existing_shd():
    tcpr = '<w:shd w:val="clear" w:color="auto" w:fill="000000"/>'
    body = tbl(
        [tr(cell(tp("{% cellbg color %}"), tcpr=tcpr), cell(tp("x")))],
        widths=(1000, 1000),
    )
    xml = render(body, {"color": "FF0000"})
    assert 'w:fill="FF0000"' in xml
    assert 'w:fill="000000"' not in xml


# ---------------- clean_tags: smart quotes & entities ----------------

def test_smart_quotes_inside_tag():
    # Word often auto-replaces quotes inside template tags
    xml = render(tp("{{ \u201cv\u201d }}"))
    assert ">v<" in xml


def test_numeric_quote_entities_inside_tag():
    xml = render(tp("{{ &#8216;v&#8217; }}"))
    assert ">v<" in xml


def test_entity_decoding_quot_apos():
    xml = render(tp("{{ x }}"), {"x": "&quot;"})
    # &quot; inside a text node is decoded like lxml would
    assert '"&quot;"' not in xml


# ---------------- comment element tags ----------------

def test_comment_tr_removes_row():
    body = tbl(
        [
            tr(cell(tp("{#tr drop this row #}"))),
            tr(cell(tp("kept"))),
        ],
        widths=(1000,),
    )
    xml = render(body)
    assert "kept" in xml and "drop this row" not in xml


def test_comment_p_removes_paragraph():
    body = tp("{#p drop me #}") + tp("kept")
    xml = render(body)
    assert "kept" in xml and "drop me" not in xml


def test_comment_tc_removes_cell():
    body = tbl(
        [tr(cell(tp("{#tc drop #}")), cell(tp("kept")))],
        widths=(1000, 1000),
    )
    xml = render(body)
    assert "kept" in xml and "drop" not in xml


# ---------------- richtext statement form ----------------
# NB: {%r ... %} is rewritten to {% rt %} upstream-style and rejected by the
# engine just like docxtpl does; only the {{r var }} form is supported.

def test_richtext_variable_form():
    xml = render(tp("{{r rt }}"), {"rt": R("bold", bold=True)})
    assert "<w:b/>" in xml and "bold" in xml


# ---------------- brace escaping ----------------

def test_escaped_braces_render_literally():
    xml = render(tp("{_{ not_a_var }_}"))
    assert "{{ not_a_var }}" in xml


def test_escaped_statement_braces():
    xml = render(tp("{_% if x %_}"))
    assert "{% if x %}" in xml


# ---------------- dash merges ----------------

def test_dash_merge_prev_cond_false():
    body = tp("A {%- if cond %}") + tp("B{% endif %}")
    xml = render(body, {"cond": False})
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert "A" in text and "B" not in text


def test_dash_merge_next():
    body = tp("{% if cond -%}") + tp("B{% endif %}")
    xml = render(body, {"cond": True})
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert "B" in text


def test_dash_merge_in_for_loop():
    body = tp("{% for i in items -%}") + tp("{{ i }} ") + tp("{%- endfor %}")
    xml = render(body, {"items": [1, 2, 3]})
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert "1" in text and "2" in text and "3" in text


# ---------------- tags split across runs ----------------

def test_jinja_tag_split_across_runs():
    # { in one run, { var } in the next — merge_split_braces joins them
    body = p(run("{"), run("{ v }}"), run(" tail"))
    xml = render(body, {"v": "joined"})
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert "joined" in text


def test_statement_tag_split_across_runs():
    body = p(run("{%"), run(" if x %}"), run("yes"), run("{%"), run(" endif %}"))
    xml = render(body, {"x": True})
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert "yes" in text
