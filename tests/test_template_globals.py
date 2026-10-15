"""Tests for the jinja2 template globals provided by the engine:
cycler / joiner / lipsum / namespace (mirrors jinja2's default globals)."""

import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, tp

from docxtplrs import DocxTemplate

# mirrors LOREM_WORDS in src/template.rs
LOREM_WORDS = {
    "lorem", "ipsum", "dolor", "sit", "amet", "consectetur", "adipiscing",
    "elit", "sed", "do", "eiusmod", "tempor", "incididunt", "ut", "labore",
    "et", "dolore", "magna", "aliqua", "enim", "ad", "minim", "veniam",
    "quis", "nostrud", "exercitation", "ullamco", "laboris", "nisi",
    "aliquip", "ex", "ea", "commodo", "consequat", "duis", "aute", "irure",
    "in", "reprehenderit", "voluptate", "velit", "esse", "cillum", "fugiat",
    "nulla", "pariatur", "excepteur", "sint", "occaecat", "cupidatat", "non",
    "proident", "sunt", "culpa", "qui", "officia", "deserunt", "mollit",
    "anim", "id", "est", "laborum",
}


def render_text(body, ctx=None):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def render_xml(body, ctx=None):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {})
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


# ---------------- cycler ----------------

def test_cycler_zebra_loop():
    body = tp(
        "{% set c = cycler('odd', 'even') %}"
        "{% for i in items %}{{ c.next() }} {% endfor %}"
    )
    assert render_text(body, {"items": [1, 2, 3, 4]}) == "odd even odd even "


def test_cycler_current():
    # before the first next() the last item is current (jinja2 semantics)
    assert render_text(tp("{% set c = cycler(1, 2, 3) %}{{ c.current }}")) == "3"
    # current tracks the last next()
    body = tp("{% set c = cycler(1, 2, 3) %}{{ c.next() }}{{ c.current }}{{ c.next() }}{{ c.current }}")
    assert render_text(body) == "1122"


def test_cycler_reset():
    body = tp(
        "{% set c = cycler('a', 'b') %}{{ c.next() }}{{ c.next() }}"
        "{% set _ = c.reset() %}{{ c.next() }}"
    )
    assert render_text(body) == "aba"


# ---------------- joiner ----------------

def test_joiner_joins_items():
    body = tp("{% set sep = joiner() %}{% for x in items %}{{ sep() }}{{ x }}{% endfor %}")
    assert render_text(body, {"items": ["a", "b", "c"]}) == "a, b, c"


def test_joiner_custom_sep():
    body = tp("{% set sep = joiner('|') %}{% for x in items %}{{ sep() }}{{ x }}{% endfor %}")
    assert render_text(body, {"items": ["a", "b"]}) == "a|b"


# ---------------- lipsum ----------------

def test_lipsum_html_paragraph_count():
    # html=true (default) wraps each paragraph in <p>; n controls the count
    xml = render_xml(tp("{{ lipsum(3) }}"))
    assert xml.count("<p>") == 3


def test_lipsum_text_uses_lorem_vocabulary():
    # html=false: plain sentences separated by blank lines; the newlines
    # become <w:br/> via resolve_listing
    xml = render_xml(tp("{{ lipsum(n=3, html=false, min=4, max=6) }}"))
    assert xml.count("<w:br") == 2 * (3 - 1)  # two newlines between paragraphs
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    # sentence breaks became <w:br/>, so sentences abut without a space
    words = [w.lower() for w in text.replace(".", " ").split()]
    assert words, "lipsum produced no text"
    assert all(w in LOREM_WORDS for w in words), words
    # each sentence ends with a period
    assert text.count(".") == 3


def test_lipsum_n_controls_paragraphs():
    xml1 = render_xml(tp("{{ lipsum(1, html=false, min=3, max=5) }}"))
    xml4 = render_xml(tp("{{ lipsum(4, html=false, min=3, max=5) }}"))
    text1 = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml1))
    text4 = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml4))
    assert text1.count(".") == 1
    assert text4.count(".") == 4


# ---------------- namespace ----------------

def test_namespace_accumulates():
    body = tp(
        "{% set ns = namespace(total=0) %}"
        "{% for x in items %}{% set ns.total = ns.total + x %}{% endfor %}"
        "{{ ns.total }}"
    )
    assert render_text(body, {"items": [1, 2, 3]}) == "6"
