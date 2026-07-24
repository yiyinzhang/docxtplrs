"""Tests for the final parity round: bool compare rewrite, missing filters,
real gettext catalogs."""

import io
import os
import struct
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, tp

from docxtplrs import DocxTemplate


def render_xml(body, context, autoescape=False, **kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body, **kw)))
    tpl.render(context, autoescape=autoescape)
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml"), tpl


def t_of(body, context, **kw):
    xml, _ = render_xml(body, context, **kw)
    return text_of(xml)


# ---------------- bool literal comparisons ----------------

def test_bool_literal_comparison():
    body = tp(
        "{% if a == true %}EQ{% endif %}{% if b != true %}NE{% endif %}"
        "{% if a == false %}BAD{% endif %}{% if c == none %}NN{% endif %}"
    )
    t = t_of(body, {"a": True, "b": False, "c": None})
    assert "EQ" in t and "NE" in t and "BAD" not in t and "NN" in t


# ---------------- missing filters ----------------

def test_filesizeformat():
    body = tp("{{ 100|filesizeformat }}|{{ 1500|filesizeformat }}|{{ 1500|filesizeformat(true) }}")
    t = t_of(body, {})
    assert "100 Bytes|1.5 kB|1.5 KiB" in t


def test_wordcount_center_forceescape():
    body = tp("{{ 'one two  three'|wordcount }}|{{ 'x'|center(5) }}|{{ s|forceescape }}")
    t = t_of(body, {"s": "a<b>"})
    assert "3" in t and "  x  " in t
    xml, _ = render_xml(body, {"s": "a<b>"})
    assert "a&lt;b&gt;" in xml


def test_truncate():
    body = tp("{{ s|truncate(10) }}|{{ s|truncate(10, true) }}")
    t = t_of(body, {"s": "one two three four five"})
    # jinja2: killwords=False cuts at the last space, killwords=True cuts hard
    assert "one...|one two..." in t


def test_xmlattr():
    body = tp("{{ d|xmlattr }}")
    t = t_of(body, {"d": {"class": "a&b", "id": "x"}})
    assert 'class="a&amp;b"' in t and 'id="x"' in t


def test_wordwrap():
    body = tp("{{ s|wordwrap(10) }}")
    t = t_of(body, {"s": "aaa bbb ccc ddd"})
    assert "aaa bbb" in t


def test_urlize():
    body = tp("{{ s|urlize }}")
    xml, _ = render_xml(body, {"s": "see https://example.com/page. end"})
    assert '<a href="https://example.com/page"' in xml
    assert "rel=\"noopener\"" in xml


def test_random():
    body = tp("{{ lst|random }}")
    t = t_of(body, {"lst": ["only"]})
    assert "only" in t


# ---------------- gettext ----------------

def make_mo(entries, plural_forms="nplurals=2; plural=(n != 1);"):
    """Build a minimal .mo file. entries: {msgid: msgstr or [forms]}"""
    header = "Content-Type: text/plain; charset=UTF-8\nPlural-Forms: %s\n" % plural_forms
    items = [("", header)]
    for k, v in entries.items():
        if isinstance(v, list):
            items.append((k, "\0".join(v)))
        else:
            items.append((k, v))
    items.sort(key=lambda x: x[0])
    n = len(items)
    orig_table_off = 28
    trans_table_off = orig_table_off + n * 8
    data_off = trans_table_off + n * 8

    out = b""
    orig_entries = []
    trans_entries = []
    blob = b""
    for k, v in items:
        kb = k.encode()
        vb = v.encode()
        orig_entries.append((len(kb), data_off + len(blob)))
        blob += kb + b"\0"
    for k, v in items:
        vb = v.encode()
        trans_entries.append((len(vb), data_off + len(blob)))
        blob += vb + b"\0"

    out = struct.pack("<IIIIIII", 0x950412DE, 0, n, orig_table_off, trans_table_off, 0, 0)
    for length, off in orig_entries:
        out += struct.pack("<II", length, off)
    for length, off in trans_entries:
        out += struct.pack("<II", length, off)
    return out + blob


def test_gettext_translation():
    mo = make_mo({"Hello %(name)s!": "Bonjour %(name)s !"})
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% trans %}Hello {{ name }}!{% endtrans %}"))))
    tpl.install_gettext(mo)
    tpl.render({"name": "World"})
    out = io.BytesIO()
    tpl.save(out)
    t = text_of(read_docx_part(out.getvalue(), "word/document.xml"))
    assert "Bonjour World !" in t


def test_gettext_plural():
    mo = make_mo({"%(count)s apple": ["%(count)s pomme", "%(count)s pommes"]})
    tpl = DocxTemplate(
        io.BytesIO(
            make_docx(
                tp("{% trans count=n %}{{ count }} apple{% pluralize %}{{ count }} apples{% endtrans %}")
            )
        )
    )
    tpl.install_gettext(mo)
    for n, expect in [(1, "1 pomme"), (3, "3 pommes")]:
        tpl.render({"n": n})
        out = io.BytesIO()
        tpl.save(out)
        assert expect in text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def test_gettext_missing_entry_fallback():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% trans %}Untranslated {{ v }}{% endtrans %}"))))
    tpl.install_gettext(make_mo({}))
    tpl.render({"v": "fallback"})
    out = io.BytesIO()
    tpl.save(out)
    assert "Untranslated fallback" in text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def test_gettext_no_catalog_identity():
    # without a catalog, trans renders the interpolated content
    t = t_of(tp("{% trans %}Hi {{ name }}{% endtrans %}"), {"name": "X"})
    assert "Hi X" in t


def test_gettext_context():
    mo = make_mo({})
    # context entries use \x04 separator
    entries = {"verb\x04Open": "Ouvrir-verbe"}
    mo = make_mo(entries)
    tpl = DocxTemplate(
        io.BytesIO(make_docx(tp('{% trans context="verb" %}Open{% endtrans %}')))
    )
    tpl.install_gettext(mo)
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    assert "Ouvrir-verbe" in text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def test_bool_compare_exact_python_semantics():
    # 1 == True and 0 == False are True in Python (jinja2)
    body = tp(
        "{% if 1 == true %}A{% endif %}{% if 0 == false %}B{% endif %}"
        "{% if 2 == true %}BAD1{% endif %}{% if '1' == true %}BAD2{% endif %}"
        "{% if a == true %}C{% endif %}{% if b != true %}D{% endif %}"
    )
    t = t_of(body, {"a": True, "b": False})
    assert "A" in t and "B" in t and "C" in t and "D" in t
    assert "BAD1" not in t and "BAD2" not in t


def test_bool_compare_ignores_literal_text():
    # the literal text "== true" outside tags must not be rewritten
    body = tp("x == true y {% if a == true %}Z{% endif %}")
    t = t_of(body, {"a": True})
    assert "x == true y" in t and "Z" in t


def test_bool_compare_ignores_string_literals():
    body = tp('{{ "== true" }}{% if a == true %}Z{% endif %}')
    t = t_of(body, {"a": True})
    assert "== true" in t and "Z" in t


def test_regex_fallback_table_fix():
    # severely broken xml (unescaped <) that even recover cannot parse:
    # regex fallback must still grow the grid
    body = (
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>'
        + "</w:tblGrid>"
        + "<w:tr>"
        + '<w:tc><w:tcPr/><w:p><w:r><w:t>{%tc for x in items %}</w:t></w:r></w:p></w:tc>'
        + '<w:tc><w:tcPr/><w:p><w:r><w:t>{{ x }}</w:t></w:r></w:p></w:tc>'
        + '<w:tc><w:tcPr/><w:p><w:r><w:t>{%tc endfor %}</w:t></w:r></w:p></w:tc>'
        + "</w:tr></w:tbl>"
    )
    # "a<b" breaks xml in a way the whitelist recovery also fixes; use a value
    # that breaks structure beyond recovery: unclosed quote inside an attribute
    xml, _ = render_xml(body, {"items": ['x"><w:', "y", "z"]})
    assert xml.count("<w:gridCol") == 3
