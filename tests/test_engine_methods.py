"""Coverage for template.rs remaining branches: dict methods, str is*/split
variants, bool-compare rewrites, printf/format internals, npgettext.
"""
import io
import re
import struct
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp  # noqa: E402
from test_final_parity import make_mo  # noqa: E402

from docxtplrs import DocxTemplate  # noqa: E402


def render_text(body, ctx=None):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


# ---------------- dict methods ----------------

def test_dict_keys_values_items():
    # engine renders dict views Python-style, like jinja2
    assert render_text(tp("{{ d.keys() }}"), {"d": {"a": 1}}) == "dict_keys(['a'])"
    assert render_text(tp("{{ d.values() }}"), {"d": {"a": 1}}) == "dict_values([1])"
    assert render_text(tp("{{ d.items() }}"), {"d": {"a": 1}}) == "dict_items([('a', 1)])"


def test_dict_get_with_default():
    assert render_text(tp("{{ d.get('a') }}"), {"d": {"a": 1}}) == "1"
    assert render_text(tp("{{ d.get('zz', 'fallback') }}"), {"d": {"a": 1}}) == "fallback"


# ---------------- str method variants ----------------

@pytest.mark.parametrize(
    "expr, ctx, expected",
    [
        ("s.isalnum()", {"s": "abc123"}, "true"),
        ("s.isalnum()", {"s": "ab c"}, "false"),
        ("s.isdecimal()", {"s": "42"}, "true"),
        ("s.isdigit()", {"s": "4a"}, "false"),
        ("s.isidentifier()", {"s": "a_b"}, "true"),
        ("s.isidentifier()", {"s": "1a"}, "false"),
        ("s.isnumeric()", {"s": "42"}, "true"),
        ("s.isspace()", {"s": "  "}, "true"),
        ("s.isupper()", {"s": "AB"}, "true"),
        ("s.isprintable()", {"s": "ab"}, "true"),
        ("s.istitle()", {"s": "ab cd"}, "false"),
        ("s.index('b')", {"s": "abc"}, "1"),
        ("s.rsplit(',')", {"s": "a,b,c"}, '["a", "b", "c"]'),
        ("s.rsplit()", {"s": "a b  c"}, '["a", "b", "c"]'),
        ("s.expandtabs(4)", {"s": "a\tb"}, "a   b"),
        ("'|'.join([])", {}, ""),
    ],
)
def test_str_method_variants(expr, ctx, expected):
    assert render_text(tp("{{ " + expr + " }}"), ctx) == expected


# ---------------- bool compare rewrites ----------------

@pytest.mark.parametrize(
    "stmt, ctx, expected",
    [
        ("{% if v == true %}Y{% endif %}", {"v": True}, "Y"),
        ("{% if v == true %}Y{% endif %}", {"v": 1}, "Y"),
        ("{% if v != true %}N{% endif %}", {"v": False}, "N"),
        ("{% if v == false %}Y{% endif %}", {"v": 0}, "Y"),
        ("{% if v != false %}Y{% endif %}", {"v": 2}, "Y"),
        ("{% if v == none %}Y{% endif %}", {"v": None}, "Y"),
        ("{% if v != none %}Y{% endif %}", {"v": 0}, "Y"),
    ],
)
def test_bool_compare_variants(stmt, ctx, expected):
    assert render_text(tp(stmt), ctx) == expected


# ---------------- printf / format internals ----------------

def test_printf_literal_percent_with_arg():
    assert render_text(tp("{{ '%s%%' % v }}"), {"v": 95}) == "95%"


def test_printf_unknown_conversion_kept():
    out = render_text(tp("{{ '%s %q' % v }}"), {"v": "x"})
    assert out.startswith("x")


def test_format_grouping_on_float():
    out = render_text(tp("{{ '{:,.1f}'.format(12345.678) }}"))
    assert out == "12,345.7"


# ---------------- npgettext (plural + context) ----------------

def test_gettext_plural_with_context():
    mo = make_mo({"fruit\x04%(count)s apple": ["%(count)s pomme", "%(count)s pommes"]})
    tpl = DocxTemplate(
        io.BytesIO(
            make_docx(
                tp(
                    '{% trans count=n context="fruit" %}{{ count }} apple'
                    "{% pluralize %}{{ count }} apples{% endtrans %}"
                )
            )
        )
    )
    tpl.install_gettext(mo)
    tpl.render({"n": 3})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert "3 pommes" in text
