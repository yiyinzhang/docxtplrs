"""Tests for jinja2 parity features: custom filters, loader, __html__,
python str methods, format specs, TemplateError, object semantics."""

import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, tp

from docxtplrs import DocxTemplate, TemplateError


def render_xml(body, context, autoescape=False, tpl=None, **kw):
    tpl = tpl or DocxTemplate(io.BytesIO(make_docx(body, **kw)))
    tpl.render(context, autoescape=autoescape)
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


def test_custom_filter():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ v|shout }}"))))
    tpl.register_filter("shout", lambda v: str(v).upper() + "!")
    xml = render_xml(None, {"v": "hey"}, tpl=tpl)
    assert "HEY!" in text_of(xml)


def test_custom_filter_with_args():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ v|rep(3, '-') }}"))))
    tpl.register_filter("rep", lambda v, n, sep: sep.join([str(v)] * n))
    xml = render_xml(None, {"v": "ab"}, tpl=tpl)
    assert "ab-ab-ab" in text_of(xml)


def test_custom_test():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% if v is even %}E{% endif %}"))))
    tpl.register_test("even", lambda v: v % 2 == 0)
    xml = render_xml(None, {"v": 4}, tpl=tpl)
    assert "E" in text_of(xml)


def test_custom_function_and_global():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ add(2, 3) }}-{{ company }}"))))
    tpl.register_function("add", lambda a, b: a + b)
    tpl.register_global("company", "ACME")
    xml = render_xml(None, {}, tpl=tpl)
    assert "5-ACME" in text_of(xml)


class FakeJinjaEnv:
    """duck-typed jinja2.Environment stand-in"""

    def __init__(self):
        self.autoescape = True
        self.filters = {"double": lambda v: str(v) * 2}
        self.globals = {"gvar": "G"}
        self.tests = {}


def test_jinja_env_duck_typing():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ v|double }} {{ gvar }} {{ esc }}"))))
    tpl.render({"v": "x", "esc": "a<b"}, jinja_env=FakeJinjaEnv())
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    t = text_of(xml)
    assert "xx G" in t
    assert "a&lt;b" in xml  # autoescape picked up from env


def test_template_loader_include():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% include 'footer_tpl' %}"))))
    tpl.set_template_loader(lambda name: "included {{ v }}" if name == "footer_tpl" else None)
    xml = render_xml(None, {"v": "V"}, tpl=tpl)
    assert "included V" in text_of(xml)


def test_template_loader_import():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% import 'macros' as m %}{{ m.hello('Bob') }}"))))
    tpl.set_template_loader(
        lambda name: "{% macro hello(n) %}Hi {{ n }}!{% endmacro %}" if name == "macros" else None
    )
    xml = render_xml(None, {}, tpl=tpl)
    assert "Hi Bob!" in text_of(xml)


def test_html_protocol():
    class SafeXml:
        def __html__(self):
            return "<b>raw</b>"

        def __str__(self):
            return "str-version"

    xml = render_xml(tp("{{ o }}"), {"o": SafeXml()})
    assert "<b>raw</b>" in xml


def test_str_methods():
    body = tp("{{ s.capitalize() }}|{{ s.title() }}|{{ s.swapcase() }}|{{ 'a,b'.split(',')[1] }}|{{ '-'.join(lst) }}")
    t = text_of(render_xml(body, {"s": "hello world", "lst": ["a", "b"]}))
    assert "Hello world|Hello World|HELLO WORLD|b|a-b" in t


def test_str_methods_extended():
    body = tp(
        "{{ 'a=b=c'.partition('=')[0] }}|{{ 'a=b=c'.rpartition('=')[2] }}"
        "|{{ 'l1\\nl2'.splitlines()|length }}|{{ 'Ab'.casefold() }}"
        "|{{ '42'.zfill(5) }}|{{ 'x'.center(5) }}|{{ s.isdigit() }}"
    )
    t = text_of(render_xml(body, {"s": "123"}))
    assert "a|c|2|ab|00042|  x  |true" in t


def test_format_method():
    body = tp("{{ '{} and {0}'.format('A', 'B') }}|{{ '{:.2f}'.format(pi) }}|{{ '{:>8}'.format(42) }}|{{ '{:,}'.format(n) }}")
    t = text_of(render_xml(body, {"pi": 3.14159, "n": 1234567}))
    assert "A and A|3.14|      42|1,234,567" in t


def test_format_spec_variants():
    body = tp("{{ '{:+}'.format(5) }}|{{ '{:#x}'.format(255) }}|{{ '{:08.3f}'.format(3.14159) }}|{{ '{:.1%}'.format(0.25) }}|{{ '{:e}'.format(12345.0) }}")
    t = text_of(render_xml(body, {}))
    assert "+5|0xff|0003.142|25.0%|1.234500e+04" in t


def test_template_error_raised():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% badtag %}"))))
    with pytest.raises(TemplateError):
        tpl.render({})


def test_undefined_length_is_zero():
    t = text_of(render_xml(tp("{{ missing|length }}"), {}))
    assert "0" in t


def test_big_int():
    big = 2**100
    t = text_of(render_xml(tp("{{ n + 1 }}"), {"n": big}))
    assert str(big + 1) in t


def test_object_truthiness():
    class Empty:
        def __len__(self):
            return 0

    t = text_of(render_xml(tp("{% if o %}T{% else %}F{% endif %}"), {"o": Empty()}))
    assert "F" in t


def test_object_comparison():
    class Money:
        def __init__(self, v):
            self.v = v

        def __eq__(self, other):
            return isinstance(other, Money) and self.v == other.v

        def __lt__(self, other):
            return self.v < other.v

    t = text_of(
        render_xml(
            tp("{% if a == b %}EQ{% endif %}{% if a < c %}LT{% endif %}"),
            {"a": Money(5), "b": Money(5), "c": Money(10)},
        )
    )
    assert "EQ" in t and "LT" in t


def test_dict_order_preserved():
    body = tp("{% for k in d %}{{ k }}{% endfor %}")
    t = text_of(render_xml(body, {"d": {"zebra": 1, "apple": 2, "mango": 3}}))
    assert "zebraapplemango" in t


def test_dict_methods():
    body = tp("{{ d.get('x', 'dflt') }}-{{ d.get('a') }}-{{ d.items()|length }}")
    t = text_of(render_xml(body, {"d": {"a": 1}}))
    assert "dflt-1-1" in t
