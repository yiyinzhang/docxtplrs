"""Tests for the remaining jinja2 parity features."""

import io
import os
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
    return read_docx_part(out.getvalue(), "word/document.xml")


def test_bool_none_rendering():
    t = text_of(render_xml(tp("{{ a }}|{{ b }}|{{ c }}"), {"a": True, "b": False, "c": None}))
    assert "True|False|None" in t


def test_bool_none_semantics():
    body = tp(
        "{% if a %}A{% endif %}{% if not b %}NB{% endif %}"
        "{% if a is true %}IT{% endif %}{% if b is false %}IF{% endif %}"
        "{% if a is boolean %}IB{% endif %}{% if c is none %}IN{% endif %}"
        "{% if a == b %}EQ{% endif %}"
    )
    t = text_of(render_xml(body, {"a": True, "b": True, "c": None}))
    assert "A" in t and "NB" not in t and "IT" in t and "IF" not in t
    assert "IB" in t and "IN" in t and "EQ" in t


def test_bool_in_filters_and_loops():
    body = tp("{% for v in vals %}{{ v }};{% endfor %}")
    t = text_of(render_xml(body, {"vals": [True, False, None, 1]}))
    assert "True;False;None;1;" in t


def test_custom_filter_receives_real_bool():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ v|kind }}"))))
    tpl.register_filter("kind", lambda v: type(v).__name__)
    tpl.render({"v": False})
    out = io.BytesIO()
    tpl.save(out)
    assert "bool" in text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def test_callable_test():
    body = tp("{% if f is callable %}C{% endif %}{% if s is not callable %}NC{% endif %}")
    t = text_of(render_xml(body, {"f": lambda: 1, "s": "str"}))
    assert "C" in t and "NC" in t


def test_in_operator_on_dict():
    body = tp("{% if 'a' in d %}Y{% endif %}{% if 'z' not in d %}N{% endif %}")
    t = text_of(render_xml(body, {"d": {"a": 1}}))
    assert "Y" in t and "N" in t


def test_in_operator_on_list():
    body = tp("{% if 2 in lst %}Y{% endif %}")
    t = text_of(render_xml(body, {"lst": [1, 2, 3]}))
    assert "Y" in t


def test_trans_block():
    body = tp("{% trans %}Hello {{ name }}!{% endtrans %}")
    t = text_of(render_xml(body, {"name": "World"}))
    assert "Hello World!" in t


def test_trans_with_assignments():
    body = tp("{% trans user=name %}Hi {{ user }}{% endtrans %}")
    t = text_of(render_xml(body, {"name": "Bob"}))
    assert "Hi Bob" in t


def test_recover_mode_fixes_tables():
    # unescaped value producing ill-formed xml; fix_tables should still run
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
    xml = render_xml(body, {"items": ["a < 3", "b < 4 & x", "e"]})
    # recovered: 3 cells -> grid grown to 3 columns
    assert xml.count("<w:gridCol") == 3


def test_zip_timestamps_preserved():
    import zipfile

    src = make_docx(tp("{{ v }}"))
    with zipfile.ZipFile(io.BytesIO(src)) as z:
        orig_time = z.getinfo("word/document.xml").date_time
    tpl = DocxTemplate(io.BytesIO(src))
    tpl.render({"v": "x"})
    out = io.BytesIO()
    tpl.save(out)
    with zipfile.ZipFile(out) as z:
        new_time = z.getinfo("word/document.xml").date_time
    assert orig_time == new_time


class EnvOptions:
    def __init__(self, **kw):
        for k, v in kw.items():
            setattr(self, k, v)
        self.autoescape = False
        self.filters = {}
        self.globals = {}
        self.tests = {}


def test_jinja_env_trim_blocks():
    class ChainableUndefined:
        pass

    env = EnvOptions(trim_blocks=True, undefined=ChainableUndefined)
    body = tp("{% if True %}\n{{ missing.attr.deep }}\n{% endif %}")
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render({}, jinja_env=env)
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    # chainable: deep access on undefined is allowed; trim_blocks eats the newline
    assert "\n" not in text_of(xml).strip()


def test_jinja_env_strict_undefined():
    class StrictUndefined:
        pass

    env = EnvOptions(undefined=StrictUndefined)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ missing }}"))))
    with pytest.raises(Exception):
        tpl.render({}, jinja_env=env)
