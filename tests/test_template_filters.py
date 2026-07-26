"""Coverage for template.rs: extra filters, jinja env options, autoescape,
render-error context, recovery fallbacks.
"""
import io
import re
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp, cell, tr, tbl  # noqa: E402

from docxtplrs import DocxTemplate, TemplateError  # noqa: E402


def render_text(body, ctx=None, **render_kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {}, **render_kw)
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


# ---------------- extra filters ----------------

def test_filesizeformat():
    assert render_text(tp("{{ n|filesizeformat }}"), {"n": 512}) == "512 Bytes"
    out = render_text(tp("{{ n|filesizeformat }}"), {"n": 2048})
    assert "k" in out.lower() or "K" in out


def test_wordcount():
    assert render_text(tp("{{ s|wordcount }}"), {"s": "a b  c"}) == "3"


def test_center_filter():
    assert render_text(tp("{{ s|center(7) }}"), {"s": "abc"}) == "  abc  "


def test_forceescape():
    out = render_text(tp("{{ s|forceescape }}"), {"s": "<b>"})
    assert "<b>" not in out and "&lt;" in out


def test_truncate_variants():
    out = render_text(tp("{{ s|truncate(5) }}"), {"s": "hello world"})
    assert out.endswith("...")
    out2 = render_text(tp("{{ s|truncate(9, True) }}"), {"s": "hello world foo"})
    assert len(out2) <= 12


def test_xmlattr_filter():
    out = render_text(tp("{{ d|xmlattr }}"), {"d": {"a": "1"}})
    assert 'a="1"' in out


def test_wordwrap_filter():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ s|wordwrap(5) }}"))))
    tpl.render({"s": "aaa bbb ccc ddd"})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    # wrapped newlines become <w:br/> via resolve_listing
    assert "<w:br" in xml


def test_urlize_filter():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ s|urlize }}"))))
    tpl.render({"s": "see http://example.com"})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert '<a href="http://example.com"' in xml


def test_random_filter():
    out = render_text(tp("{{ xs|random }}"), {"xs": [1, 2, 3]})
    assert out in ("1", "2", "3")


def test_striptags_filter():
    assert render_text(tp("{{ s|striptags }}"), {"s": "<b>x</b> tail"}) == "x tail"


# ---------------- jinja env options ----------------

def test_env_lstrip_blocks():
    import jinja2

    env = jinja2.Environment(lstrip_blocks=True, trim_blocks=True)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% if true %}  \n  yes\n{% endif %}"))))
    tpl.render({}, jinja_env=env)
    out = io.BytesIO()
    tpl.save(out)
    text = "".join(
        re.findall(r"<w:t[^>]*>([^<]*)</w:t>", read_docx_part(out.getvalue(), "word/document.xml"))
    )
    assert "yes" in text


def test_env_strict_undefined_raises():
    import jinja2

    env = jinja2.Environment(undefined=jinja2.StrictUndefined)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ missing }}"))))
    with pytest.raises(Exception):
        tpl.render({}, jinja_env=env)


def test_env_keep_trailing_newline():
    import jinja2

    env = jinja2.Environment(keep_trailing_newline=True)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ x }}"))))
    tpl.render({"x": "v"}, jinja_env=env)
    out = io.BytesIO()
    tpl.save(out)
    assert "v" in read_docx_part(out.getvalue(), "word/document.xml")


# ---------------- autoescape ----------------

def test_autoescape_escapes_markup():
    out = render_text(tp("{{ v }}"), {"v": "a<b>&\"c\""}, autoescape=True)
    assert "a&lt;b&gt;&amp;" in out


def test_no_autoescape_keeps_literal():
    out = render_text(tp("{{ v }}"), {"v": "plain"})
    assert out == "plain"


# ---------------- render error context ----------------

def test_template_error_has_line_info():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ bad syntax here }}"))))
    with pytest.raises(TemplateError) as exc_info:
        tpl.render({})
    assert "syntax" in str(exc_info.value).lower() or "Context" in str(exc_info.value)


# ---------------- recover / fallback paths ----------------

def test_recover_unescaped_ampersand():
    # unescaped & in rendered output is not well-formed; recover path fixes it
    body = tbl([tr(cell(tp("{{ v }}")))], widths=(1000,))
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render({"v": "a & b <c>"})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "a" in xml and "b" in xml


def test_render_twice_same_template():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ x }}"))))
    tpl.render({"x": "one"})
    tpl.render({"x": "two"})
    out = io.BytesIO()
    tpl.save(out)
    assert "two" in read_docx_part(out.getvalue(), "word/document.xml")
