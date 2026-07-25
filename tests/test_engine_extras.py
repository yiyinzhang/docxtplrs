"""Coverage for template.rs engine paths: str methods, format specs, printf
operator, undeclared variables, replace_* APIs, get_xml/build_url_id.
"""
import io
import struct
import sys
import zipfile
import zlib

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp  # noqa: E402

from docxtplrs import DocxTemplate, InlineImage  # noqa: E402


def render(body, ctx=None, **kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body, **kw)))
    tpl.render(ctx or {})
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


def render_text(body, ctx=None, **kw):
    import re

    xml = render(body, ctx, **kw)
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def png(tag=0x11, w=4, h=3):
    def chunk(t, d):
        return (
            struct.pack(">I", len(d)) + t + d
            + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + bytes([tag, 0, 0]) * w for _ in range(h))
    return (
        b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    )


# ---------------- str methods ----------------

@pytest.mark.parametrize(
    "expr, expected",
    [
        ("s.upper()", "ABC DEF"),
        ("s.lower()", "abc def"),
        ("s.capitalize()", "Abc def"),
        ("s.title()", "Abc Def"),
        ("s.casefold()", "abc def"),
        ("s.swapcase()", "ABC DEF"),
        ("s.strip()", "abc def"),
        ("s.replace('a', 'X')", "Xbc def"),
        # NB: minijinja renders sequences/bools JSON-style (see README)
        ("s.split('|')", '["abc def"]'),
        ("s.zfill(10)", "000abc def"),
        ("s.center(11, '*')", "  abc def  "),
        ("s.ljust(9, '-')", "abc def--"),
        ("s.rjust(9, '-')", "--abc def"),
        ("s.startswith('ab')", "true"),
        ("s.endswith('x')", "false"),
        ("s.find('d')", "4"),
        ("s.count('b')", "1"),
        ("s.islower()", "true"),
        ("s.isalpha()", "false"),
        ("s.removeprefix('ab')", "c def"),
        ("s.removesuffix('ef')", "abc d"),
        ("'|'.join(s.split())", "abc|def"),
        ("s.partition(' ')", '["abc", " ", "def"]'),
        ("s.rpartition(' ')", '["abc", " ", "def"]'),
        ("s.expandtabs(4)", "abc def"),
    ],
)
def test_str_methods(expr, expected):
    body = tp("{{ " + expr + " }}")
    assert render_text(body, {"s": "abc def"}) == expected


def test_str_splitlines_and_misc():
    assert render_text(tp("{{ s.splitlines() }}"), {"s": "a\nb"}) == '["a", "b"]'
    assert render_text(tp("{{ s.isdigit() }}"), {"s": "42"}) == "true"
    assert render_text(tp("{{ s.istitle() }}"), {"s": "Abc Def"}) == "true"


# ---------------- format specs (str.format / %-operator) ----------------

@pytest.mark.parametrize(
    "expr, expected",
    [
        ("'{:>8}'.format(42)", "      42"),
        ("'{:<8}'.format(42)", "42      "),
        ("'{:^8}'.format(42)", "   42   "),
        ("'{:08.2f}'.format(3.14159)", "00003.14"),
        ("'{:,}'.format(1234567)", "1,234,567"),
        ("'{:_}'.format(1234567)", "1_234_567"),
        ("'{:.2f}'.format(3.14159)", "3.14"),
        ("'{:#x}'.format(255)", "0xff"),
        ("'{:#X}'.format(255)", "0XFF"),
        ("'{:#o}'.format(8)", "0o10"),
        ("'{:#b}'.format(5)", "0b101"),
        ("'{:e}'.format(12345.678)", "1.234568e+04"),
        ("'{:.1%}'.format(0.256)", "25.6%"),
        ("'{:+d}'.format(5)", "+5"),
        ("'{:g}'.format(0.00001234)", "1.234e-5"),
        ("'{:g}'.format(123456.0)", "123456"),
    ],
)
def test_format_specs(expr, expected):
    assert render_text(tp("{{ " + expr + " }}")) == expected


@pytest.mark.parametrize(
    "expr, ctx, expected",
    [
        ("'%s' % v", {"v": "x"}, "x"),
        ("'%d items' % 3", {}, "3 items"),
        ("'%.2f' % 3.14159", {}, "3.14"),
        ("'%s-%s' % (a, b)", {"a": "x", "b": "y"}, "x-y"),
        ("'%d%%' % 95", {}, "95%"),
        ("'%05.1f' % 3.14", {}, "003.1"),
        ("'%x' % 255", {}, "ff"),
        ("'%e' % 12345.678", {}, "1.234568e+04"),
        ("'%g' % 0.00001234", {}, "1.234e-5"),
    ],
)
def test_printf_operator(expr, ctx, expected):
    assert render_text(tp("{{ " + expr + " }}"), ctx) == expected


# ---------------- undeclared variables ----------------

def test_undeclared_variables():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ a }} {{ b.c }} {% if d %}x{% endif %}"))))
    assert tpl.get_undeclared_template_variables() == {"a", "b", "d"}


def test_undeclared_variables_with_context_filter():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ a }} {{ b }}"))))
    assert tpl.get_undeclared_template_variables(context={"a": 1}) == {"b"}


# ---------------- get_xml / get_docx_bytes ----------------

def test_get_xml_contains_body():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hello"))))
    xml = tpl.get_xml()
    assert "hello" in xml


def test_get_docx_bytes_roundtrip():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ x }}"))))
    tpl.render({"x": "done"})
    data = tpl.get_docx_bytes()
    assert zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")


# ---------------- build_url_id ----------------

def test_build_url_id_adds_external_rel():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    rid = tpl.build_url_id("https://example.com")
    assert rid.startswith("rId")
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    rels = read_docx_part(out.getvalue(), "word/_rels/document.xml.rels")
    assert "https://example.com" in rels


# ---------------- replace_* APIs ----------------

def test_replace_media_by_crc():
    old, new = png(0x11), png(0x22)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"), media={"image1.png": old})))
    tpl.replace_media(io.BytesIO(old), io.BytesIO(new))
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert z.read("word/media/image1.png") == new


def test_replace_zipname():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.replace_zipname("word/document.xml", io.BytesIO(b"<w:document/>"))
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    assert read_docx_part(out.getvalue(), "word/document.xml") == "<w:document/>"


def test_reset_replacements():
    old, new = png(0x11), png(0x22)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"), media={"image1.png": old})))
    tpl.replace_media(io.BytesIO(old), io.BytesIO(new))
    tpl.reset_replacements()
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert z.read("word/media/image1.png") == old


def test_replace_pic_by_name(tmp_path):
    old, new = png(0x11), png(0x22)
    p = tmp_path / "logo.png"
    p.write_bytes(old)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ img }}"))))
    tpl.replace_pic("logo.png", io.BytesIO(new))
    tpl.render({"img": InlineImage(tpl, str(p))})
    out = io.BytesIO()
    tpl.save(out)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert z.read("word/media/image1.png") == new


def test_replace_pic_missing_raises():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.replace_pic("nope.png", io.BytesIO(png()))
    tpl.render({})
    with pytest.raises(Exception):
        tpl.save(io.BytesIO())


def test_allow_missing_pics():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.allow_missing_pics = True
    tpl.replace_pic("nope.png", io.BytesIO(png()))
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    assert out.getvalue()
