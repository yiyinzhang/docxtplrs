"""Coverage for pyclasses.rs: constructor input forms, save variants,
Length units, RichText/RP/Listing, register_*, template loader, pic_map.
"""
import io
import pathlib
import re
import sys
import zipfile

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp  # noqa: E402

from docxtplrs import (  # noqa: E402
    Cm,
    DocxTemplate,
    Emu,
    Inches,
    Length,
    Listing,
    Mm,
    Pt,
    R,
    RP,
    Twips,
)


def render_text(body, ctx=None):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


# ---------------- constructor input forms ----------------

def test_ctor_from_bytes():
    tpl = DocxTemplate(make_docx(tp("{{ x }}")))
    tpl.render({"x": "ok"})
    assert tpl.get_xml()


def test_ctor_from_str_path(tmp_path):
    f = tmp_path / "t.docx"
    f.write_bytes(make_docx(tp("{{ x }}")))
    tpl = DocxTemplate(str(f))
    tpl.render({"x": "ok"})
    out = io.BytesIO()
    tpl.save(out)
    assert "ok" in read_docx_part(out.getvalue(), "word/document.xml")


def test_ctor_from_pathlike(tmp_path):
    f = tmp_path / "t.docx"
    f.write_bytes(make_docx(tp("{{ x }}")))
    tpl = DocxTemplate(pathlib.Path(f))
    tpl.render({"x": "ok"})
    assert tpl.get_xml()


def test_ctor_from_file_like(tmp_path):
    f = tmp_path / "t.docx"
    f.write_bytes(make_docx(tp("{{ x }}")))
    with open(f, "rb") as fh:
        tpl = DocxTemplate(fh)
    tpl.render({"x": "ok"})
    assert tpl.get_xml()


def test_ctor_invalid_raises():
    with pytest.raises(Exception):
        DocxTemplate(12345)


def test_ctor_missing_file_raises():
    with pytest.raises(Exception):
        DocxTemplate("/nonexistent/path/t.docx")


# ---------------- save variants ----------------

def test_save_to_str_path(tmp_path):
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hi"))))
    tpl.render({})
    target = tmp_path / "out.docx"
    tpl.save(str(target))
    with zipfile.ZipFile(target) as z:
        assert "hi" in z.read("word/document.xml").decode()


def test_write_xml(tmp_path):
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hi"))))
    target = tmp_path / "doc.xml"
    tpl.write_xml(str(target))
    assert "hi" in target.read_text()


# ---------------- Length ----------------

def test_length_units():
    assert Emu(914400).emu == 914400
    assert Inches(1).emu == 914400
    assert Cm(1).emu == 360000
    assert Mm(1).emu == 36000
    assert Pt(1).emu == 12700
    assert Twips(1).emu == 635
    assert Length(914400).emu == 914400


def test_length_conversions():
    x = Inches(1)
    assert x.inches == 1.0
    assert abs(x.cm - 2.54) < 1e-9
    assert abs(x.mm - 25.4) < 1e-9
    assert abs(x.pt - 72.0) < 1e-9
    assert x.twips == 1440.0
    assert int(x) == 914400
    assert "914400" in repr(x)


# ---------------- RichText / RP / Listing ----------------

def test_richtext_add_and_dunder():
    rt = R("plain")
    rt.add("bold", bold=True)
    rt.add("colored", color="#FF0000")
    s = str(rt)
    assert "<w:b/>" in s and 'w:val="FF0000"' in s
    assert rt.__html__() == s
    xml = render_text(tp("{{r rt }}"), {"rt": rt})
    assert "plainboldcolored" == xml


def test_richtextparagraph():
    rp = RP("first")
    rp.add(" second", italic=True)
    assert rp.__html__() == str(rp)
    xml = render_text(tp("{{p rp }}"), {"rp": rp})
    assert "first second" == xml


def test_listing_newlines_tabs():
    lst = Listing("line1\nline2\ta")
    assert lst.__html__() == str(lst)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ l }}"))))
    tpl.render({"l": lst})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "<w:br/>" in xml and "<w:tab/>" in xml


# ---------------- register_function / register_global ----------------

def test_register_function_and_global():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ add(1, 2) }} {{ company }}"))))
    tpl.register_function("add", lambda a, b: a + b)
    tpl.register_global("company", "ACME")
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    text = "".join(
        re.findall(r"<w:t[^>]*>([^<]*)</w:t>", read_docx_part(out.getvalue(), "word/document.xml"))
    )
    assert text == "3 ACME"


# ---------------- template loader: include / import ----------------

def test_include_via_loader():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp('{% include "part" %}'))))
    tpl.set_template_loader(lambda name: "included {{ v }}" if name == "part" else None)
    tpl.render({"v": "YES"})
    out = io.BytesIO()
    tpl.save(out)
    text = "".join(
        re.findall(r"<w:t[^>]*>([^<]*)</w:t>", read_docx_part(out.getvalue(), "word/document.xml"))
    )
    assert text == "included YES"


def test_import_macro_via_loader():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp('{% import "m" as m %}{{ m.hi() }}'))))
    tpl.set_template_loader(
        lambda name: "{% macro hi() %}MACRO{% endmacro %}" if name == "m" else None
    )
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    text = "".join(
        re.findall(r"<w:t[^>]*>([^<]*)</w:t>", read_docx_part(out.getvalue(), "word/document.xml"))
    )
    assert "MACRO" in text


# ---------------- get_pic_map ----------------

def test_get_pic_map(tmp_path):
    import struct
    import zlib

    def chunk(t, d):
        return (
            struct.pack(">I", len(d)) + t + d
            + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", 2, 2, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * 2 for _ in range(2))
    blob = (
        b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    )
    p = tmp_path / "logo.png"
    p.write_bytes(blob)
    from docxtplrs import InlineImage

    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ img }}"))))
    # pic_map is populated by the replace_pic pre-processing scan
    tpl.replace_pic("logo.png", io.BytesIO(blob))
    tpl.render({"img": InlineImage(tpl, str(p))})
    out = io.BytesIO()
    tpl.save(out)
    pic_map = tpl.get_pic_map()
    assert isinstance(pic_map, dict)
    assert any("logo.png" in k for k in pic_map)
