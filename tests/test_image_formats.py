"""Coverage for image.rs: JPEG/GIF/BMP/TIFF/PNG-pHYs size & DPI parsing.

Exercised through InlineImage rendering; assertions check the computed
wp:extent EMU values (EMU = px * 914400 / dpi).
"""
import io
import struct
import sys
import zipfile
import zlib

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx  # noqa: E402

from docxtplrs import DocxTemplate, InlineImage, Mm, TemplateError  # noqa: E402

EMU = 914400


def render_img(blob: bytes, **kw) -> str:
    tpl = DocxTemplate(io.BytesIO(make_docx("<w:p><w:r><w:t>{{ img }}</w:t></w:r></w:p>")))
    tpl.render({"img": InlineImage(tpl, io.BytesIO(blob), **kw)})
    out = io.BytesIO()
    tpl.save(out)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        return z.read("word/document.xml").decode()


def extent(xml: str):
    import re

    m = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', xml)
    return int(m.group(1)), int(m.group(2))


def jpeg(dpi_units=1, xden=200, yden=200, w=100, h=50, jfif=True) -> bytes:
    app0 = b""
    if jfif:
        app0 = (
            b"\xff\xe0\x00\x10JFIF\x00\x01\x02"
            + bytes([dpi_units])
            + struct.pack(">HH", xden, yden)
            + b"\x00\x00"
        )
    sof = b"\xff\xc0\x00\x11\x08" + struct.pack(">HH", h, w) + b"\x03" + b"\x00" * 9
    return b"\xff\xd8" + app0 + sof + b"\xff\xd9"


def gif(w=100, h=50) -> bytes:
    return b"GIF89a" + struct.pack("<HH", w, h) + b"\x00" * 8


def bmp(w=100, h=50, xppm=0, yppm=0, core=False) -> bytes:
    if core:
        dib = struct.pack("<IHHHH", 12, w, h, 1, 24)  # BITMAPCOREHEADER
    else:
        dib = struct.pack("<IiiHHIIii", 40, w, h, 1, 24, 0, 0, xppm, yppm)
        dib += b"\x00" * 8  # clrUsed/clrImportant
    return b"BM" + struct.pack("<III", 14 + len(dib), 0, 14 + len(dib)) + dib


def tiff(little=True, w=100, h=50, num=300, den=1, unit=2, with_res=True) -> bytes:
    end = "<" if little else ">"
    magic = b"II*\x00" if little else b"MM\x00*"
    entries = [
        struct.pack(end + "HHII", 256, 4, 1, w),
        struct.pack(end + "HHII", 257, 4, 1, h),
    ]
    data = b""
    if with_res:
        base = 8 + 2 + 5 * 12 + 4
        entries += [
            struct.pack(end + "HHII", 282, 5, 1, base),
            struct.pack(end + "HHII", 283, 5, 1, base + 8),
            struct.pack(end + "HHIH", 296, 3, 1, unit) + b"\x00\x00",
        ]
        data = struct.pack(end + "II", num, den) + struct.pack(end + "II", num, den)
    ifd = struct.pack(end + "H", len(entries)) + b"".join(entries) + struct.pack(end + "I", 0)
    return magic + struct.pack(end + "I", 8) + ifd + data


def png_with_phys(w=100, h=50, ppux=0, ppuy=0) -> bytes:
    def chunk(t, d):
        return (
            struct.pack(">I", len(d)) + t + d
            + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    phys = struct.pack(">IIB", ppux, ppuy, 1) if ppux else b""
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * w for _ in range(h))
    out = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
    if phys:
        out += chunk(b"pHYs", phys)
    return out + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


# ---------------- format parsing ----------------

def test_jpeg_jfif_dpi_inch():
    cx, cy = extent(render_img(jpeg(dpi_units=1, xden=200, yden=100)))
    assert (cx, cy) == (EMU * 100 // 200, EMU * 50 // 100)


def test_jpeg_jfif_dpi_cm():
    # units=2 -> per-cm density * 2.54
    cx, cy = extent(render_img(jpeg(dpi_units=2, xden=100, yden=100)))
    assert (cx, cy) == (EMU * 100 // 254, EMU * 50 // 254)


def test_jpeg_no_jfif_defaults_72dpi():
    cx, cy = extent(render_img(jpeg(jfif=False)))
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_jpeg_no_sof_raises():
    with pytest.raises(Exception):
        render_img(b"\xff\xd8\xff\xd9")


def test_gif():
    cx, cy = extent(render_img(gif()))
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_gif87a():
    cx, cy = extent(render_img(b"GIF87a" + struct.pack("<HH", 100, 50) + b"\x00" * 8))
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_bmp_infoheader_dpi():
    # 3937 px/m == 100 dpi
    cx, cy = extent(render_img(bmp(xppm=3937, yppm=3937)))
    assert (cx, cy) == (EMU * 100 // 100, EMU * 50 // 100)


def test_bmp_infoheader_no_dpi():
    cx, cy = extent(render_img(bmp()))
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_bmp_coreheader():
    cx, cy = extent(render_img(bmp(core=True)))
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_bmp_top_down_height():
    cx, cy = extent(render_img(bmp(h=-50)))  # negative height = top-down
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_tiff_little_endian():
    cx, cy = extent(render_img(tiff(little=True)))
    assert (cx, cy) == (EMU * 100 // 300, EMU * 50 // 300)


def test_tiff_big_endian():
    cx, cy = extent(render_img(tiff(little=False)))
    assert (cx, cy) == (EMU * 100 // 300, EMU * 50 // 300)


def test_tiff_res_unit_cm():
    # unit=3 -> per-cm * 2.54 ; 100/cm -> 254 dpi
    cx, cy = extent(render_img(tiff(num=100, den=1, unit=3)))
    assert (cx, cy) == (EMU * 100 // 254, EMU * 50 // 254)


def test_tiff_no_resolution_defaults_72():
    cx, cy = extent(render_img(tiff(with_res=False)))
    assert (cx, cy) == (EMU * 100 // 72, EMU * 50 // 72)


def test_png_phys_dpi():
    # 3937 px/m == 100 dpi
    cx, cy = extent(render_img(png_with_phys(ppux=3937, ppuy=3937)))
    assert (cx, cy) == (EMU * 100 // 100, EMU * 50 // 100)


def test_unrecognized_format_raises():
    with pytest.raises(Exception):
        render_img(b"\x00\x01\x02\x03garbage")


# ---------------- scaled_dimensions paths ----------------

def test_width_and_height_explicit():
    xml = render_img(gif(), width=Mm(40), height=Mm(20))
    cx, cy = extent(xml)
    assert (cx, cy) == (Mm(40).emu, Mm(20).emu)


def test_height_only_scales_proportionally():
    # native: 100x50 @72dpi -> 1270000 x 635000 EMU; height=Mm(25)=900000
    cx, cy = extent(render_img(gif(), height=Mm(25)))
    assert cy == Mm(25).emu
    assert cx == round(1270000 * (Mm(25).emu / 635000))
