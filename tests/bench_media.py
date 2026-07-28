"""Benchmark for media-heavy templates and repeated render() calls.

Covers the P0 blind spots of tests/benchmark.py:
- save cost dominated by re-deflating incompressible media (P0-1)
- repeated render() re-inflating the whole zip (P0-2)
- repeated render() re-compiling the jinja template (P0-5)

Usage: .venv/bin/python tests/bench_media.py
"""

import io
import os
import sys
import time
import zipfile

sys.path.insert(0, os.path.dirname(__file__))
from helpers import document_xml, p, run  # noqa: E402

from docxtplrs import DocxTemplate  # noqa: E402


def make_png(size_kb: int) -> bytes:
    """A minimal valid PNG with `size_kb` of incompressible IDAT payload."""
    import struct
    import zlib

    sig = b"\x89PNG\r\n\x1a\n"

    def chunk(typ: bytes, data: bytes) -> bytes:
        c = struct.pack(">I", len(data)) + typ + data
        return c + struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", 2048, 2048, 8, 2, 0, 0, 0)
    raw = os.urandom(size_kb * 1024)  # incompressible
    idat = zlib.compress(raw, 0)  # stored blocks: keeps size, valid PNG
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


def make_media_docx(path: str, n_images: int = 4, img_kb: int = 1024) -> None:
    """docx with a table-row loop template and n_images PNGs in word/media."""
    tpl = (
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>"
        "<w:tr><w:tc>" + p(run("{%tr for u in users %}"))
        + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + p(run("{{ u.name }}"))
        + "</w:tc><w:tc>" + p(run("{{ u.price }}"))
        + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + p(run("{%tr endfor %}"))
        + "</w:tc></w:tr></w:tbl>"
    )
    doc = document_xml(p(run("{{ title }}")) + tpl)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", doc)
        for i in range(n_images):
            z.writestr(f"word/media/image{i}.png", make_png(img_kb))


def timeit(fn, n=5):
    best = float("inf")
    for _ in range(n):
        t0 = time.perf_counter()
        fn()
        best = min(best, time.perf_counter() - t0)
    return best


def main() -> None:
    path = "/tmp/bench_media.docx"
    make_media_docx(path)
    size_mb = os.path.getsize(path) / 1e6
    users = [{"name": f"user{i}", "price": i * 1.5} for i in range(50)]

    def full_cycle():
        tpl = DocxTemplate(path)
        tpl.render({"title": "Report", "users": users})
        tpl.save(io.BytesIO())

    t_full = timeit(full_cycle)

    tpl = DocxTemplate(path)
    ctx = {"title": "Report", "users": users}
    tpl.render(ctx)
    t_rerender = timeit(lambda: tpl.render(ctx))
    t_save = timeit(lambda: tpl.save(io.BytesIO()))

    print(f"template size: {size_mb:.1f} MB (4x1MB PNG)")
    print(f"full cycle (init+render+save): {t_full*1e3:8.2f} ms")
    print(f"repeated render (same object): {t_rerender*1e3:8.2f} ms")
    print(f"save only:                     {t_save*1e3:8.2f} ms")


if __name__ == "__main__":
    main()
