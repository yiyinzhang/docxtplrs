"""Benchmark for listing-heavy templates (values with \n / \t / \x07).

Targets P1-a (resolve_listing hand-rolled scan). Usage:
    .venv/bin/python tests/bench_listing.py
"""

import io
import os
import sys
import time
import zipfile

sys.path.insert(0, os.path.dirname(__file__))
from helpers import document_xml, p, run  # noqa: E402

from docxtplrs import DocxTemplate  # noqa: E402


def make_docx(path: str, n_rows: int = 300) -> None:
    tpl = (
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>"
        "<w:tr><w:tc>" + p(run("{%tr for a in addresses %}"))
        + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + p(run("{{ a }}"))
        + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + p(run("{%tr endfor %}"))
        + "</w:tc></w:tr></w:tbl>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", document_xml(p(run("{{ title }}")) + tpl))


def main() -> None:
    path = "/tmp/bench_listing.docx"
    make_docx(path)
    # multi-line addresses with tabs: every rendered cell hits resolve_listing
    addresses = [
        "Alice Smith\tCEO\n123 Main Street\nSuite 400\x07Springfield, IL 62701"
        for _ in range(300)
    ]
    ctx = {"title": "Directory", "addresses": addresses}

    best = float("inf")
    for _ in range(10):
        t0 = time.perf_counter()
        tpl = DocxTemplate(path)
        tpl.render(ctx)
        tpl.save(io.BytesIO())
        best = min(best, time.perf_counter() - t0)
    print(f"listing-heavy full cycle (300 multi-line cells): {best*1e3:.2f} ms")


if __name__ == "__main__":
    main()
