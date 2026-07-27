#!/usr/bin/env python3
"""Engine-vs-engine benchmark: docxtplrs vs python-docx-template (docxtpl).

Two cases (templates generated on the fly, identical for both engines):

- micro: title + `{%p %}` paragraph loop + `{%tr %}` table-row loop over
  10/100/400 items, 30 full cycles (load + render + save to BytesIO) each,
  3 interleaved rounds, best median reported.
- big (--big): ~8.9MB document.xml (20k paragraphs, table loop, nested loop),
  one full cycle per round, best time reported.

Usage:
    .venv/bin/python tests/benchmark.py          # micro only
    .venv/bin/python tests/benchmark.py --big    # micro + large document
"""

import io
import statistics
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import docxtpl
import docxtplrs

ENGINES = [("docxtplrs", docxtplrs), ("docxtpl", docxtpl)]

CT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)
DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)


def write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", document_xml)


def para(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def make_micro_template(path: Path) -> None:
    """title + paragraph loop + {%tr %} table-row loop (README's micro case)."""
    body = [
        para("{{ title }}"),
        para("{%p for x in items %}"),
        para("{{ x.name }} costs {{ x.price }}"),
        para("{%p endfor %}"),
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
        '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
        "<w:tr><w:tc>" + para("name") + "</w:tc><w:tc>" + para("price") + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + para("{%tr for x in items %}") + "</w:tc><w:tc>" + para("") + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + para("{{ x.name }}") + "</w:tc><w:tc>" + para("{{ x.price }}") + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + para("{%tr endfor %}") + "</w:tc><w:tc>" + para("") + "</w:tc></w:tr>"
        "</w:tbl>",
    ]
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>" + "".join(body) + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
    )
    write_docx(path, xml)


def make_big_template(path: Path) -> None:
    """~8.9MB document.xml: 20k paragraphs + table loop + nested loop."""
    filler = "The quick brown fox jumps over the lazy dog. " * 6
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<w:body>"
    ]
    for i in range(20000):
        parts.append(
            '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
            f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{filler} para {i}</w:t></w:r>'
            "<w:r><w:t>{{ user.name }} tail</w:t></w:r></w:p>"
        )
    parts.append(para("{%p for row in rows %}"))
    for _ in range(200):
        parts.append(
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            "<w:tr><w:tc>" + para("{{ row.name }}") + "</w:tc>"
            "<w:tc>" + para("{{ row.value }}") + "</w:tc></w:tr></w:tbl>"
        )
        parts.append(para(f"static {filler}"))
    parts.append(para("{%p endfor %}"))
    parts.append(para("{%p for g in groups %}"))
    parts.append(para("{%p for it in g.entries %}"))
    for _ in range(50):
        parts.append(para(f"{{{{ g.title }}}} - {{{{ it }}}} - {filler}"))
    parts.append(para("{%p endfor %}"))
    parts.append(para("{%p endfor %}"))
    parts.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>')
    write_docx(path, "".join(parts))


def big_ctx() -> dict:
    return {
        "user": {"name": "张三 Zhang"},
        "rows": [{"name": f"rowname-{i}", "value": i * 7} for i in range(10)],
        "groups": [
            {"title": f"group-{g}", "entries": [f"item-{i}" for i in range(10)]}
            for g in range(3)
        ],
    }


def one_cycle(mod, path: str, ctx: dict) -> float:
    t0 = time.perf_counter()
    tpl = mod.DocxTemplate(path)
    tpl.render(ctx)
    bio = io.BytesIO()
    tpl.save(bio)
    return time.perf_counter() - t0


def bench_micro(workdir: Path) -> None:
    tpl_path = workdir / "micro.docx"
    make_micro_template(tpl_path)
    print(f"micro template: {tpl_path.stat().st_size} bytes "
          "(title + {%p %} loop + {%tr %} row loop)")
    for n_items in (10, 100, 400):
        ctx = {
            "title": "Quarterly Report",
            "items": [{"name": f"item-{i}", "price": i * 1.5} for i in range(n_items)],
        }
        best = {}
        for _round in range(3):  # interleaved rounds cancel machine drift
            for name, mod in ENGINES:
                times = [one_cycle(mod, str(tpl_path), ctx) for _ in range(30)]
                med = statistics.median(times)
                best[name] = min(med, best.get(name, med))
        rs, py = best["docxtplrs"], best["docxtpl"]
        print(f"  items={n_items:>3}: docxtplrs {rs*1000:6.2f}ms | docxtpl {py*1000:6.2f}ms "
              f"| speedup {py/rs:5.1f}x")


def bench_big(workdir: Path) -> None:
    tpl_path = workdir / "big.docx"
    make_big_template(tpl_path)
    xml_size = zipfile.ZipFile(tpl_path).getinfo("word/document.xml").file_size
    print(f"big template: document.xml {xml_size/1e6:.1f}MB "
          "(20k paragraphs + table loop + nested loop)")
    ctx = big_ctx()
    best = {}
    for _round in range(2):
        for name, mod in ENGINES:
            t = one_cycle(mod, str(tpl_path), ctx)
            best[name] = min(t, best.get(name, t))
    rs, py = best["docxtplrs"], best["docxtpl"]
    print(f"  full cycle: docxtplrs {rs:.2f}s | docxtpl {py:.2f}s | speedup {py/rs:.1f}x")


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="docxtpl-bench-") as td:
        workdir = Path(td)
        bench_micro(workdir)
        if "--big" in sys.argv:
            bench_big(workdir)


if __name__ == "__main__":
    main()
