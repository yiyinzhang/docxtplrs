"""Cross-validate docxtplrs against the reference docxtpl implementation.

Run with:
  python3.14 tests/crosscheck.py docxtpl   > /tmp/ref.json     (needs docxtpl on PYTHONPATH)
  .venv/bin/python tests/crosscheck.py docxtplrs > /tmp/rs.json
  python3 tests/crosscheck.py compare /tmp/ref.json /tmp/rs.json
"""

import io
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, run, p, tp, cell, tr, tbl

TCPR = '<w:tcW w:w="1000" w:type="dxa"/>'


def cases():
    """Each case: (name, body, context_fn(engine_module), kwargs)"""
    out = []

    def add(name, body, ctx=None, autoescape=False, **kw):
        out.append((name, body, ctx or {}, autoescape, kw))

    add("basic", tp("Hello {{ name }}!"), {"name": "World"})
    add("split_runs", p(run("Hello {{"), run(" na", rpr="<w:b/>"), run("me }}!")), {"name": "World"})
    add("for_loop", tp("{% for x in items %}{{ x }}{{ ',' if not loop.last }}{% endfor %}"), {"items": [1, 2, 3]})
    add("if_else", tp("{% if c %}Y{% else %}N{% endif %}"), {"c": True})
    add("filters", tp("{{ s|upper }} {{ l|join('+') }} {{ l|length }}"), {"s": "ab", "l": [1, 2]})
    add("autoescape", tp("{{ v }}"), {"v": "a<b>&"}, autoescape=True)
    add("newline", tp("{{ v }}"), {"v": "l1\nl2"})
    add("tab", tp("{{ v }}"), {"v": "a\tb"})
    add("p_if", tp("{%p if show %}") + tp("keep {{ v }}") + tp("{%p endif %}"), {"show": True, "v": 1})
    add("p_for", tp("{%p for x in items %}") + tp("i{{ x }}") + tp("{%p endfor %}"), {"items": [1, 2]})
    add(
        "tr_loop",
        tbl([
            tr(cell(tp("{%tr for item in items %}")), cell(tp(""))),
            tr(cell(tp("{{ item }}")), cell(tp("x"))),
            tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
        ]),
        {"items": ["a", "b", "c"]},
    )
    add("colspan", tbl([tr(cell(tp("{% colspan s %}")), cell(tp("{{ v }}")))]), {"s": 2, "v": "x"})
    add("cellbg", tbl([tr(cell(tp("{% cellbg c %}")), cell(tp("x")))]), {"c": "FF0000"})
    add(
        "vmerge",
        tbl([
            tr(cell(tp("{%tr for x in items %}")), cell(tp("h"))),
            tr(cell(tp("{% vm %}m"), tcpr=TCPR), cell(tp("{{ x }}"))),
            tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
        ]),
        {"items": [1, 2]},
    )
    add(
        "hmerge",
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>'
        + "</w:tblGrid><w:tr>"
        + cell(tp("{% for x in items %}"))
        + cell(tp("{% hm %}{{ x }}"), tcpr=TCPR)
        + cell(tp("{% endfor %}"))
        + "</w:tr></w:tbl>",
        {"items": [1, 2]},
    )
    add(
        "tc_loop_fix",
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>'
        + "</w:tblGrid><w:tr>"
        + cell(tp("{%tc for x in items %}"))
        + cell(tp("{{ x }}"))
        + cell(tp("{%tc endfor %}"))
        + "</w:tr></w:tbl>",
        {"items": [1, 2, 3, 4]},
    )
    add("header_footer", tp("body {{ v }}"), {"v": "X"},
        headers={"header1.xml": tp("head {{ v }}")}, footers={"footer1.xml": tp("foot {{ v }}")})
    add("dash_merge", tp("A {%- if c %}") + tp("B{% endif %}"), {"c": True})
    add("smartquote", tp("{{ \u201ck\u201d }}"), {})
    add("escaped_delim", tp("{_{ x }_}"), {})
    return out


def features(data: bytes) -> dict:
    doc = read_docx_part(data, "word/document.xml")
    names = []
    try:
        hdr = read_docx_part(data, "word/header1.xml")
        ftr = read_docx_part(data, "word/footer1.xml")
    except KeyError:
        hdr = ftr = ""
    return {
        "text": text_of(doc),
        "n_tr": doc.count("<w:tr"),
        "n_tc": doc.count("<w:tc>") + doc.count("<w:tc "),
        "n_gridcol": doc.count("<w:gridCol"),
        "gridspan_vals": sorted(re.findall(r'<w:gridSpan w:val="([^"]+)"', doc)),
        "shd_fills": sorted(re.findall(r'<w:shd[^>]*w:fill="([^"]+)"', doc)),
        "vmerge": sorted(re.findall(r'<w:vMerge w:val="([^"]+)"', doc)),
        "n_br": doc.count("<w:br/>"),
        "n_tab": doc.count("<w:tab/>"),
        "n_p": len(re.findall(r"<w:p>", doc)),
        "hdr_text": text_of(hdr),
        "ftr_text": text_of(ftr),
        "runs": sorted(re.findall(r"<w:r><w:rPr>.{0,400}?</w:r>", doc)),
        "extent": sorted(re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc)),
        "n_drawing": doc.count("<w:drawing>"),
        "hyperlinks": sorted(re.findall(r'<w:hyperlink[^>]*>', doc)),
    }


def extra_cases(engine, lib):
    """Cases needing engine-specific objects (RichText, Listing, InlineImage)."""
    import helpers

    png_path = "/tmp/xchk_img.png"
    with open(png_path, "wb") as f:
        f.write(helpers.make_png(100, 50))

    if engine == "docxtpl":
        from docx.shared import Mm
    else:
        Mm = lib.Mm

    out = []

    def add(name, body, ctx_fn, autoescape=False, **kw):
        out.append((name, body, ctx_fn, autoescape, kw))

    add("richtext", tp("{{r rt }}"),
        lambda tpl: {"rt": lib.RichText("Hi", bold=True, color="#FF0000", size=28)})
    add("richtext2", tp("{{r rt }}"),
        lambda tpl: {"rt": lib.RichText("u<l>", underline=True, italic=True, font="Arial")})
    add("richtext_multi", tp("{{r rt }}"),
        lambda tpl: {"rt": (lambda rt: (rt.add("b", strike=True), rt)[1])(lib.RichText("a", bold=True))})
    add("richtext_url", tp("{{r rt }}"),
        lambda tpl: {"rt": lib.RichText("click", url_id=tpl.build_url_id("https://example.com"))})
    add("listing", tp("{{ lst }}"), lambda tpl: {"lst": lib.Listing("a\nb\tc")})
    add("listing_bell", tp("{{ lst }}"), lambda tpl: {"lst": lib.Listing("p1\ap2")})
    add("inline_image", tp("{{ img }}"),
        lambda tpl: {"img": lib.InlineImage(tpl, png_path, width=Mm(20))})
    add("inline_image_native", tp("{{ img }}"),
        lambda tpl: {"img": lib.InlineImage(tpl, png_path)})
    return out


def run_engine(engine):
    if engine == "docxtpl":
        import docxtpl as lib
    else:
        import docxtplrs as lib

    all_cases = [(n, b, (lambda t, c=c: c), a, k) for n, b, c, a, k in cases()]
    all_cases += extra_cases(engine, lib)

    results = {}
    for name, body, ctx_fn, autoescape, kw in all_cases:
        try:
            tpl = lib.DocxTemplate(io.BytesIO(make_docx(body, **kw)))
            ctx = ctx_fn(tpl)
            tpl.render(ctx, autoescape=autoescape)
            out = io.BytesIO()
            tpl.save(out)
            results[name] = features(out.getvalue())
        except Exception as e:
            results[name] = {"error": f"{type(e).__name__}: {e}"}
    return results


def compare(ref_path, rs_path):
    ref = json.load(open(ref_path))
    rs = json.load(open(rs_path))
    ok = True
    for name in ref:
        a, b = ref[name], rs.get(name)
        if a != b:
            ok = False
            print(f"=== DIFF in case '{name}' ===")
            keys = set(a) | set(b or {})
            for k in sorted(keys):
                va, vb = a.get(k), (b or {}).get(k)
                if va != vb:
                    print(f"  {k}:\n    docxtpl:   {va!r}\n    docxtplrs: {vb!r}")
    if ok:
        print("ALL CASES MATCH")
    return 0 if ok else 1


if __name__ == "__main__":
    if sys.argv[1] == "compare":
        sys.exit(compare(sys.argv[2], sys.argv[3]))
    else:
        json.dump(run_engine(sys.argv[1]), sys.stdout, indent=1, sort_keys=True)
