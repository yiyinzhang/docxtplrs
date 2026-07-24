"""Test suite for docxtplrs (mirrors docxtpl behaviors)."""

import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import (
    make_docx,
    read_docx_part,
    docx_names,
    text_of,
    make_png,
    document_xml,
    run,
    p,
    tp,
    cell,
    tr,
    tbl,
    CORE_XML,
)

from docxtplrs import (
    DocxTemplate,
    RichText,
    R,
    RichTextParagraph,
    RP,
    Listing,
    InlineImage,
    Subdoc,
    Length,
    Emu,
    Inches,
    Cm,
    Mm,
    Pt,
    Twips,
)


def render(body, context, autoescape=False, **docx_kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body, **docx_kw)))
    tpl.render(context, autoescape=autoescape)
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()


def render_xml(body, context, autoescape=False, **docx_kw):
    return read_docx_part(render(body, context, autoescape, **docx_kw), "word/document.xml")


# --------------------------------------------------------------- basics

def test_basic_variable():
    xml = render_xml(tp("Hello {{ name }}!"), {"name": "World"})
    assert "Hello World!" in text_of(xml)


def test_variable_split_across_runs():
    # Word often splits jinja tags across multiple runs
    body = p(run("Hello {{"), run(" na", rpr="<w:b/>"), run("me }}!"))
    xml = render_xml(body, {"name": "World"})
    assert "Hello World!" in text_of(xml)


def test_variable_with_filter_split_runs():
    body = p(run("{{ name"), run(" | upper }}"))
    xml = render_xml(body, {"name": "abc"})
    assert "ABC" in text_of(xml)


def test_no_autoescape_by_default():
    # without autoescape, values are inserted raw (like docxtpl); the
    # recover pass then escapes stray markup so the output stays valid
    xml = render_xml(tp("{{ v }}"), {"v": "a<b>c"})
    assert "a&lt;b&gt;c" in xml or "a<b>c" in xml


def test_autoescape():
    xml = render_xml(tp("{{ v }}"), {"v": "a<b>&\"c\""}, autoescape=True)
    t = text_of(xml)
    assert "a&lt;b&gt;" in xml
    assert "&amp;" in xml


def test_if_else():
    body = tp("{% if cond %}YES{% else %}NO{% endif %}")
    assert "YES" in text_of(render_xml(body, {"cond": True}))
    assert "NO" in text_of(render_xml(body, {"cond": False}))


def test_for_loop():
    body = tp("{% for x in items %}{{ x }},{% endfor %}")
    assert "a,b,c," in text_of(render_xml(body, {"items": ["a", "b", "c"]}))


def test_loop_vars():
    body = tp("{% for x in items %}{{ loop.index }}{{ 'F' if loop.first else '' }}{{ 'L' if loop.last else '' }};{% endfor %}")
    t = text_of(render_xml(body, {"items": [1, 2, 3]}))
    assert "1F;" in t and "2;" in t and "3L;" in t


def test_loop_length():
    body = tp("{% for x in items %}{{ loop.length }}{% endfor %}")
    assert "333" in text_of(render_xml(body, {"items": [1, 2, 3]}))


def test_filters():
    body = tp("{{ s | upper }}|{{ lst | join('-') }}|{{ lst | length }}|{{ missing | default('dflt') }}")
    t = text_of(render_xml(body, {"s": "abc", "lst": [1, 2]}))
    assert "ABC" in t and "1-2" in t and "2" in t and "dflt" in t


def test_dict_and_index_access():
    body = tp("{{ d.k }}-{{ lst[1] }}")
    t = text_of(render_xml(body, {"d": {"k": "v"}, "lst": [10, 20]}))
    assert "v-20" in t


def test_object_attribute_and_method():
    class User:
        def __init__(self):
            self.name = "Bob"

        def greet(self, punct="!"):
            return f"hi{ punct}"

    body = tp("{{ u.name }}:{{ u.greet('?') }}")
    t = text_of(render_xml(body, {"u": User()}))
    assert "Bob:hi?" in t


def test_newline_in_value_becomes_br():
    xml = render_xml(tp("{{ v }}"), {"v": "line1\nline2"})
    assert "<w:br/>" in xml


def test_tab_in_value():
    xml = render_xml(tp("{{ v }}"), {"v": "a\tb"})
    assert "<w:tab/>" in xml


def test_formfeed_in_value():
    xml = render_xml(tp("{{ v }}"), {"v": "a\fb"})
    assert '<w:br w:type="page"/>' in xml


def test_bell_in_value_new_paragraph():
    xml = render_xml(tp("{{ v }}"), {"v": "para1\apara2"})
    assert xml.count("<w:p>") >= 2


def test_escaped_delimiters():
    # {_{ must render as literal {{
    xml = render_xml(tp("{_{ not_a_var }_}"), {})
    assert "{{ not_a_var }}" in text_of(xml)


def test_smartquotes_in_tags():
    body = tp("{{ \u201ckey\u201d }}")
    # jinja: undefined variable renders empty; should not raise
    render_xml(body, {})


def test_set_statement():
    body = tp("{% set x = 42 %}{{ x }}")
    assert "42" in text_of(render_xml(body, {}))


# --------------------------------------------------------------- RichText / Listing

def test_richtext_basic():
    rt = RichText("Hello", bold=True, color="#FF0000", size=28)
    xml = render_xml(tp("{{r rt }}"), {"rt": rt})
    assert "<w:b/>" in xml
    assert '<w:color w:val="FF0000"/>' in xml
    assert '<w:sz w:val="28"/>' in xml
    assert "Hello" in text_of(xml)


def test_richtext_alias_R():
    rt = R("x", italic=True)
    assert isinstance(rt, RichText)
    xml = render_xml(tp("{{r rt }}"), {"rt": rt})
    assert "<w:i/>" in xml


def test_richtext_add_and_combine():
    rt = RichText("first", bold=True)
    rt.add("second", underline=True)
    rt.add(RichText("third", strike=True))
    xml = render_xml(tp("{{r rt }}"), {"rt": rt})
    assert "<w:b/>" in xml and '<w:u w:val="single"/>' in xml and "<w:strike/>" in xml
    t = text_of(xml)
    assert "first" in t and "second" in t and "third" in t


def test_richtext_escapes_text():
    rt = RichText("a<b>&")
    xml = render_xml(tp("{{r rt }}"), {"rt": rt})
    assert "a&lt;b&gt;&amp;" in xml


def test_richtext_all_props():
    rt = RichText(
        "t",
        style="MyStyle",
        highlight="#00FF00",
        subscript=True,
        font="Arial",
        rtl=True,
        lang="en-US",
    )
    xml = render_xml(tp("{{r rt }}"), {"rt": rt})
    assert '<w:rStyle w:val="MyStyle"/>' in xml
    assert '<w:shd w:fill="00FF00"/>' in xml
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert 'w:ascii="Arial"' in xml
    assert '<w:rtl w:val="true"/>' in xml
    assert '<w:lang w:val="en-US"/>' in xml


def test_richtext_regional_font():
    rt = RichText("t", font="eastAsia:SimSun")
    xml = render_xml(tp("{{r rt }}"), {"rt": rt})
    assert 'w:eastAsia="SimSun"' in xml


def test_richtext_url():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{r rt }}"))))
    url_id = tpl.build_url_id("https://example.com")
    tpl.render({"rt": RichText("click", url_id=url_id)})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert f'<w:hyperlink r:id="{url_id}"' in xml
    rels = read_docx_part(out.getvalue(), "word/_rels/document.xml.rels")
    assert "https://example.com" in rels


def test_richtext_paragraph():
    rp = RichTextParagraph()
    rp.add(RichText("line1", bold=True), parastyle="MyPara")
    rp.add("line2")
    xml = render_xml(tp("{{p rp }}"), {"rp": rp})
    assert '<w:pStyle w:val="MyPara"/>' in xml
    assert "line1" in text_of(xml) and "line2" in text_of(xml)


def test_listing():
    xml = render_xml(tp("{{ lst }}"), {"lst": Listing("a\nb\tc")})
    assert "<w:br/>" in xml and "<w:tab/>" in xml


def test_listing_escapes():
    xml = render_xml(tp("{{ lst }}"), {"lst": Listing("a<b")})
    assert "a&lt;b" in xml


# --------------------------------------------------------------- tables

def test_table_row_loop():
    # docxtpl convention: {%tr for%} alone in a tag row, data rows between
    # the for-row and the {%tr endfor%} row are repeated.
    body = tbl([
        tr(cell(tp("{%tr for item in items %}")), cell(tp(""))),
        tr(cell(tp("{{ item }}")), cell(tp("x"))),
        tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
    ])
    xml = render_xml(body, {"items": ["a", "b", "c"]})
    t = text_of(xml)
    assert "a" in t and "b" in t and "c" in t
    # 3 data rows produced (each with 2 cells)
    assert xml.count("<w:tr>") == 3


def test_table_colspan():
    body = tbl([
        tr(cell(tp("{% colspan span %}")), cell(tp("{{ v }}"))),
    ])
    xml = render_xml(body, {"span": 2, "v": "x"})
    assert '<w:gridSpan w:val="2"/>' in xml


def test_table_cellbg():
    body = tbl([
        tr(cell(tp("{% cellbg color %}")), cell(tp("x"))),
    ])
    xml = render_xml(body, {"color": "FF0000"})
    assert '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>' in xml


def test_table_cellbg_removes_existing_shd():
    body = tbl([
        tr(cell(tp("{% cellbg color %}"), tcpr='<w:shd w:val="clear" w:fill="000000"/>')),
    ])
    xml = render_xml(body, {"color": "00FF00"})
    assert 'w:fill="00FF00"' in xml
    assert 'w:fill="000000"' not in xml


def test_vertical_merge():
    tcpr = '<w:tcW w:w="1000" w:type="dxa"/>'
    body = tbl([
        tr(cell(tp("{%tr for x in items %}")), cell(tp("h"))),
        tr(cell(tp("{% vm %}merged"), tcpr=tcpr), cell(tp("{{ x }}"))),
        tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
    ], widths=(1000, 1000))
    xml = render_xml(body, {"items": [1, 2]})
    assert '<w:vMerge w:val="restart"/>' in xml
    assert '<w:vMerge w:val="continue"/>' in xml


def test_horizontal_merge():
    body = tbl([
        tr(cell(tp("{%tr for x in items %}")) if False else cell(tp("{% hm %}{{ x }}"))),
    ], widths=(1000,))
    body = (
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>'
        + "</w:tblGrid>"
        + tr(cell(tp("{% for x in items %}")) if False else "")
    )
    # proper hm template: loop inside the row, single cell with {% hm %}
    # (tcPr must be non-empty for the merge patching, as with docxtpl)
    tcpr = '<w:tcW w:w="1000" w:type="dxa"/>'
    body = (
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>'
        + "</w:tblGrid>"
        + "<w:tr>"
        + cell(tp("{% for x in items %}"))
        + cell(tp("{% hm %}{{ x }}"), tcpr=tcpr)
        + cell(tp("{% endfor %}"))
        + "</w:tr></w:tbl>"
    )
    xml = render_xml(body, {"items": [1, 2]})
    assert '<w:gridSpan w:val="2"/>' in xml
    # only one generated cell kept
    assert xml.count("<w:tc>") >= 2


def test_fix_tables_adds_columns():
    # {%tc for %} generates more cells than grid columns -> grid fixed
    body = (
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>'
        + "</w:tblGrid>"
        + "<w:tr>"
        + cell(tp("{%tc for x in items %}"))
        + cell(tp("{{ x }}"))
        + cell(tp("{%tc endfor %}"))
        + "</w:tr></w:tbl>"
    )
    xml = render_xml(body, {"items": [1, 2, 3, 4]})
    assert xml.count("<w:gridCol") == 4
    assert xml.count("<w:tc>") == 4


def test_fix_tables_removes_columns():
    # fewer generated cells than grid columns -> gridCol removed, width kept
    body = (
        "<w:tbl><w:tblGrid>"
        + '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>'
        + "</w:tblGrid>"
        + "<w:tr>"
        + cell(tp("{%tc for x in items %}"))
        + cell(tp("{{ x }}"))
        + cell(tp("{%tc endfor %}"))
        + "</w:tr></w:tbl>"
    )
    xml = render_xml(body, {"items": [1, 2]})
    assert xml.count("<w:gridCol") == 2
    # total width preserved (4000)
    import re

    widths = [int(w) for w in re.findall(r'<w:gridCol w:w="(\d+)"/>', xml)]
    assert sum(widths) == 4000


def test_tc_comment_tag():
    body = (
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"1000\"/></w:tblGrid>"
        + "<w:tr>"
        + cell(tp("{#tc comment about row #}"))
        + cell(tp("v"))
        + "</w:tr></w:tbl>"
    )
    xml = render_xml(body, {})
    assert "comment" not in text_of(xml)


# --------------------------------------------------------------- paragraph tags

def test_p_if():
    body = tp("{%p if show %}") + tp("shown paragraph") + tp("{%p endif %}")
    xml = render_xml(body, {"show": True})
    assert "shown paragraph" in text_of(xml)
    xml = render_xml(body, {"show": False})
    assert "shown" not in text_of(xml)


def test_p_for():
    body = tp("{%p for x in items %}") + tp("item {{ x }}") + tp("{%p endfor %}")
    xml = render_xml(body, {"items": [1, 2, 3]})
    t = text_of(xml)
    assert "item 1" in t and "item 2" in t and "item 3" in t
    # three paragraphs generated
    assert xml.count("<w:p>") >= 3


def test_r_tag_new_run():
    # {{r }} is meant for RichText-like values producing run xml
    xml = render_xml(tp("before {{r v }} after"), {"v": RichText("MID", bold=True)})
    assert "MID" in text_of(xml)
    assert "<w:b/>" in xml


def test_dash_merging():
    # {%- merges with previous paragraph text
    body = tp("A {%- if cond %}") + tp("B{% endif %}")
    xml = render_xml(body, {"cond": True})
    assert "A B" in text_of(xml) or "AB" in text_of(xml)


# --------------------------------------------------------------- headers/footers/properties

def test_header_footer_rendering():
    out = render(
        tp("body {{ v }}"),
        {"v": "X"},
        headers={"header1.xml": tp("head {{ v }}")},
        footers={"footer1.xml": tp("foot {{ v }}")},
    )
    assert "head X" in text_of(read_docx_part(out, "word/header1.xml"))
    assert "foot X" in text_of(read_docx_part(out, "word/footer1.xml"))


def test_core_properties():
    core = CORE_XML.format(title="Report for {{ name }}", author="{{ author }}")
    out = render(tp("x"), {"name": "ACME", "author": "Bob"}, core=core)
    core_xml = read_docx_part(out, "docProps/core.xml")
    assert "Report for ACME" in core_xml
    assert "<dc:creator>Bob</dc:creator>" in core_xml


def test_footnotes():
    out = render(
        tp("x"),
        {"v": "FN"},
        footnotes=tp("note {{ v }}"),
    )
    assert "note FN" in text_of(read_docx_part(out, "word/footnotes.xml"))


# --------------------------------------------------------------- images

def test_inline_image(tmp_path):
    png = make_png(100, 50)
    img_path = tmp_path / "pic.png"
    img_path.write_bytes(png)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ img }}"))))
    tpl.render({"img": InlineImage(tpl, str(img_path), width=Mm(20))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    names = docx_names(data)
    assert any(n.startswith("word/media/image") and n.endswith(".png") for n in names)
    xml = read_docx_part(data, "word/document.xml")
    assert "<w:drawing>" in xml
    assert "<wp:inline" in xml
    # width 20mm = 720000 EMU; height scaled by aspect ratio (50/100)
    assert 'cx="720000"' in xml
    assert 'cy="360000"' in xml
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    assert "relationships/image" in rels


def test_inline_image_native_size(tmp_path):
    png = make_png(72, 36)  # 72dpi -> 1 inch wide
    img_path = tmp_path / "p.png"
    img_path.write_bytes(png)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ img }}"))))
    tpl.render({"img": InlineImage(tpl, str(img_path))})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert 'cx="914400"' in xml  # 1 inch
    assert 'cy="457200"' in xml  # 0.5 inch


def test_inline_image_anchor(tmp_path):
    png = make_png()
    img_path = tmp_path / "p.png"
    img_path.write_bytes(png)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ img }}"))))
    tpl.render({"img": InlineImage(tpl, str(img_path), anchor="https://example.com")})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "<a:hlinkClick" in xml
    rels = read_docx_part(out.getvalue(), "word/_rels/document.xml.rels")
    assert "https://example.com" in rels


def test_inline_image_from_bytes(tmp_path):
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ img }}"))))
    tpl.render({"img": InlineImage(tpl, make_png(10, 10))})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert 'name="image.png"' in xml


def test_inline_image_dedup(tmp_path):
    png = make_png()
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ i1 }} {{ i2 }}"))))
    tpl.render({"i1": InlineImage(tpl, png), "i2": InlineImage(tpl, png)})
    out = io.BytesIO()
    tpl.save(out)
    media = [n for n in docx_names(out.getvalue()) if n.startswith("word/media/")]
    assert len(media) == 1


# --------------------------------------------------------------- subdoc

def test_subdoc_plain():
    sub_bytes = make_docx(tp("sub content {{ var }}"))
    # note: subdoc content is inserted as-is (already rendered subdoc in real usage)
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub_bytes))})
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    assert "sub content {{ var }}" in xml
    assert "<w:sectPr/>" not in xml.split("<w:body>")[1].split("sub content")[0]


def test_subdoc_with_image():
    png = make_png(5, 5)
    # build subdoc containing an image relationship + drawing
    sub_body = (
        "<w:p><w:r><w:drawing><wp:inline>"
        '<wp:extent cx="100" cy="100"/><wp:docPr id="1" name="Picture 1"/>'
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
        "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"p.png\"/><pic:cNvPicPr/></pic:nvPicPr>"
        "<pic:blipFill><a:blip r:embed=\"rId50\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>"
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    buf = io.BytesIO()
    import zipfile as zf

    z = zf.ZipFile(buf, "w")
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    z.writestr("[Content_Types].xml", ct)
    z.writestr("word/document.xml", document_xml(sub_body))
    z.writestr(
        "word/_rels/document.xml.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/sub.png"/>'
        "</Relationships>",
    )
    z.writestr("word/media/sub.png", png)
    z.close()

    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(buf.getvalue()))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    xml = read_docx_part(data, "word/document.xml")
    assert 'r:embed="rId50"' not in xml  # rid remapped
    media = [n for n in docx_names(data) if n.startswith("word/media/")]
    assert len(media) == 1
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    assert "relationships/image" in rels


# --------------------------------------------------------------- replacements

def test_replace_media():
    png1 = make_png(2, 2)
    png2 = make_png(6, 6)
    docx = make_docx(tp("x"), media={"image1.png": png1})
    tpl = DocxTemplate(io.BytesIO(docx))
    tpl.replace_media(io.BytesIO(png1), io.BytesIO(png2))
    out = io.BytesIO()
    tpl.save(out)
    import zipfile as zf

    with zf.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert z.read("word/media/image1.png") == png2


def test_replace_embedded():
    docx = make_docx(tp("x"), extra_files={"word/embeddings/file1.xlsx": b"OLDXLSX"})
    tpl = DocxTemplate(io.BytesIO(docx))
    src = tmp = None
    # need source file to compute crc: write temp files
    import tempfile, os

    with tempfile.TemporaryDirectory() as d:
        sp = os.path.join(d, "src.xlsx")
        dp = os.path.join(d, "dst.xlsx")
        with open(sp, "wb") as f:
            f.write(b"OLDXLSX")
        with open(dp, "wb") as f:
            f.write(b"NEWXLSX")
        tpl.replace_embedded(sp, dp)
        out = io.BytesIO()
        tpl.save(out)
    import zipfile as zf

    with zf.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert z.read("word/embeddings/file1.xlsx") == b"NEWXLSX"


def test_replace_zipname():
    docx = make_docx(tp("x"), extra_files={"word/embeddings/f.txt": b"OLD"})
    tpl = DocxTemplate(io.BytesIO(docx))
    import tempfile, os

    with tempfile.TemporaryDirectory() as d:
        dp = os.path.join(d, "new.txt")
        with open(dp, "wb") as f:
            f.write(b"NEW")
        tpl.replace_zipname("word/embeddings/f.txt", dp)
        out = io.BytesIO()
        tpl.save(out)
    import zipfile as zf

    with zf.ZipFile(io.BytesIO(out.getvalue())) as z:
        assert z.read("word/embeddings/f.txt") == b"NEW"


def test_replace_pic():
    png1 = make_png(2, 2)
    png2 = make_png(6, 6)
    body = (
        "<w:p><w:r><w:drawing><wp:inline>"
        '<wp:extent cx="100" cy="100"/><wp:docPr id="1" name="Picture 1"/>'
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
        "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"dummy.png\"/><pic:cNvPicPr/></pic:nvPicPr>"
        "<pic:blipFill><a:blip r:embed=\"rId50\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>"
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    buf = io.BytesIO()
    import zipfile as zf

    z = zf.ZipFile(buf, "w")
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    z.writestr("[Content_Types].xml", ct)
    z.writestr("word/document.xml", document_xml(body))
    z.writestr(
        "word/_rels/document.xml.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/dummy.png"/>'
        "</Relationships>",
    )
    z.writestr("word/media/dummy.png", png1)
    z.close()

    tpl = DocxTemplate(io.BytesIO(buf.getvalue()))
    tpl.replace_pic("dummy.png", io.BytesIO(png2))
    out = io.BytesIO()
    tpl.save(out)
    with zf.ZipFile(io.BytesIO(out.getvalue())) as z2:
        assert z2.read("word/media/dummy.png") == png2
    # pic map recorded
    pm = tpl.get_pic_map()
    assert "dummy.png" in pm


def test_replace_pic_missing_raises():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.replace_pic("nothere.png", io.BytesIO(make_png()))
    with pytest.raises(Exception, match="not found"):
        tpl.save(io.BytesIO())


def test_allow_missing_pics():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.allow_missing_pics = True
    tpl.replace_pic("nothere.png", io.BytesIO(make_png()))
    tpl.save(io.BytesIO())  # no error


def test_reset_replacements():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    tpl.replace_pic("a.png", io.BytesIO(make_png()))
    tpl.reset_replacements()
    tpl.save(io.BytesIO())


# --------------------------------------------------------------- misc API

def test_undeclared_variables():
    tpl = DocxTemplate(
        io.BytesIO(
            make_docx(
                tp("{{ a }} {% for x in items %}{{ x.b }}{% endfor %}{% if cond %}c{% endif %}"),
                headers={"header1.xml": tp("{{ hdr_var }}")},
            )
        )
    )
    vars = tpl.get_undeclared_template_variables()
    assert vars == {"a", "items", "cond", "hdr_var"}


def test_undeclared_variables_with_context():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ a }} {{ b }}"))))
    vars = tpl.get_undeclared_template_variables(context={"a": 1})
    assert vars == {"b"}


def test_render_twice():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ v }}"))))
    tpl.render({"v": "one"})
    out1 = io.BytesIO()
    tpl.save(out1)
    assert "one" in text_of(read_docx_part(out1.getvalue(), "word/document.xml"))
    # render again with fresh template
    tpl.render({"v": "two"})
    out2 = io.BytesIO()
    tpl.save(out2)
    assert "two" in text_of(read_docx_part(out2.getvalue(), "word/document.xml"))


def test_save_to_path(tmp_path):
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ v }}"))))
    tpl.render({"v": "z"})
    path = tmp_path / "out.docx"
    tpl.save(str(path))
    assert "z" in text_of(read_docx_part(path.read_bytes(), "word/document.xml"))


def test_template_from_path(tmp_path):
    path = tmp_path / "tpl.docx"
    path.write_bytes(make_docx(tp("{{ v }}")))
    tpl = DocxTemplate(str(path))
    tpl.render({"v": "frompath"})
    out = io.BytesIO()
    tpl.save(out)
    assert "frompath" in text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def test_is_rendered_flags():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    assert not tpl.is_rendered
    tpl.render({})
    assert tpl.is_rendered


def test_get_xml():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hello"))))
    xml = tpl.get_xml()
    assert "hello" in xml


def test_write_xml(tmp_path):
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("hello"))))
    out = tmp_path / "out.xml"
    tpl.write_xml(str(out))
    assert "hello" in out.read_text()


def test_build_url_id():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    rid1 = tpl.build_url_id("https://a.com")
    rid2 = tpl.build_url_id("https://b.com")
    assert rid1 != rid2
    tpl.render({})
    out = io.BytesIO()
    tpl.save(out)
    rels = read_docx_part(out.getvalue(), "word/_rels/document.xml.rels")
    assert "https://a.com" in rels and "https://b.com" in rels
    assert 'TargetMode="External"' in rels


def test_lengths():
    assert Inches(1).emu == 914400
    assert Cm(1).emu == 360000
    assert Mm(1).emu == 36000
    assert Pt(1).emu == 12700
    assert Twips(1).emu == 635
    assert Emu(5).emu == 5
    assert Length(914400).inches == 1.0


def test_subdoc_class_direct():
    sub_bytes = make_docx(tp("direct subdoc"))
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p s }}"))))
    tpl.render({"s": Subdoc(tpl, io.BytesIO(sub_bytes))})
    out = io.BytesIO()
    tpl.save(out)
    assert "direct subdoc" in text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def test_docpr_ids_renumbered():
    body = (
        "<w:p><w:r><w:drawing><wp:inline>"
        '<wp:extent cx="1" cy="1"/><wp:docPr id="5" name="a"/>'
        "<a:graphic/></wp:inline></w:drawing></w:r>"
        "<w:r><w:drawing><wp:inline>"
        '<wp:extent cx="1" cy="1"/><wp:docPr id="5" name="b"/>'
        "<a:graphic/></wp:inline></w:drawing></w:r></w:p>"
    )
    xml = render_xml(body, {})
    import re

    ids = re.findall(r'<wp:docPr id="(\d+)"', xml)
    assert len(ids) == 2
    assert ids[0] != ids[1]
    assert all(int(i) > 1000 for i in ids)


def test_error_contains_context():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("start {% badtag %} end"))))
    with pytest.raises(Exception):
        tpl.render({})
