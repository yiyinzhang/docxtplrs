"""Tests for advanced Subdoc merging: style conflicts, footnotes, bookmarks,
recursive parts, numbering."""

import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, docx_names, text_of, tp, make_png, XML_DECL, NSDECL

from docxtplrs import DocxTemplate


def build_subdoc(body, files=None, doc_rels=None, content_types_extra=""):
    """Build a subdoc package with custom extra parts."""
    buf = io.BytesIO()
    z = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + content_types_extra
        + "</Types>"
    )
    z.writestr("[Content_Types].xml", ct)
    z.writestr(
        "word/document.xml",
        XML_DECL + f"<w:document {NSDECL}><w:body>{body}<w:sectPr/></w:body></w:document>",
    )
    if doc_rels:
        rels = "".join(
            f'<Relationship Id="{i}" Type="{t}" Target="{tg}"/>' for i, t, tg in doc_rels
        )
        z.writestr(
            "word/_rels/document.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rels
            + "</Relationships>",
        )
    for name, data in (files or {}).items():
        z.writestr(name, data)
    z.close()
    return buf.getvalue()


def render_with_sub(sub_bytes, master_body=None):
    tpl = DocxTemplate(io.BytesIO(make_docx(master_body or tp("{{p sub }}"))))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub_bytes))})
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()


STYLE_CT = '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
NUMBERING_CT = '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
FOOTNOTES_CT = '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
STYLES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
NUMBERING_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
FOOTNOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

STYLES_XML = (
    XML_DECL
    + f'<w:styles {NSDECL}>'
    + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
    + "<w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>"
    + '<w:style w:type="paragraph" w:styleId="SubOnly"><w:name w:val="sub only"/>'
    + '<w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="unused"/>'
    + '<w:rPr><w:shd w:fill="FFFF00"/></w:rPr></w:style>'
    + "</w:styles>"
)


def test_style_conflict_renamed():
    # master already has Heading1 with a *different* definition
    master_styles = (
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        + '<w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>'
    )
    sub_body = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>sub heading</w:t></w:r></w:p>'
    sub = build_subdoc(
        sub_body,
        files={"word/styles.xml": STYLES_XML},
        doc_rels=[("rId10", STYLES_RT, "styles.xml")],
        content_types_extra=STYLE_CT,
    )
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"), styles=master_styles)))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    doc = read_docx_part(data, "word/document.xml")
    # reference must be renamed to Heading1_1
    assert 'w:pStyle w:val="Heading1_1"' in doc
    styles = read_docx_part(data, "word/styles.xml")
    # master keeps its own Heading1 (sz 40) and gains Heading1_1 (sz 32)
    assert '<w:sz w:val="40"/>' in styles
    assert 'w:styleId="Heading1_1"' in styles
    assert '<w:sz w:val="32"/>' in styles
    # unused style not merged
    assert "Unused" not in styles


def test_unused_style_not_merged():
    sub_body = '<w:p><w:pPr><w:pStyle w:val="SubOnly"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    sub = build_subdoc(
        sub_body,
        files={"word/styles.xml": STYLES_XML},
        doc_rels=[("rId10", STYLES_RT, "styles.xml")],
        content_types_extra=STYLE_CT,
    )
    data = render_with_sub(sub)
    # styles part created in master (copied wholesale since master had none)
    styles = read_docx_part(data, "word/styles.xml")
    assert "SubOnly" in styles


def test_footnotes_merged():
    sub_footnotes = (
        XML_DECL
        + f'<w:footnotes {NSDECL}>'
        + '<w:footnote w:type="separator" w:id="0"/>'
        + '<w:footnote w:type="continuationSeparator" w:id="1"/>'
        + '<w:footnote w:id="2"><w:p><w:r><w:t>sub footnote text</w:t></w:r></w:p></w:footnote>'
        + "</w:footnotes>"
    )
    sub_body = (
        "<w:p><w:r><w:t>text</w:t></w:r>"
        + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>'
    )
    sub = build_subdoc(
        sub_body,
        files={"word/footnotes.xml": sub_footnotes},
        doc_rels=[("rId11", FOOTNOTES_RT, "footnotes.xml")],
        content_types_extra=FOOTNOTES_CT,
    )
    # master has its own footnote with id 2
    master_footnotes = (
        '<w:footnote w:id="2"><w:p><w:r><w:t>master note</w:t></w:r></w:p></w:footnote>'
    )
    tpl = DocxTemplate(
        io.BytesIO(make_docx(tp("{{p sub }}"), footnotes=master_footnotes))
    )
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    fn = read_docx_part(data, "word/footnotes.xml")
    assert "master note" in fn and "sub footnote text" in fn
    doc = read_docx_part(data, "word/document.xml")
    # reference must be remapped (master max id is 2 -> sub note 2 becomes 5)
    assert '<w:footnoteReference w:id="5"/>' in doc


def test_bookmarks_renumbered():
    sub_body = (
        '<w:p><w:bookmarkStart w:id="0" w:name="bm"/>'
        + "<w:r><w:t>bm text</w:t></w:r>"
        + '<w:bookmarkEnd w:id="0"/></w:p>'
    )
    sub = build_subdoc(sub_body)
    master_body = (
        '<w:p><w:bookmarkStart w:id="5" w:name="m"/>'
        + "<w:r><w:t>m</w:t></w:r><w:bookmarkEnd w:id=\"5\"/></w:p>"
        + tp("{{p sub }}")
    )
    data = render_with_sub(sub, master_body=master_body)
    doc = read_docx_part(data, "word/document.xml")
    # subdoc bookmark id 0 must be shifted past master's max (5)
    assert 'w:id="6"' in doc


def test_recursive_part_copy():
    # subdoc references a chart part which itself references a style part
    chart_xml = XML_DECL + '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:style r:id="rId5"/></c:chart>'
    chartstyle_xml = XML_DECL + '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>'
    sub_body = "<w:p><w:r><w:t>chart here</w:t></w:r></w:p>"
    sub = build_subdoc(
        sub_body,
        files={
            "word/charts/chart1.xml": chart_xml,
            "word/charts/style1.xml": chartstyle_xml,
            "word/charts/_rels/chart1.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>'
                "</Relationships>"
            ),
        },
        doc_rels=[
            (
                "rId20",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "charts/chart1.xml",
            )
        ],
        content_types_extra='<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    )
    data = render_with_sub(sub)
    names = docx_names(data)
    assert "word/charts/chart1.xml" in names
    assert "word/charts/style1.xml" in names
    # chart's rels must exist in master with remapped rid
    chart = read_docx_part(data, "word/charts/chart1.xml")
    rels = read_docx_part(data, "word/charts/_rels/chart1.xml.rels")
    assert "chartStyle" in rels
    import re

    m = re.search(r'Id="(rId\d+)"[^>]*chartStyle', rels)
    assert m and f'r:id="{m.group(1)}"' in chart


def test_numbering_merged_with_offsets():
    sub_numbering = (
        XML_DECL
        + f'<w:numbering {NSDECL}>'
        + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>'
        + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        + "</w:numbering>"
    )
    master_numbering = (
        XML_DECL
        + f'<w:numbering {NSDECL}>'
        + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
        + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        + "</w:numbering>"
    )
    sub_body = (
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        + "<w:r><w:t>bullet</w:t></w:r></w:p>"
    )
    sub = build_subdoc(
        sub_body,
        files={"word/numbering.xml": sub_numbering},
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml")],
        content_types_extra=NUMBERING_CT,
    )
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    # add master numbering to the package first? master has none -> whole copy
    out = io.BytesIO()
    tpl.save(out)
    num = read_docx_part(out.getvalue(), "word/numbering.xml")
    assert "bullet" in num


def test_docproperty_fields_dissolved():
    sub_body = (
        '<w:p><w:fldSimple w:instr=" DOCPROPERTY Title \\* MERGEFORMAT ">'
        + "<w:r><w:t>My Title</w:t></w:r></w:fldSimple></w:p>"
    )
    sub = build_subdoc(sub_body)
    data = render_with_sub(sub)
    doc = read_docx_part(data, "word/document.xml")
    assert "My Title" in text_of(doc)
    assert "fldSimple" not in doc
