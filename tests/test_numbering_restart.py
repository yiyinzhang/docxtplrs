"""Test docxcompose-style numbering restart for subdocs."""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, tp, XML_DECL, NSDECL

from docxtplrs import DocxTemplate
from test_subdoc_merge import build_subdoc, NUMBERING_CT, NUMBERING_RT, STYLE_CT, STYLES_RT

SUB_NUMBERING = (
    XML_DECL
    + f'<w:numbering {NSDECL}>'
    + '<w:abstractNum w:abstractNumId="0">'
    + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
    + "</w:abstractNum>"
    + '<w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>'
    + "</w:numbering>"
)

MASTER_NUMBERING = (
    XML_DECL
    + f'<w:numbering {NSDECL}>'
    + '<w:abstractNum w:abstractNumId="0">'
    + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
    + "</w:abstractNum>"
    + '<w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>'
    + "</w:numbering>"
)

SUB_STYLES = (
    XML_DECL
    + f'<w:styles {NSDECL}>'
    + '<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/></w:style>'
    + "</w:styles>"
)

MASTER_STYLES = '<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/></w:style>'


def numbered_para(num_id, style="ListPara", text="item"):
    return (
        f'<w:p><w:pPr><w:pStyle w:val="{style}"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
        f"<w:r><w:t>{text}</w:t></w:r></w:p>"
    )


def test_numbering_restarts_for_subdoc():
    sub_body = numbered_para(7, text="first of sub")
    sub = build_subdoc(
        sub_body,
        files={"word/numbering.xml": SUB_NUMBERING, "word/styles.xml": SUB_STYLES},
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml"), ("rId13", STYLES_RT, "styles.xml")],
        content_types_extra=NUMBERING_CT + STYLE_CT,
    )
    tpl = DocxTemplate(
        io.BytesIO(
            make_docx(tp("{{p sub }}"), numbering=MASTER_NUMBERING, styles=MASTER_STYLES)
        )
    )
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    doc = read_docx_part(data, "word/document.xml")
    numbering = read_docx_part(data, "word/numbering.xml")
    # subdoc's num 7 is remapped to 8 by the merge; restart creates a new num 9
    # with startOverride and retargets the paragraph
    assert '<w:startOverride w:val="1"/>' in numbering
    assert numbering.count("<w:num ") >= 3
    # the paragraph must reference the restart num, not the plain merged one
    import re

    m = re.search(r'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="(\d+)"/></w:numPr>', doc)
    assert m
    restart_id = m.group(1)
    assert f'<w:num w:numId="{restart_id}">' in numbering
    # and that num must contain the startOverride
    seg = numbering.split(f'<w:num w:numId="{restart_id}">')[1]
    assert "<w:startOverride" in seg.split("</w:num>")[0]


def test_bullet_not_restarted():
    sub_numbering = SUB_NUMBERING.replace('w:numFmt w:val="decimal"', 'w:numFmt w:val="bullet"')
    master_numbering = MASTER_NUMBERING.replace('w:numFmt w:val="decimal"', 'w:numFmt w:val="bullet"')
    sub_body = numbered_para(7, text="bullet item")
    sub = build_subdoc(
        sub_body,
        files={"word/numbering.xml": sub_numbering, "word/styles.xml": SUB_STYLES},
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml"), ("rId13", STYLES_RT, "styles.xml")],
        content_types_extra=NUMBERING_CT + STYLE_CT,
    )
    tpl = DocxTemplate(
        io.BytesIO(make_docx(tp("{{p sub }}"), numbering=master_numbering, styles=MASTER_STYLES))
    )
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    numbering = read_docx_part(out.getvalue(), "word/numbering.xml")
    assert "startOverride" not in numbering
