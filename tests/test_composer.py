"""Tests for the docxcompose-style Composer API (document concatenation)."""

import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, docx_names, tp, XML_DECL, NSDECL

from docxtplrs import Composer
from test_subdoc_merge import (
    build_subdoc,
    NUMBERING_CT,
    NUMBERING_RT,
    STYLE_CT,
    STYLES_RT,
    STYLES_XML,
)
from test_numbering_restart import (
    SUB_NUMBERING,
    MASTER_NUMBERING,
    SUB_STYLES,
    MASTER_STYLES,
    numbered_para,
)


def compose(master_bytes, *subs):
    c = Composer(io.BytesIO(master_bytes))
    for s in subs:
        c.append(io.BytesIO(s))
    out = io.BytesIO()
    c.save(out)
    return out.getvalue()


def test_append_order_and_page_breaks():
    master = make_docx(tp("master"))
    sub1 = make_docx(tp("one"))
    sub2 = make_docx(tp("two"))
    doc = read_docx_part(compose(master, sub1, sub2), "word/document.xml")

    i_master = doc.index(">master<")
    i_br1 = doc.index("<w:br")
    i_one = doc.index(">one<")
    i_br2 = doc.rindex("<w:br")
    i_two = doc.index(">two<")
    i_sectpr = doc.index("<w:sectPr")
    # master -> page break -> sub1 -> page break -> sub2, sectPr stays last
    assert i_master < i_br1 < i_one < i_br2 < i_two < i_sectpr
    assert doc.count('w:type="page"') == 2


def test_save_to_path(tmp_path):
    master = make_docx(tp("master"))
    sub = make_docx(tp("sub"))
    c = Composer(io.BytesIO(master))
    c.append(io.BytesIO(sub))
    out_path = tmp_path / "out.docx"
    c.save(str(out_path))
    doc = read_docx_part(out_path.read_bytes(), "word/document.xml")
    assert ">master<" in doc and ">sub<" in doc


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
    data = compose(make_docx(tp("master"), styles=master_styles), sub)
    doc = read_docx_part(data, "word/document.xml")
    # reference must be renamed to Heading1_1
    assert 'w:pStyle w:val="Heading1_1"' in doc
    styles = read_docx_part(data, "word/styles.xml")
    # master keeps its own Heading1 (sz 40) and gains Heading1_1 (sz 32)
    assert 'w:styleId="Heading1"' in styles
    assert 'w:styleId="Heading1_1"' in styles


def test_numbering_restarts_for_appended_doc():
    sub_body = numbered_para(7, text="first of sub")
    sub = build_subdoc(
        sub_body,
        files={"word/numbering.xml": SUB_NUMBERING, "word/styles.xml": SUB_STYLES},
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml"), ("rId13", STYLES_RT, "styles.xml")],
        content_types_extra=NUMBERING_CT + STYLE_CT,
    )
    master = make_docx(
        numbered_para(7, text="master item"),
        numbering=MASTER_NUMBERING,
        styles=MASTER_STYLES,
    )
    data = compose(master, sub)
    doc = read_docx_part(data, "word/document.xml")
    numbering = read_docx_part(data, "word/numbering.xml")
    # subdoc's list is remapped and restarted with a startOverride num
    assert '<w:startOverride w:val="1"/>' in numbering
    # two numbered paragraphs: master keeps numId 7, sub references a new num
    ids = re.findall(r'<w:numId w:val="(\d+)"/>', doc)
    assert len(ids) == 2
    assert ids[0] == "7"
    restart_id = ids[1]
    seg = numbering.split(f'<w:num w:numId="{restart_id}">')[1]
    assert "<w:startOverride" in seg.split("</w:num>")[0]


def test_append_with_image():
    png = (
        bytes.fromhex("89504e470d0a1a0a")  # PNG signature is enough for the merge
    )
    image_rt = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    sub_body = (
        '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    sub = build_subdoc(
        sub_body,
        files={"word/media/image1.png": png},
        doc_rels=[("rId20", image_rt, "media/image1.png")],
        content_types_extra='<Default Extension="png" ContentType="image/png"/>',
    )
    data = compose(make_docx(tp("master")), sub)
    names = docx_names(data)
    assert "word/media/image1.png" in names
    doc = read_docx_part(data, "word/document.xml")
    # rId remapped to a master rel that points at the copied media part
    m = re.search(r'r:embed="(rId\d+)"', doc)
    assert m
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    assert f'Id="{m.group(1)}"' in rels
    assert "media/image1.png" in rels
