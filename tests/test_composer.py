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
    # master already has Heading1 with a *different* definition; with
    # preserve_styles=True conflicting styles are kept under a renamed id
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
    master = make_docx(tp("master"), styles=master_styles)
    c = Composer(io.BytesIO(master), preserve_styles=True)
    assert c.preserve_styles is True
    c.append(io.BytesIO(sub))
    out = io.BytesIO()
    c.save(out)
    data = out.getvalue()
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


# ---------------- insert() / append_index() ----------------

def test_insert_at_middle_of_body():
    master = make_docx(tp("a") + tp("b") + tp("c"))
    sub = make_docx(tp("X"))
    c = Composer(io.BytesIO(master))
    # the append index is the master body's trailing sectPr
    assert c.append_index() == 3
    c.insert(1, io.BytesIO(sub))
    out = io.BytesIO()
    c.save(out)
    doc = read_docx_part(out.getvalue(), "word/document.xml")
    i_a = doc.index(">a<")
    i_br = doc.index("<w:br")
    i_x = doc.index(">X<")
    i_b = doc.index(">b<")
    i_c = doc.index(">c<")
    # inserted between a and b, preceded by a page break
    assert i_a < i_br < i_x < i_b < i_c


# ---------------- remove_property_fields ----------------

CUSTOM_PROPS_CT = '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'

CUSTOM_XML = (
    XML_DECL
    + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    + 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    + '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Firma">'
    + "<vt:lpwstr>ACME</vt:lpwstr></property>"
    + '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Ort">'
    + "<vt:lpwstr>Bern</vt:lpwstr></property>"
    + "</Properties>"
)

# one fldSimple field and one complex (fldChar) field, both DOCPROPERTY
PROPERTY_FIELDS_BODY = (
    '<w:p><w:fldSimple w:instr=" DOCPROPERTY &quot;Firma&quot; \\* MERGEFORMAT ">'
    + "<w:r><w:t>stale-simple</w:t></w:r></w:fldSimple></w:p>"
    + "<w:p>"
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> DOCPROPERTY "Ort" \\* MERGEFORMAT </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + "<w:r><w:t>stale-complex</w:t></w:r>"
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    + "</w:p>"
)


def build_property_fields_sub():
    return build_subdoc(
        PROPERTY_FIELDS_BODY,
        files={"docProps/custom.xml": CUSTOM_XML},
        content_types_extra=CUSTOM_PROPS_CT,
    )


def test_remove_property_fields_dissolves_both_field_kinds():
    # remove_property_fields=True (the default) dissolves fldSimple and
    # complex DOCPROPERTY fields into the custom property's literal value
    c = Composer(io.BytesIO(make_docx(tp("master"))))
    c.append(io.BytesIO(build_property_fields_sub()))
    out = io.BytesIO()
    c.save(out)
    doc = read_docx_part(out.getvalue(), "word/document.xml")
    assert ">ACME<" in doc and ">Bern<" in doc
    assert "stale" not in doc
    assert "DOCPROPERTY" not in doc
    assert "fldSimple" not in doc
    assert "fldChar" not in doc


def test_remove_property_fields_false_keeps_fields():
    c = Composer(io.BytesIO(make_docx(tp("master"))))
    c.append(io.BytesIO(build_property_fields_sub()), remove_property_fields=False)
    out = io.BytesIO()
    c.save(out)
    doc = read_docx_part(out.getvalue(), "word/document.xml")
    assert doc.count("DOCPROPERTY") == 2
    assert "stale-simple" in doc and "stale-complex" in doc
    assert "fldSimple" in doc and "fldChar" in doc


# ---------------- preserve_styles ----------------

def test_preserve_styles_default_is_false():
    assert Composer(io.BytesIO(make_docx(tp("m")))).preserve_styles is False


def test_preserve_styles_false_maps_conflict_by_name():
    # default: a style whose name exists in the master is mapped to the
    # master's same-named style (no copy, no X_1 rename)
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
    # the reference points at the master's Heading1, unchanged
    assert 'w:pStyle w:val="Heading1"' in doc
    assert "Heading1_1" not in doc
    styles = read_docx_part(data, "word/styles.xml")
    # no renamed copy of the sub's style
    assert "Heading1_1" not in styles
    assert '<w:sz w:val="32"/>' not in styles


def test_preserve_styles_false_maps_different_ids_by_name():
    # the sub's style B is named "Foo", which exists in the master as A:
    # the reference is remapped to A and B is not copied; the unknown
    # style C is still merged
    master_styles = (
        '<w:style w:type="paragraph" w:styleId="A"><w:name w:val="Foo"/>'
        + "<w:rPr><w:b/></w:rPr></w:style>"
    )
    sub_styles = (
        XML_DECL
        + f"<w:styles {NSDECL}>"
        + '<w:style w:type="paragraph" w:styleId="B"><w:name w:val="Foo"/>'
        + "<w:rPr><w:i/></w:rPr></w:style>"
        + '<w:style w:type="paragraph" w:styleId="C"><w:name w:val="Bar"/></w:style>'
        + "</w:styles>"
    )
    sub_body = (
        '<w:p><w:pPr><w:pStyle w:val="B"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>'
        + '<w:p><w:pPr><w:pStyle w:val="C"/></w:pPr><w:r><w:t>other</w:t></w:r></w:p>'
    )
    sub = build_subdoc(
        sub_body,
        files={"word/styles.xml": sub_styles},
        doc_rels=[("rId10", STYLES_RT, "styles.xml")],
        content_types_extra=STYLE_CT,
    )
    data = compose(make_docx(tp("master"), styles=master_styles), sub)
    doc = read_docx_part(data, "word/document.xml")
    assert '<w:pStyle w:val="A"/>' in doc  # B mapped to the master's A
    assert '<w:pStyle w:val="B"/>' not in doc
    assert '<w:pStyle w:val="C"/>' in doc
    styles = read_docx_part(data, "word/styles.xml")
    assert 'w:styleId="B"' not in styles
    assert 'w:styleId="C"' in styles


# ---------------- restart_numbering ----------------

def test_restart_numbering_property_roundtrip():
    c = Composer(io.BytesIO(make_docx(tp("m"))))
    assert c.restart_numbering is True  # default
    c.restart_numbering = False
    assert c.restart_numbering is False


def test_restart_numbering_false_skips_restart():
    sub = build_subdoc(
        numbered_para(7, text="first of sub"),
        files={"word/numbering.xml": SUB_NUMBERING, "word/styles.xml": SUB_STYLES},
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml"), ("rId13", STYLES_RT, "styles.xml")],
        content_types_extra=NUMBERING_CT + STYLE_CT,
    )
    master = make_docx(
        numbered_para(7, text="master item"),
        numbering=MASTER_NUMBERING,
        styles=MASTER_STYLES,
    )
    c = Composer(io.BytesIO(master))
    c.restart_numbering = False
    c.append(io.BytesIO(sub))
    out = io.BytesIO()
    c.save(out)
    data = out.getvalue()
    numbering = read_docx_part(data, "word/numbering.xml")
    # no restart: the merged list keeps counting where the master left off
    assert "startOverride" not in numbering
    doc = read_docx_part(data, "word/document.xml")
    # both numbered paragraphs survive, with distinct num ids
    ids = re.findall(r'<w:numId w:val="(\d+)"/>', doc)
    assert len(ids) == 2 and ids[0] != ids[1]


# ---------------- doc facade snapshot ----------------

def test_doc_returns_snapshot_facade():
    c = Composer(io.BytesIO(make_docx(tp("master"))))
    c.append(io.BytesIO(make_docx(tp("one"))))
    doc = c.doc
    texts = [p.text for p in doc.paragraphs]
    assert "master" in texts and "one" in texts
    # snapshot semantics: later appends are not reflected in the old facade
    c.append(io.BytesIO(make_docx(tp("two"))))
    assert "two" not in [p.text for p in doc.paragraphs]
    assert "two" in [p.text for p in c.doc.paragraphs]
