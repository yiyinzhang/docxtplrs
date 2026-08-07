"""Subdoc keep_sections (sectPr preservation incl. headers/footers) and
comments merging — both go beyond docxcompose parity."""

import io
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, tp, XML_DECL, NSDECL

from docxtplrs import DocxTemplate, Composer

HEADER_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
COMMENTS_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
HEADER_CT = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
COMMENTS_CT = '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'

PORTRAIT_SECTPR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
LANDSCAPE_SECTPR = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>'


def build_doc(body, sectpr="<w:sectPr/>", files=None, doc_rels=None, content_types_extra=""):
    """Build a docx package with a custom body-level sectPr."""
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
        XML_DECL + f"<w:document {NSDECL}><w:body>{body}{sectpr}</w:body></w:document>",
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


def comment_body(cid, text="annotated"):
    return (
        f"<w:p><w:commentRangeStart w:id=\"{cid}\"/>"
        f"<w:r><w:t>{text}</w:t></w:r>"
        f"<w:commentRangeEnd w:id=\"{cid}\"/>"
        f'<w:r><w:commentReference w:id="{cid}"/></w:r></w:p>'
    )


def comments_xml(*comments):
    """comments: list of (id, author, text)"""
    inner = "".join(
        f'<w:comment w:id="{i}" w:author="{a}"><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:comment>'
        for i, a, t in comments
    )
    return XML_DECL + f"<w:comments {NSDECL}>{inner}</w:comments>"


def render_with_sub(sub_bytes, master_body=None, master_sectpr="<w:sectPr/>", **kw):
    tpl = DocxTemplate(io.BytesIO(build_doc(master_body or tp("{{p sub }}"), sectpr=master_sectpr)))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub_bytes), **kw)})
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()


# ---------------- keep_sections ----------------


def test_keep_sections_preserves_page_setup():
    sub = build_doc(tp("sub content"), sectpr=LANDSCAPE_SECTPR)
    doc = read_docx_part(
        render_with_sub(sub, master_sectpr=PORTRAIT_SECTPR, keep_sections=True),
        "word/document.xml",
    )
    # subdoc sectPr became a paragraph-level sectPr right after its content
    assert '<w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr>' in doc
    # master's own (non-empty) sectPr was cloned into a resume break *before*
    # the subdoc content, so master content keeps its page setup
    assert '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>' in doc
    resume_pos = doc.index('w:w="11906" w:h="16838"/></w:sectPr></w:pPr>')
    sub_pos = doc.index("sub content")
    land_pos = doc.index('orient="landscape"')
    assert resume_pos < sub_pos < land_pos
    # master's body-level sectPr still comes last
    assert doc.rindex("<w:sectPr>") > land_pos


def test_keep_sections_default_drops_sectpr():
    sub = build_doc(tp("sub content"), sectpr=LANDSCAPE_SECTPR)
    doc = read_docx_part(render_with_sub(sub), "word/document.xml")
    assert "landscape" not in doc


def test_keep_sections_empty_sectpr_ignored():
    # an empty sectPr carries no page setup: nothing is inserted
    sub = build_doc(tp("sub content"))
    doc = read_docx_part(render_with_sub(sub, keep_sections=True), "word/document.xml")
    assert doc.count("<w:p>") == 1  # only the sub content paragraph
    assert doc.count("<w:sectPr") == 1  # only the master's (empty) one


def test_keep_sections_merges_header_with_rename():
    # master already has header1.xml -> subdoc's header must be renamed
    master = build_doc(
        tp("{{p sub }}"),
        sectpr=PORTRAIT_SECTPR,
        files={"word/header1.xml": XML_DECL + f"<w:hdr {NSDECL}>" + tp("master header") + "</w:hdr>"},
        doc_rels=[("rId5", HEADER_RT, "header1.xml")],
        content_types_extra=HEADER_CT,
    )
    sub_sectpr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/>'
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>'
    )
    sub = build_doc(
        tp("sub content"),
        sectpr=sub_sectpr,
        files={"word/header1.xml": XML_DECL + f"<w:hdr {NSDECL}>" + tp("sub header") + "</w:hdr>"},
        doc_rels=[("rId7", HEADER_RT, "header1.xml")],
        content_types_extra=HEADER_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub), keep_sections=True)})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    doc = read_docx_part(data, "word/document.xml")
    # the preserved sectPr references a header rel in the master
    m = re.search(r'<w:headerReference w:type="default" r:id="(rId\d+)"/>', doc)
    assert m
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    rm = re.search(rf'Id="{m.group(1)}"[^>]*Target="([^"]+)"', rels)
    assert rm
    target = rm.group(1)
    assert target != "header1.xml"  # renamed: master keeps its own header1
    sub_header = read_docx_part(data, f"word/{target}")
    assert "sub header" in sub_header
    # master's own header is untouched
    assert "master header" in read_docx_part(data, "word/header1.xml")


# ---------------- comments merge ----------------


def test_comments_merged_with_offset():
    master = make_docx(
        comment_body(0, "master text") + tp("{{p sub }}"),
        extra_files={"word/comments.xml": comments_xml((0, "Ann", "master note"))},
    )
    sub = build_doc(
        comment_body(0, "sub text"),
        files={"word/comments.xml": comments_xml((0, "Bob", "sub note"))},
        doc_rels=[("rId9", COMMENTS_RT, "comments.xml")],
        content_types_extra=COMMENTS_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    comments = read_docx_part(data, "word/comments.xml")
    assert 'w:id="0"' in comments and "master note" in comments
    assert 'w:id="1"' in comments and "sub note" in comments
    doc = read_docx_part(data, "word/document.xml")
    # master's references keep id 0, the subdoc's were remapped to 1
    assert '<w:commentReference w:id="0"/>' in doc
    assert '<w:commentReference w:id="1"/>' in doc
    assert '<w:commentRangeStart w:id="1"/>' in doc
    assert '<w:commentRangeEnd w:id="1"/>' in doc


def test_comments_copied_when_master_has_none():
    master = make_docx(tp("{{p sub }}"))
    sub = build_doc(
        comment_body(0, "sub text"),
        files={"word/comments.xml": comments_xml((0, "Bob", "sub note"))},
        doc_rels=[("rId9", COMMENTS_RT, "comments.xml")],
        content_types_extra=COMMENTS_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    comments = read_docx_part(data, "word/comments.xml")
    assert "sub note" in comments and 'w:id="0"' in comments
    doc = read_docx_part(data, "word/document.xml")
    assert '<w:commentReference w:id="0"/>' in doc
    # relationship + content type registered
    rels = read_docx_part(data, "word/_rels/document.xml.rels")
    assert "comments.xml" in rels


# ---------------- Composer + sections ----------------


def test_composer_preserves_sections_and_skips_page_break():
    master = build_doc(tp("master content"), sectpr=PORTRAIT_SECTPR)
    sub_sectpr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/>'
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>'
    )
    sub = build_doc(
        tp("sub content"),
        sectpr=sub_sectpr,
        files={"word/header1.xml": XML_DECL + f"<w:hdr {NSDECL}>" + tp("sub header") + "</w:hdr>"},
        doc_rels=[("rId7", HEADER_RT, "header1.xml")],
        content_types_extra=HEADER_CT,
    )
    c = Composer(io.BytesIO(master))
    c.append(io.BytesIO(sub))
    out = io.BytesIO()
    c.save(out)
    data = out.getvalue()

    doc = read_docx_part(data, "word/document.xml")
    # landscape section preserved as paragraph-level sectPr
    assert 'orient="landscape"' in doc
    # no explicit page break: the section break starts the new page
    assert 'w:type="page"' not in doc
    # master content -> resume break (portrait) -> sub content -> landscape break
    assert doc.index("master content") < doc.index('w:w="11906"')
    assert doc.index("sub content") < doc.index('orient="landscape"')
    # master's body-level sectPr still comes last
    assert doc.rindex("<w:sectPr>") > doc.index('orient="landscape"')
    # header part copied (no collision here: master has none)
    assert "sub header" in read_docx_part(data, "word/header1.xml")


# ---------------- consecutive appends / multi-section / hf styles ----------------

from test_subdoc_merge import STYLES_XML  # noqa: E402
from test_numbering_restart import (  # noqa: E402
    MASTER_NUMBERING,
    SUB_NUMBERING,
    SUB_STYLES,
    MASTER_STYLES,
    numbered_para,
)

NUMBERING_CT = '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
NUMBERING_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
STYLES_CT = '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
STYLES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
COMMENTS_EXTENDED_CT = '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>'
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def test_composer_consecutive_sections_no_blank_page():
    """Second append must not add a resume break: the previous append already
    ended with a section break, a resume break here would be an empty section
    (a blank page)."""
    master = build_doc(tp("master content"), sectpr=PORTRAIT_SECTPR)
    sub1 = build_doc(tp("chapter one"), sectpr=LANDSCAPE_SECTPR)
    sub2 = build_doc(tp("chapter two"), sectpr=LANDSCAPE_SECTPR)
    c = Composer(io.BytesIO(master))
    c.append(io.BytesIO(sub1))
    c.append(io.BytesIO(sub2))
    out = io.BytesIO()
    c.save(out)
    doc = read_docx_part(out.getvalue(), "word/document.xml")

    # exactly one resume break (portrait clone), before chapter one
    assert doc.count('<w:pPr><w:sectPr><w:pgSz w:w="11906"') == 1
    # two preserved landscape section breaks
    assert doc.count('w:orient="landscape"') == 2
    # order: master -> resume -> ch1 -> landscape -> ch2 -> landscape -> body sectPr
    assert (
        doc.index("master content")
        < doc.index("chapter one")
        < doc.index("chapter two")
    )
    assert doc.rindex("<w:sectPr>") > doc.rindex('orient="landscape"')


def test_multi_section_subdoc_preserved():
    """A subdoc with an intermediate (paragraph-level) section keeps both
    section setups."""
    mid_sectpr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/><w:cols w:num="2"/></w:sectPr>'
    sub = build_doc(
        tp("first section") + f"<w:p><w:pPr>{mid_sectpr}</w:pPr></w:p>" + tp("second section"),
        sectpr=LANDSCAPE_SECTPR,
    )
    doc = read_docx_part(
        render_with_sub(sub, master_sectpr=PORTRAIT_SECTPR, keep_sections=True),
        "word/document.xml",
    )
    assert '<w:cols w:num="2"/>' in doc  # intermediate section kept
    assert 'orient="landscape"' in doc  # final section kept
    assert doc.index("first section") < doc.index('<w:cols w:num="2"/>') < doc.index(
        "second section"
    ) < doc.index('orient="landscape"')


def _header_with(ppr_inner):
    return (
        XML_DECL
        + f"<w:hdr {NSDECL}>"
        + f"<w:p><w:pPr>{ppr_inner}</w:pPr><w:r><w:t>styled header</w:t></w:r></w:p>"
        + "</w:hdr>"
    )


def test_header_styles_merged_consistently():
    """A style referenced from both the body and a header maps to ONE new id."""
    master_styles = (
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        + '<w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>'
    )
    sub_sectpr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/>'
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>'
    )
    sub = build_doc(
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>sub heading</w:t></w:r></w:p>',
        sectpr=sub_sectpr,
        files={
            "word/styles.xml": STYLES_XML,
            "word/header1.xml": _header_with('<w:pStyle w:val="Heading1"/>'),
        },
        doc_rels=[("rId10", STYLES_RT, "styles.xml"), ("rId7", HEADER_RT, "header1.xml")],
        content_types_extra=STYLES_CT + HEADER_CT,
    )
    # master has its own conflicting Heading1 (sz 40)
    master = build_doc(tp("{{p sub }}"), files={
        "word/styles.xml": XML_DECL + f"<w:styles {NSDECL}>" + master_styles + "</w:styles>",
    }, doc_rels=[("rId3", STYLES_RT, "styles.xml")], content_types_extra=STYLES_CT)
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub), keep_sections=True)})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    styles = read_docx_part(data, "word/styles.xml")
    # exactly one renamed copy
    assert styles.count('w:styleId="Heading1_1"') == 1
    doc = read_docx_part(data, "word/document.xml")
    assert 'w:pStyle w:val="Heading1_1"' in doc
    header = read_docx_part(data, "word/header1.xml")
    # the header references the SAME renamed style (not a duplicate)
    assert 'w:pStyle w:val="Heading1_1"' in header


def test_header_numbering_merged_consistently():
    """A numId referenced from both the body and a header maps to ONE new id;
    the docxcompose numbering restart still only touches the body."""
    sub_sectpr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/>'
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>'
    )
    # body paragraph carries a pStyle so the docxcompose-style numbering
    # restart triggers for it (restart requires a style)
    sub = build_doc(
        numbered_para(7),
        sectpr=sub_sectpr,
        files={
            "word/numbering.xml": SUB_NUMBERING,
            "word/styles.xml": SUB_STYLES,
            "word/header1.xml": _header_with('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'),
        },
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml"), ("rId7", HEADER_RT, "header1.xml")],
        content_types_extra=NUMBERING_CT + HEADER_CT,
    )
    master = build_doc(tp("{{p sub }}"), files={
        "word/numbering.xml": MASTER_NUMBERING,
        "word/styles.xml": XML_DECL + f"<w:styles {NSDECL}>" + MASTER_STYLES + "</w:styles>",
    }, doc_rels=[("rId3", NUMBERING_RT, "numbering.xml"), ("rId4", STYLES_RT, "styles.xml")],
        content_types_extra=NUMBERING_CT + STYLES_CT)
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub), keep_sections=True)})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    header = read_docx_part(data, "word/header1.xml")
    doc = read_docx_part(data, "word/document.xml")
    numbering = read_docx_part(data, "word/numbering.xml")
    # the header's numId was remapped (away from the subdoc's 7)...
    header_id = re.search(r'<w:numId w:val="(\d+)"/>', header).group(1)
    assert header_id != "7"
    assert f'<w:num w:numId="{header_id}">' in numbering
    # ...to the SAME id as the body's (one merged numbering entry, not two):
    # master num 7 + merged sub num + restart num = exactly 3 <w:num> entries
    assert numbering.count("<w:num ") == 3
    # the body's first list was restarted (new num with startOverride),
    # the header's reference is NOT restarted
    body_id = re.search(r'<w:numId w:val="(\d+)"/>', doc).group(1)
    assert body_id != "7" and body_id != header_id
    seg = numbering.split(f'<w:num w:numId="{body_id}">')[1]
    assert "<w:startOverride" in seg.split("</w:num>")[0]


# ---------------- commentsExtended (w15) ----------------


def _comments_ex(*para_ids):
    inner = "".join(
        f'<w15:commentEx w15:paraId="{p}" w15:done="0"/>' for p in para_ids
    )
    return XML_DECL + f'<w15:commentsEx xmlns:w15="{W15}">{inner}</w15:commentsEx>'


def test_comments_extended_copied_when_master_lacks_it():
    master = make_docx(tp("{{p sub }}"))
    sub = build_doc(
        comment_body(0, "sub text"),
        files={
            "word/comments.xml": comments_xml((0, "Bob", "sub note")),
            "word/commentsExtended.xml": _comments_ex("AAAA0001"),
        },
        doc_rels=[("rId9", COMMENTS_RT, "comments.xml")],
        content_types_extra=COMMENTS_CT + COMMENTS_EXTENDED_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    ex = read_docx_part(data, "word/commentsExtended.xml")
    assert "AAAA0001" in ex


def test_comments_extended_appended():
    master = make_docx(
        comment_body(0, "master text") + tp("{{p sub }}"),
        extra_files={
            "word/comments.xml": comments_xml((0, "Ann", "master note")),
            "word/commentsExtended.xml": _comments_ex("BBBB0002"),
        },
    )
    sub = build_doc(
        comment_body(0, "sub text"),
        files={
            "word/comments.xml": comments_xml((0, "Bob", "sub note")),
            "word/commentsExtended.xml": _comments_ex("AAAA0001"),
        },
        doc_rels=[("rId9", COMMENTS_RT, "comments.xml")],
        content_types_extra=COMMENTS_CT + COMMENTS_EXTENDED_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()
    ex = read_docx_part(data, "word/commentsExtended.xml")
    assert "BBBB0002" in ex and "AAAA0001" in ex


# ---------------- footnotes/comments style & numbering consistency ----------------

FOOTNOTES_CT = '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
FOOTNOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

SUB_STYLES_SUBONLY = (
    XML_DECL
    + f'<w:styles {NSDECL}>'
    + '<w:style w:type="paragraph" w:styleId="SubOnly"><w:name w:val="sub only"/>'
    + '<w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>'
    + "</w:styles>"
)
# master's conflicting definition of the same style id (different color)
MASTER_STYLES_SUBONLY = (
    '<w:style w:type="paragraph" w:styleId="SubOnly"><w:name w:val="sub only"/>'
    '<w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>'
)


def footnotes_xml(*notes):
    inner = (
        '<w:footnote w:type="separator" w:id="0"/>'
        '<w:footnote w:type="continuationSeparator" w:id="1"/>' + "".join(notes)
    )
    return XML_DECL + f"<w:footnotes {NSDECL}>{inner}</w:footnotes>"


def test_footnote_style_merged_consistently():
    """A pStyle referenced from both the body and a footnote maps to ONE new
    id (the footnotes content joins the body's styles merge)."""
    styled = '<w:pStyle w:val="SubOnly"/>'
    sub = build_doc(
        f"<w:p><w:pPr>{styled}</w:pPr><w:r><w:t>styled body</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        files={
            "word/styles.xml": SUB_STYLES_SUBONLY,
            "word/footnotes.xml": footnotes_xml(
                f'<w:footnote w:id="2"><w:p><w:pPr>{styled}</w:pPr>'
                "<w:r><w:t>sub footnote text</w:t></w:r></w:p></w:footnote>"
            ),
        },
        doc_rels=[("rId10", STYLES_RT, "styles.xml"), ("rId11", FOOTNOTES_RT, "footnotes.xml")],
        content_types_extra=STYLES_CT + FOOTNOTES_CT,
    )
    master = build_doc(
        tp("{{p sub }}"),
        files={
            "word/styles.xml": XML_DECL + f"<w:styles {NSDECL}>" + MASTER_STYLES_SUBONLY + "</w:styles>",
        },
        doc_rels=[("rId3", STYLES_RT, "styles.xml")],
        content_types_extra=STYLES_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    styles = read_docx_part(data, "word/styles.xml")
    # exactly one renamed copy of the conflicting style
    assert styles.count('w:styleId="SubOnly_1"') == 1
    doc = read_docx_part(data, "word/document.xml")
    assert 'w:pStyle w:val="SubOnly_1"' in doc
    footnotes = read_docx_part(data, "word/footnotes.xml")
    # the footnote references the SAME renamed style as the body
    assert 'w:pStyle w:val="SubOnly_1"' in footnotes
    assert "sub footnote text" in footnotes


def test_comment_numbering_merged_consistently():
    """A numId referenced from both the body and a comment maps to ONE new
    id. The body paragraph carries no pStyle, so the docxcompose numbering
    restart does not trigger and both references stay identical."""
    commented = comment_body(0, "annotated item")
    sub_comments = (
        XML_DECL + f"<w:comments {NSDECL}>"
        '<w:comment w:id="0" w:author="Bob"><w:p><w:pPr>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'
        "</w:pPr><w:r><w:t>sub note</w:t></w:r></w:p></w:comment></w:comments>"
    )
    sub = build_doc(
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>'
        "<w:r><w:t>item</w:t></w:r></w:p>" + commented,
        files={
            "word/numbering.xml": SUB_NUMBERING,
            "word/comments.xml": sub_comments,
        },
        doc_rels=[("rId12", NUMBERING_RT, "numbering.xml"), ("rId9", COMMENTS_RT, "comments.xml")],
        content_types_extra=NUMBERING_CT + COMMENTS_CT,
    )
    master = build_doc(
        tp("{{p sub }}"),
        files={"word/numbering.xml": MASTER_NUMBERING},
        doc_rels=[("rId3", NUMBERING_RT, "numbering.xml")],
        content_types_extra=NUMBERING_CT,
    )
    tpl = DocxTemplate(io.BytesIO(master))
    tpl.render({"sub": tpl.new_subdoc(io.BytesIO(sub))})
    out = io.BytesIO()
    tpl.save(out)
    data = out.getvalue()

    doc = read_docx_part(data, "word/document.xml")
    comments = read_docx_part(data, "word/comments.xml")
    numbering = read_docx_part(data, "word/numbering.xml")
    body_id = re.search(r'<w:numId w:val="(\d+)"/>', doc).group(1)
    comment_id = re.search(r'<w:numId w:val="(\d+)"/>', comments).group(1)
    # both remapped away from the subdoc's 7, to the SAME merged id
    assert body_id != "7"
    assert body_id == comment_id
    assert f'<w:num w:numId="{body_id}">' in numbering
    # master num 7 + one merged sub num (no restart: no pStyle on the body para)
    assert numbering.count("<w:num ") == 2
