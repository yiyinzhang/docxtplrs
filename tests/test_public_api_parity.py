"""Tests for the docxtpl semi-public pipeline methods and public attributes
exposed on DocxTemplate (string-based variants of docxtpl's lxml-based API):

patch_xml / resolve_listing / fix_tables / fix_docpr_ids / xml_to_string /
render_xml_part / render_properties / render_footnotes / build_xml /
map_tree / get_part_xml / get_headers_footers_encoding /
build_headers_footers_xml / map_headers_footers_xml / init_docx /
render_init / pre_processing / post_processing, plus the template_file /
docx / pic_map / current_rendering_part properties, the four replacement
dicts and the is_saved / is_rendered setters.
"""

import io
import os
import re
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import (
    CORE_XML,
    make_docx,
    read_docx_part,
    tp,
    p,
    run,
    cell,
    tr,
    tbl,
    document_xml,
)

from docxtplrs import DocxTemplate


def tpl_of(body, **kw):
    return DocxTemplate(io.BytesIO(make_docx(body, **kw)))


# ---------------- patch_xml / resolve_listing ----------------

def test_patch_xml_merges_split_tags():
    tpl = tpl_of(tp("x"))
    frag = p(run("{{"), run("name"), run("}}"))
    out = tpl.patch_xml(frag)
    assert isinstance(out, str)
    assert "{{name}}" in out
    # the three runs are merged into a single text node
    assert out.count("<w:t") == 1


def test_patch_xml_decodes_text_entities():
    tpl = tpl_of(tp("x"))
    out = tpl.patch_xml(p(run("a&quot;b")))
    assert 'a"b' in out


def test_resolve_listing_newline():
    tpl = tpl_of(tp("x"))
    out = tpl.resolve_listing(p(run("a\nb")))
    assert "<w:br/>" in out


def test_resolve_listing_tab():
    tpl = tpl_of(tp("x"))
    out = tpl.resolve_listing(p(run("a\tb")))
    assert "<w:tab/>" in out


# ---------------- fix_tables / fix_docpr_ids / map_tree ----------------

BROKEN_TBL = document_xml(
    tbl([tr(cell(tp("a")), cell(tp("b")), cell(tp("c")))], widths=(2000, 2000))
)


def test_fix_tables_adds_gridcol():
    tpl = tpl_of(tp("x"))
    out = tpl.fix_tables(BROKEN_TBL)
    assert isinstance(out, str)
    assert out.count("<w:gridCol") == 3


def test_fix_tables_noop_on_valid():
    tpl = tpl_of(tp("x"))
    xml = document_xml(tbl([tr(cell(tp("a")), cell(tp("b")))]))
    assert tpl.fix_tables(xml) == xml


def test_fix_docpr_ids_renumbers():
    tpl = tpl_of(tp("x"))
    tpl.render_init()  # docx_ids_index = 1000, like before a render
    xml = document_xml(
        '<w:p><w:r><w:drawing>'
        '<wp:docPr id="1"/><wp:docPr id="1"/>'
        '<pic:cNvPr id="7"/><pic:cNvPr id="7"/>'
        "</w:drawing></w:r></w:p>"
    )
    out = tpl.fix_docpr_ids(xml)
    docpr = re.findall(r'wp:docPr id="(\d+)"', out)
    cnvpr = re.findall(r'pic:cNvPr id="(\d+)"', out)
    assert len(docpr) == 2 and len(set(docpr)) == 2
    assert len(cnvpr) == 2 and len(set(cnvpr)) == 2
    assert all(int(i) > 1000 for i in docpr)


def test_map_tree_fixes_tables_and_docpr():
    tpl = tpl_of(tp("x"))
    tpl.render_init()
    xml = document_xml(
        tbl([tr(cell(tp("a")), cell(tp("b")), cell(tp("c")))], widths=(2000, 2000))
        + '<w:p><w:r><w:drawing><wp:docPr id="1"/></w:drawing></w:r></w:p>'
    )
    out = tpl.map_tree(xml)
    assert out.count("<w:gridCol") == 3
    assert 'wp:docPr id="1"' not in out


# ---------------- xml_to_string / get_headers_footers_encoding ----------------

def test_xml_to_string_str_passthrough():
    tpl = tpl_of(tp("x"))
    assert tpl.xml_to_string("<a>é</a>") == "<a>é</a>"


def test_xml_to_string_bytes_decoding():
    tpl = tpl_of(tp("x"))
    assert tpl.xml_to_string(b"<a/>") == "<a/>"
    assert tpl.xml_to_string("<a>é</a>".encode("latin-1"), encoding="latin-1") == "<a>é</a>"
    with pytest.raises(Exception):
        tpl.xml_to_string(42)


def test_get_headers_footers_encoding():
    tpl = tpl_of(tp("x"))
    decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr/>'
    assert tpl.get_headers_footers_encoding(decl) == "UTF-8"
    assert tpl.get_headers_footers_encoding(decl.encode("utf-8")) == "UTF-8"
    assert tpl.get_headers_footers_encoding("<w:hdr/>") == "utf-8"
    decl1252 = '<?xml version="1.0" encoding="windows-1252"?><w:hdr/>'
    assert tpl.get_headers_footers_encoding(decl1252) == "windows-1252"


# ---------------- render_xml_part / build_xml ----------------

def test_render_xml_part():
    tpl = tpl_of(tp("x"))
    out = tpl.render_xml_part(p(run("Hello {{ name }}")), "word/document.xml", {"name": "World"})
    assert isinstance(out, str)
    assert "Hello World" in out
    assert "{{" not in out


def test_render_xml_part_sets_current_rendering_part():
    tpl = tpl_of(tp("x"))
    seen = []
    tpl.register_function("probe", lambda: seen.append(tpl.current_rendering_part) or "")
    tpl.render_xml_part(p(run("{{ probe() }}")), "word/header1.xml", {})
    assert seen == ["word/header1.xml"]
    assert tpl.current_rendering_part is None


def test_current_rendering_part_during_render():
    tpl = tpl_of(tp("{{ probe() }}"), headers={"header1.xml": tp("{{ probe() }}")})
    seen = []
    tpl.register_function("probe", lambda: seen.append(tpl.current_rendering_part) or "")
    tpl.render({})
    assert seen == ["word/document.xml", "word/header1.xml"]
    assert tpl.current_rendering_part is None


def test_build_xml_does_not_touch_package():
    tpl = tpl_of(tp("Hello {{ name }}"))
    out = tpl.build_xml({"name": "World"})
    assert "Hello World" in out
    # the stored template is unchanged
    assert "{{ name }}" in tpl.get_xml()


def test_build_xml_map_tree_equivalent_to_render():
    blob = make_docx(tp("Hello {{ name }}"))
    manual = DocxTemplate(io.BytesIO(blob))
    manual.render_init()
    xml = manual.map_tree(manual.build_xml({"name": "World"}))
    auto = DocxTemplate(io.BytesIO(blob))
    auto.render({"name": "World"})
    assert xml == auto.get_xml()


# ---------------- render_properties / render_footnotes / get_part_xml ----------------

def test_render_properties():
    core = CORE_XML.format(title="{{ t }}", author="me")
    tpl = tpl_of(tp("x"), core=core)
    tpl.render_properties({"t": "Hello"})
    part = tpl.get_part_xml("docProps/core.xml")
    assert "<dc:title>Hello</dc:title>" in part
    assert "<dc:creator>me</dc:creator>" in part


def test_render_footnotes():
    fn = '<w:footnote w:id="1">' + p(run("note {{ x }}")) + "</w:footnote>"
    tpl = tpl_of(tp("x"), footnotes=fn)
    tpl.render_footnotes({"x": "42"})
    assert "note 42" in tpl.get_part_xml("word/footnotes.xml")


def test_render_footnotes_no_part():
    tpl = tpl_of(tp("x"))
    tpl.render_footnotes({"x": "42"})  # no footnotes part: no-op


def test_get_part_xml():
    tpl = tpl_of(tp("Hello {{ name }}"))
    assert "{{ name }}" in tpl.get_part_xml("word/document.xml")
    with pytest.raises(Exception):
        tpl.get_part_xml("word/nonexistent.xml")


# ---------------- headers/footers build & map ----------------

def test_build_headers_footers_xml():
    tpl = tpl_of(
        tp("x"),
        headers={"header1.xml": tp("H {{ name }}")},
        footers={"footer1.xml": tp("F {{ name }}")},
    )
    headers = tpl.build_headers_footers_xml({"name": "N"}, DocxTemplate.HEADER_URI)
    assert isinstance(headers, dict) and len(headers) == 1
    assert "H N" in next(iter(headers.values()))
    footers = tpl.build_headers_footers_xml({"name": "N"}, DocxTemplate.FOOTER_URI)
    assert "F N" in next(iter(footers.values()))


def test_map_headers_footers_xml():
    tpl = tpl_of(
        tp("x"),
        headers={"header1.xml": tp("h")},
        footers={"footer1.xml": tp("f")},
    )
    tpl.render_init()  # docx_ids_index = 1000
    hdr_rid = tpl.get_headers_footers(DocxTemplate.HEADER_URI)[0][0]
    ftr_rid = tpl.get_headers_footers(DocxTemplate.FOOTER_URI)[0][0]

    # tables are fixed for both headers and footers
    broken = tbl([tr(cell(tp("a")), cell(tp("b")), cell(tp("c")))], widths=(2000, 2000))
    assert tpl.map_headers_footers_xml(hdr_rid, broken).count("<w:gridCol") == 3
    assert tpl.map_headers_footers_xml(ftr_rid, broken).count("<w:gridCol") == 3

    # docPr ids are renumbered for headers only
    frag = '<w:hdr><w:p><w:r><w:drawing><wp:docPr id="1"/></w:drawing></w:r></w:p></w:hdr>'
    out_hdr = tpl.map_headers_footers_xml(hdr_rid, frag)
    assert 'wp:docPr id="1"' not in out_hdr
    out_ftr = tpl.map_headers_footers_xml(ftr_rid, frag)
    assert 'wp:docPr id="1"' in out_ftr


# ---------------- init_docx / render_init / pre_processing / post_processing ----------------

def test_init_docx_reload_semantics():
    tpl = tpl_of(tp("{{ x }}"))
    tpl.render({"x": "1"})
    assert tpl.is_rendered
    tpl.init_docx(reload=False)  # already loaded + rendered: no-op
    assert tpl.is_rendered
    tpl.init_docx()  # reload=True: package reloaded, rendered flag reset
    assert not tpl.is_rendered
    assert "{{ x }}" in tpl.get_xml()


def test_render_init_resets_state():
    tpl = tpl_of(tp("{{ x }}"))
    tpl.render({"x": "1"})
    tpl.save(io.BytesIO())
    assert tpl.is_saved and tpl.is_rendered
    tpl.render_init()
    assert not tpl.is_saved and not tpl.is_rendered
    assert tpl.pic_map == {}


def test_pre_processing_missing_pic():
    tpl = tpl_of(tp("x"))
    tpl.pre_processing()  # empty pics_to_replace: no-op
    tpl.pics_to_replace = {"missing.png": b"data"}
    with pytest.raises(Exception):
        tpl.pre_processing()
    tpl.allow_missing_pics = True
    tpl.pre_processing()  # tolerated now


def test_post_processing_zipname(tmp_path):
    blob = make_docx(tp("x"), extra_files={"word/embeddings/obj.bin": b"OLD"})
    tpl = DocxTemplate(io.BytesIO(blob))
    path = tmp_path / "out.docx"
    tpl.save(str(path))
    with zipfile.ZipFile(path) as z:
        assert z.read("word/embeddings/obj.bin") == b"OLD"
    tpl.zipname_to_replace = {"word/embeddings/obj.bin": b"NEW"}
    tpl.post_processing(str(path))
    with zipfile.ZipFile(path) as z:
        assert z.read("word/embeddings/obj.bin") == b"NEW"
        # unrelated entries untouched
        assert b"x" in z.read("word/document.xml")


def test_post_processing_media_crc(tmp_path):
    blob = make_docx(tp("x"), extra_files={"word/media/img.png": b"OLDIMG"})
    tpl = DocxTemplate(io.BytesIO(blob))
    path = tmp_path / "out.docx"
    tpl.save(str(path))
    crc = DocxTemplate.get_file_crc(b"OLDIMG")
    tpl.crc_to_new_media = {crc: b"NEWIMG"}
    tpl.post_processing(str(path))
    with zipfile.ZipFile(path) as z:
        assert z.read("word/media/img.png") == b"NEWIMG"


def test_post_processing_noop_when_empty(tmp_path):
    tpl = tpl_of(tp("x"))
    # all replacement dicts empty: does not even require the file to exist
    tpl.post_processing(str(tmp_path / "does_not_exist.docx"))


# ---------------- replacement dicts getter/setter ----------------

def test_replacement_dicts_roundtrip():
    tpl = tpl_of(tp("x"))
    assert tpl.crc_to_new_media == {}
    assert tpl.crc_to_new_embedded == {}
    assert tpl.zipname_to_replace == {}
    assert tpl.pics_to_replace == {}

    tpl.crc_to_new_media = {123: b"abc"}
    assert tpl.crc_to_new_media == {123: b"abc"}
    tpl.crc_to_new_embedded = {456: b"def"}
    assert tpl.crc_to_new_embedded == {456: b"def"}
    tpl.zipname_to_replace = {"word/a.bin": b"g"}
    assert tpl.zipname_to_replace == {"word/a.bin": b"g"}
    tpl.pics_to_replace = {"p.png": b"h"}
    assert tpl.pics_to_replace == {"p.png": b"h"}

    # getters return snapshots: in-place mutation does not stick
    snap = tpl.crc_to_new_media
    snap[999] = b"x"
    assert tpl.crc_to_new_media == {123: b"abc"}

    # setter replaces the whole map
    tpl.crc_to_new_media = {}
    assert tpl.crc_to_new_media == {}


def test_replacement_dicts_updated_by_replace_methods():
    tpl = tpl_of(tp("x"))
    tpl.replace_zipname("word/embeddings/x.bin", io.BytesIO(b"zz"))
    assert tpl.zipname_to_replace == {"word/embeddings/x.bin": b"zz"}
    tpl.reset_replacements()
    assert tpl.zipname_to_replace == {}


# ---------------- template_file / docx / pic_map / flags ----------------

def test_template_file_path(tmp_path):
    path = tmp_path / "t.docx"
    path.write_bytes(make_docx(tp("x")))
    assert DocxTemplate(str(path)).template_file == str(path)
    assert DocxTemplate(path).template_file == str(path)  # PathLike


def test_template_file_none_for_bytes_and_filelike():
    assert DocxTemplate(make_docx(tp("x"))).template_file is None
    assert tpl_of(tp("x")).template_file is None


def test_docx_property():
    tpl = tpl_of(tp("hello"))
    doc = tpl.docx  # explicit property, not the __getattr__ delegation
    assert [para.text for para in doc.paragraphs] == ["hello"]


def test_pic_map_property_matches_get_pic_map():
    tpl = tpl_of(tp("x"))
    assert tpl.pic_map == {}
    assert tpl.pic_map == tpl.get_pic_map()


def test_is_saved_is_rendered_setters():
    tpl = tpl_of(tp("x"))
    assert not tpl.is_rendered and not tpl.is_saved
    tpl.is_rendered = True
    tpl.is_saved = True
    assert tpl.is_rendered and tpl.is_saved
    tpl.is_rendered = False
    tpl.is_saved = False
    assert not tpl.is_rendered and not tpl.is_saved
