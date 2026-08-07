"""Tests for field-code support: Paragraph.fields / add_field and
settings.update_fields_on_open."""
import io
import sys

import pytest

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp, p, run  # noqa: E402

from docxtplrs import DocxTemplate  # noqa: E402


def new_tpl(body=None, **kw):
    return DocxTemplate(io.BytesIO(make_docx(body or tp("x"), **kw)))


def saved_xml(tpl, part="word/document.xml"):
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), part)


# complex field: PAGE with cached "7"
COMPLEX = (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:t>7</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
)
SIMPLE = '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>'


def test_read_complex_field():
    tpl = new_tpl(p(COMPLEX))
    fields = tpl.paragraphs[0].fields
    assert len(fields) == 1
    f = fields[0]
    assert f.kind == "complex"
    assert f.instr == "PAGE"
    assert f.text == "7"


def test_read_simple_field():
    tpl = new_tpl(p(SIMPLE))
    f = tpl.paragraphs[0].fields[0]
    assert f.kind == "simple"
    assert f.instr == "NUMPAGES"
    assert f.text == "3"


def test_write_complex_field():
    tpl = new_tpl(p(COMPLEX))
    f = tpl.paragraphs[0].fields[0]
    f.instr = "NUMPAGES"
    f.text = "42"
    assert f.instr == "NUMPAGES"
    assert f.text == "42"
    xml = saved_xml(tpl)
    assert "> NUMPAGES </w:instrText>" in xml
    assert "<w:t xml:space=\"preserve\">42</w:t>" in xml
    # structure still intact: begin/separate/end each once
    assert xml.count('w:fldCharType="begin"') == 1
    assert xml.count('w:fldCharType="separate"') == 1
    assert xml.count('w:fldCharType="end"') == 1


def test_write_simple_field():
    tpl = new_tpl(p(SIMPLE))
    f = tpl.paragraphs[0].fields[0]
    f.instr = "PAGE"
    f.text = "9"
    xml = saved_xml(tpl)
    assert 'w:instr=" PAGE "' in xml
    assert "<w:t xml:space=\"preserve\">9</w:t>" in xml


def test_add_field():
    tpl = new_tpl(tp("page: "))
    f = tpl.paragraphs[0].add_field("PAGE", cached="1")
    assert f.instr == "PAGE"
    assert f.text == "1"
    xml = saved_xml(tpl)
    assert 'w:fldCharType="begin"' in xml
    assert "> PAGE </w:instrText>" in xml
    assert 'w:fldCharType="end"' in xml
    # mixed with existing fields: counts accumulate
    tpl2 = new_tpl(p(COMPLEX))
    tpl2.paragraphs[0].add_field("NUMPAGES")
    fields = tpl2.paragraphs[0].fields
    assert len(fields) == 2
    assert [f.instr for f in fields] == ["PAGE", "NUMPAGES"]


def test_update_fields_on_open():
    tpl = new_tpl()
    st = tpl.get_docx().settings
    assert st.update_fields_on_open is False
    st.update_fields_on_open = True
    assert st.update_fields_on_open is True
    xml = saved_xml(tpl, "word/settings.xml")
    assert "<w:updateFields/>" in xml
    st.update_fields_on_open = False
    assert st.update_fields_on_open is False
