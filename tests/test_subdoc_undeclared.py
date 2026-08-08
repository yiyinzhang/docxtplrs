"""Tests for Subdoc attribute delegation (__getattr__/docx facade over the
subdoc content) and get_undeclared_template_variables(jinja_env=...)."""

import io
import os
import sys

import pytest

jinja2 = pytest.importorskip("jinja2")

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, tp, tbl, tr, cell

from docxtplrs import DocxTemplate


# ---------------- Subdoc.__getattr__ delegation ----------------

def test_subdoc_delegation_file_based():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("master {{p sub }}"))))
    sub = tpl.new_subdoc(io.BytesIO(make_docx(
        tp("sub para one") + tp("sub para two") + tbl([tr(cell(tp("c1")))])
    )))
    assert [p.text for p in sub.paragraphs] == ["sub para one", "sub para two"]
    assert len(sub.tables) == 1
    assert len(sub.sections) == 1
    assert type(sub.styles).__name__ == "Styles"


def test_subdoc_docx_property_file_based():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sub = tpl.new_subdoc(io.BytesIO(make_docx(tp("hello"))))
    doc = sub.docx
    assert [p.text for p in doc.paragraphs] == ["hello"]


def test_subdoc_delegation_builder():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sub = tpl.new_subdoc()
    sub.add_paragraph("built para")
    sub.add_table(2, 3)
    assert [p.text for p in sub.paragraphs] == ["built para"]
    assert len(sub.tables) == 1
    assert len(sub.tables[0].rows) == 2
    assert len(sub.tables[0].columns) == 3
    # docx property reflects the accumulated content too
    assert [p.text for p in sub.docx.paragraphs] == ["built para"]


def test_subdoc_delegation_unknown_attr_raises():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("x"))))
    sub = tpl.new_subdoc(io.BytesIO(make_docx(tp("y"))))
    with pytest.raises(AttributeError):
        sub.no_such_attribute


def test_subdoc_delegation_does_not_break_render():
    # accessing facade attributes must not disturb lazy materialization
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("master:") + tp("{{p sub }}"))))
    sub = tpl.new_subdoc(io.BytesIO(make_docx(tp("sub content"))))
    assert sub.paragraphs[0].text == "sub content"
    tpl.render({"sub": sub})
    assert "sub content" in tpl.get_xml()

    tpl2 = DocxTemplate(io.BytesIO(make_docx(tp("{{p sub }}"))))
    b = tpl2.new_subdoc()
    b.add_paragraph("built")
    assert b.paragraphs[0].text == "built"
    tpl2.render({"sub": b})
    assert "built" in tpl2.get_xml()


# ---------------- get_undeclared_template_variables(jinja_env) ----------------

def test_undeclared_no_env_unchanged():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ a }} {{ b }}"))))
    assert tpl.get_undeclared_template_variables() == {"a", "b"}
    assert tpl.get_undeclared_template_variables(context={"a": 1}) == {"b"}


def test_undeclared_with_env_basic():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ a }} {{ b }}"))))
    env = jinja2.Environment(trim_blocks=True, lstrip_blocks=True,
                             keep_trailing_newline=True)
    assert tpl.get_undeclared_template_variables(jinja_env=env) == {"a", "b"}
    # context filtering still applies
    assert tpl.get_undeclared_template_variables(jinja_env=env, context={"a": 1}) == {"b"}


def test_undeclared_trans_requires_env():
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{% trans %}{{ name }}{% endtrans %}"))))
    # without an env the engine parses with default syntax (no i18n), like
    # docxtpl's default jinja2 Environment without the i18n extension
    with pytest.raises(Exception):
        tpl.get_undeclared_template_variables()
    env = jinja2.Environment()
    assert tpl.get_undeclared_template_variables(jinja_env=env) == {"name"}


def test_undeclared_trans_assignments_and_plural():
    env = jinja2.Environment()
    tpl = DocxTemplate(io.BytesIO(make_docx(
        tp("{% trans user=name %}Hi {{ user }}{% endtrans %}"))))
    # assigned inside trans: user is not undeclared (jinja2 semantics)
    assert tpl.get_undeclared_template_variables(jinja_env=env) == {"name"}

    tpl2 = DocxTemplate(io.BytesIO(make_docx(
        tp("{% trans cnt=items|length %}one{% pluralize cnt %}{{ cnt }} items{% endtrans %}"))))
    assert tpl2.get_undeclared_template_variables(jinja_env=env) == {"items"}
