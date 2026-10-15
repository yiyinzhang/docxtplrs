"""Tests for the CLI: python -m docxtplrs template.docx data.json out.docx."""

import json
import os
import subprocess
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, tp


def run_cli(*args, stdin=""):
    return subprocess.run(
        [sys.executable, "-m", "docxtplrs", *args],
        input=stdin,
        capture_output=True,
        text=True,
        timeout=60,
    )


@pytest.fixture
def template_and_json(tmp_path):
    tpl_path = tmp_path / "template.docx"
    tpl_path.write_bytes(make_docx(tp("Hello {{ name }}")))
    json_path = tmp_path / "data.json"
    json_path.write_text(json.dumps({"name": "World"}), encoding="utf-8")
    return tpl_path, json_path


def test_cli_renders_template(tmp_path, template_and_json):
    tpl_path, json_path = template_and_json
    out = tmp_path / "out.docx"
    r = run_cli(str(tpl_path), str(json_path), str(out))
    assert r.returncode == 0, r.stderr
    assert "Saved:" in r.stdout
    text = text_of(read_docx_part(out.read_bytes(), "word/document.xml"))
    assert "Hello World" in text


def test_cli_existing_output_requires_flag_or_confirmation(tmp_path, template_and_json):
    tpl_path, json_path = template_and_json
    out = tmp_path / "out.docx"
    out.write_bytes(b"sentinel")
    # without -o the CLI asks; answering "n" aborts with a non-zero exit code
    r = run_cli(str(tpl_path), str(json_path), str(out), stdin="n\n")
    assert r.returncode != 0
    assert out.read_bytes() == b"sentinel"
    # -o overwrites without asking
    r = run_cli("-o", str(tpl_path), str(json_path), str(out))
    assert r.returncode == 0, r.stderr
    assert "Hello World" in text_of(read_docx_part(out.read_bytes(), "word/document.xml"))


def test_cli_missing_arguments_fails(tmp_path):
    r = run_cli()
    assert r.returncode != 0
    # argparse reports usage on stderr
    assert "usage" in r.stderr.lower()


def test_cli_template_error_fails(tmp_path):
    tpl_path = tmp_path / "bad.docx"
    tpl_path.write_bytes(make_docx(tp("{% endfor %}")))
    json_path = tmp_path / "data.json"
    json_path.write_text("{}", encoding="utf-8")
    out = tmp_path / "out.docx"
    r = run_cli(str(tpl_path), str(json_path), str(out))
    assert r.returncode != 0
    assert "Template error" in r.stderr
    assert not out.exists()


def test_cli_rejects_non_docx_template(tmp_path, template_and_json):
    _, json_path = template_and_json
    not_docx = tmp_path / "template.txt"
    not_docx.write_text("nope", encoding="utf-8")
    r = run_cli(str(not_docx), str(json_path), str(tmp_path / "out.docx"))
    assert r.returncode != 0
