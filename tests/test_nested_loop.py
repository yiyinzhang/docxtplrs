"""Test nested for loops in docxtplrs."""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, tp, cell, tr, tbl

from docxtplrs import DocxTemplate


def render_xml(body, context):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(context)
    out = io.BytesIO()
    tpl.save(out)
    return read_docx_part(out.getvalue(), "word/document.xml")


def test_nested_for_loops():
    """Test nested for loops - inner loop should access outer loop variable."""
    # Simple nested loop: outer loop over categories, inner loop over items
    body = tp("{% for cat in categories %}{{ cat.name }}: {% for item in cat.items %}{{ item }}, {% endfor %}{% endfor %}")
    xml = render_xml(body, {
        "categories": [
            {"name": "A", "items": ["a1", "a2"]},
            {"name": "B", "items": ["b1", "b2"]},
        ]
    })
    text = text_of(xml)
    print(f"Simple nested loop output: {text}")
    assert "A:" in text and "a1" in text and "a2" in text
    assert "B:" in text and "b1" in text and "b2" in text


def test_nested_for_in_table():
    """Test nested for loops inside table rows (the reported issue)."""
    # This mimics the structure: outer loop over mutations, inner loop over
    # each mutation's entries (jinja2 iterates dict keys, so the inner loop
    # must run over a list, like test_nested_for_simple below)
    body = tbl([
        tr(cell(tp("{%tr for mutation in mutations %}")), cell(tp(""))),
        tr(cell(tp("{{ mutation.name }}")), cell(tp(""))),
        tr(cell(tp("{%tr for mut in mutation.muts %}")), cell(tp(""))),
        tr(cell(tp("{{ mut.pdot }}")), cell(tp(""))),
        tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
        tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
    ])
    xml = render_xml(body, {
        "mutations": [
            {"name": "mut1", "muts": [{"pdot": "p1"}]},
            {"name": "mut2", "muts": [{"pdot": "p2"}]},
        ]
    })
    text = text_of(xml)
    print(f"Table nested loop output: {text}")
    # Inner loop should render mut.pdot values
    assert "p1" in text, f"Expected 'p1' in output, got: {text}"
    assert "p2" in text, f"Expected 'p2' in output, got: {text}"


def test_nested_for_simple():
    """Test the exact structure from the reported issue."""
    # Structure from treport.docx:
    # {%tr for mutation in [wildtype, target, fusion, cnv, longindel, hrd ] %}
    #   {%tr for mut in mutation %}
    #     {{ mut.pdot }}
    #   {%tr endfor %}
    # {%tr endfor %}

    body = tbl([
        tr(cell(tp("{%tr for mutation in mutation_list %}")), cell(tp(""))),
        tr(cell(tp("{%tr for mut in mutation %}")), cell(tp(""))),
        tr(cell(tp("{{ mut.pdot }}")), cell(tp(""))),
        tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
        tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
    ])

    # mutation_list contains lists of mutation dicts
    context = {
        "mutation_list": [
            [{"pdot": "p.Ala1"}, {"pdot": "p.Ala2"}],  # first mutation group
            [{"pdot": "p.Bla1"}],                       # second mutation group
        ]
    }

    xml = render_xml(body, context)
    text = text_of(xml)
    print(f"Exact structure output: {text}")

    # Should render all pdot values
    assert "p.Ala1" in text, f"Expected 'p.Ala1' in output"
    assert "p.Ala2" in text, f"Expected 'p.Ala2' in output"
    assert "p.Bla1" in text, f"Expected 'p.Bla1' in output"


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-v", "-s"])
