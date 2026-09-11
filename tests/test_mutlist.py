"""Coverage for template-local mutable lists: jinja2-style
`{% set x = [] %}` + `x.append(...)` in-place mutation (MutList).
"""
import io
import re
import sys

sys.path.insert(0, "tests")
from helpers import make_docx, read_docx_part, tp  # noqa: E402

from docxtplrs import DocxTemplate  # noqa: E402


def render_text(body, ctx=None, **render_kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(ctx or {}, **render_kw)
    out = io.BytesIO()
    tpl.save(out)
    xml = read_docx_part(out.getvalue(), "word/document.xml")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def test_append_in_loop():
    body = tp(
        "{% set xs = [] %}"
        "{% for i in items %}{% set _ = xs.append(i) %}{% endfor %}"
        "{{ xs | join(',') }}"
    )
    assert render_text(body, {"items": [1, 2, 3]}) == "1,2,3"


def test_append_with_condition():
    body = tp(
        "{% set xs = [] %}"
        "{% for k, v in d.items() %}"
        "{% if v %}{% set _ = xs.append(k ~ '=' ~ v) %}{% endif %}"
        "{% endfor %}"
        "{{ xs | join(';') if xs | length > 0 else 'empty' }}"
    )
    assert render_text(body, {"d": {"a": "1", "b": "", "c": "3"}}) == "a=1;c=3"


def test_append_empty_list_fallback():
    body = tp(
        "{% set xs = [] %}"
        "{% for i in items %}{% set _ = xs.append(i) %}{% endfor %}"
        "{{ xs | join(',') if xs | length > 0 else 'empty' }}"
    )
    assert render_text(body, {"items": []}) == "empty"


def test_appended_list_truthiness_and_index():
    body = tp(
        "{% set xs = [] %}"
        "{% set _ = xs.append('a') %}"
        "{% if xs %}first={{ xs[0] }} len={{ xs | length }}{% endif %}"
    )
    assert render_text(body) == "first=a len=1"


def test_append_whitespace_control_set_tag():
    body = tp(
        "{%- set xs = [] -%}"
        "{% set _ = xs.append('x') %}"
        "{{ xs | length }}"
    )
    assert render_text(body) == "1"


def test_extend_insert_pop_remove_clear():
    body = tp(
        "{% set xs = [] %}"
        "{% set _ = xs.extend([1, 2]) %}"
        "{% set _ = xs.insert(1, 9) %}"
        "{{ xs | join(',') }}"
        "{% set _ = xs.remove(9) %}"
        "|{{ xs.pop() }}|{{ xs | join(',') }}"
        "{% set _ = xs.clear() %}"
        "|{{ xs | length }}"
    )
    assert render_text(body) == "1,9,2|2|1|0"


def test_list_repr():
    body = tp("{% set xs = [] %}{% set _ = xs.append('a') %}{% set _ = xs.append('b') %}{{ xs }}")
    assert render_text(body) == "['a', 'b']"


def test_plain_list_literal_without_append_untouched():
    # `{% set x = [] %}` without any append() keeps the native seq path
    body = tp("{% set xs = [] %}{{ xs | length }}{% set ys = [1, 2] %}{{ ys | sum }}")
    assert render_text(body) == "03"


def test_append_result_is_none():
    # list.append returns None in python/jinja2
    body = tp("{% set xs = [] %}{% set r = xs.append('a') %}{{ r }}{{ xs | length }}")
    assert render_text(body) == "None1"
