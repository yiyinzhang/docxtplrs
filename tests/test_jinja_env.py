"""Tests for passing a real jinja2 Environment to render().

jinja2's own builtin filters/tests/globals must be handled by minijinja's
native implementations (detected by object identity against jinja2.defaults);
only user-added or user-overridden entries are imported as python callables.
"""

import io
import os
import sys

import pytest

jinja2 = pytest.importorskip("jinja2")

sys.path.insert(0, os.path.dirname(__file__))
from helpers import make_docx, read_docx_part, text_of, tp

from docxtplrs import DocxTemplate, TemplateError


def render_text(body, context, env=None, **kw):
    tpl = DocxTemplate(io.BytesIO(make_docx(body, **kw)))
    tpl.render(context, jinja_env=env)
    out = io.BytesIO()
    tpl.save(out)
    return text_of(read_docx_part(out.getvalue(), "word/document.xml"))


def make_env(**kw):
    return jinja2.Environment(**kw)


def test_bare_environment_basic_render():
    assert render_text(tp("{{ name }}"), {"name": "hello"}, make_env()) == "hello"


def test_env_custom_filter_with_args():
    env = make_env()
    env.filters["wrap"] = lambda v, left="[", right="]": f"{left}{v}{right}"
    assert render_text(tp("{{ x | wrap('(', ')') }}"), {"x": "v"}, env) == "(v)"


def test_env_custom_globals():
    env = make_env()
    env.globals["BRAND"] = "ACME"
    env.globals["double"] = lambda x: x * 2
    assert render_text(tp("{{ BRAND }}-{{ double(21) }}"), {}, env) == "ACME-42"


def test_env_custom_test():
    env = make_env()
    env.tests["big"] = lambda x: x > 100
    t = render_text(tp("{% if n is big %}BIG{% else %}small{% endif %}"), {"n": 500}, env)
    assert t == "BIG"


def test_builtin_default_filter_on_missing_variable():
    # jinja2's do_default relies on jinja2.Undefined; it must NOT be imported
    # as a python callback (minijinja's native default handles undefined).
    env = make_env()
    assert render_text(tp("{{ missing | default('DF') }}"), {}, env) == "DF"
    assert render_text(tp("{{ missing | d('DF') }}"), {}, env) == "DF"


def test_builtin_defined_undefined_tests():
    env = make_env()
    t = render_text(
        tp("{{ m is defined }}/{{ m is undefined }}/{{ x is defined }}"),
        {"x": 1},
        env,
    )
    assert t == "false/true/true"


def test_builtin_tojson_filter():
    env = make_env()
    assert render_text(tp("{{ d | tojson }}"), {"d": {"k": 1}}, env) == '{"k":1}'


def test_builtin_urlencode_filter():
    env = make_env()
    assert render_text(tp("{{ q | urlencode }}"), {"q": "a b&c"}, env) == "a%20b%26c"


def test_builtin_filters_still_work():
    env = make_env()
    t = render_text(
        tp("{{ items | join('+') }} {{ name | upper }} {{ items | length }}"),
        {"items": ["a", "b"], "name": "xy"},
        env,
    )
    assert t == "a+b XY 2"


def test_user_override_of_builtin_wins():
    env = make_env()
    env.filters["upper"] = lambda s: "CUSTOM"
    assert render_text(tp("{{ x | upper }}"), {"x": "ab"}, env) == "CUSTOM"


def test_env_entry_overrides_register_filter():
    env = make_env()
    env.filters["f"] = lambda s: "env"
    tpl = DocxTemplate(io.BytesIO(make_docx(tp("{{ x | f }}"))))
    tpl.register_filter("f", lambda s: "reg")
    tpl.render({"x": ""}, jinja_env=env)
    out = io.BytesIO()
    tpl.save(out)
    assert text_of(read_docx_part(out.getvalue(), "word/document.xml")) == "env"


def test_strict_undefined_raises():
    env = make_env(undefined=jinja2.StrictUndefined)
    with pytest.raises(TemplateError):
        render_text(tp("{{ missing }}"), {}, env)


def test_chainable_undefined():
    env = make_env(undefined=jinja2.ChainableUndefined)
    assert render_text(tp("[{{ a.b.c }}]"), {}, env) == "[]"


def test_trim_blocks_passthrough():
    env = make_env(trim_blocks=True, lstrip_blocks=True)
    assert render_text(tp("{% if True %}\nYES\n{% endif %}"), {}, env) == "YES"


def test_duck_typed_fake_env_still_works():
    """Environments not backed by jinja2 keep the legacy import behavior."""

    class FakeEnv:
        autoescape = False
        filters = {"shout": lambda s: str(s).upper()}
        globals = {"G": 7}
        tests = {"neg": lambda x: x < 0}

    t = render_text(
        tp("{{ x | shout }} {{ G }} {{ n is neg }}"),
        {"x": "ab", "n": -1},
        FakeEnv(),
    )
    assert t == "AB 7 true"
