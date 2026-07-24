# docxtplrs

> ⚠️ **Vibecoding project**: this codebase was written by an AI (Kimi Code) in a
> pair-programming session with the user, without line-by-line human review.
> It is verified against the official docxtpl test suite (32 real-world template
> cases) plus 173 in-house tests, but undiscovered issues may remain — evaluate
> before production use.

A **Rust implementation of [python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template)**,
usable both as a **Rust crate** and as a **Python package** (via PyO3) with an API
almost identical to docxtpl. The template engine is
[minijinja](https://docs.rs/minijinja) (Jinja2-compatible syntax); all docx zip/XML
handling, template preprocessing, table fixing, relationship management and the
document object model are native Rust — **no dependency** on
python-docx / lxml / jinja2.

---

## 1. Usage as a Rust crate

Add to your `Cargo.toml` (path dependency for now):

```toml
[dependencies]
docxtplrs = { path = "../docxtplrs" }
minijinja = "2"
```

Render a template:

```rust
use docxtplrs::template::TplCore;
use minijinja::Value;
use std::collections::BTreeMap;

// build the rendering context
let mut ctx = BTreeMap::new();
ctx.insert("title".to_string(), Value::from("Quarterly Report"));
ctx.insert("items".to_string(), Value::from_serialize(&vec!["alpha", "beta"]));

// load, render (autoescape = false), save
let mut tpl = TplCore::new(std::fs::read("template.docx")?);
tpl.render(false, &|_core, _part| Ok(Value::from_serialize(&ctx)))?;
std::fs::write("out.docx", tpl.save_bytes()?)?;
```

Useful Rust API surface (all in `docxtplrs::`):

| Module | Highlights |
|---|---|
| `template` | `TplCore`: `new/render/save_bytes/get_xml/undeclared_variables/build_url_id`, `Deferred`, replacement maps (`crc_to_new_media`, `pics_to_replace`, …) |
| `package` | `Package` (docx zip entries, `.rels`, content types), `Rels`, CRC32/SHA1 helpers |
| `patch` | `patch_xml`, `resolve_listing`, `decode_text_entities` (docxtpl's XML preprocessing) |
| `richtext` | `richtext_run`, `richtext_paragraph`, `listing_xml`, `TextProps` |
| `image` | `ImageInfo` (PNG/JPEG/GIF/BMP/TIFF size & DPI), EMU conversions |
| `gettext` | `Catalog` (.mo parsing, plural-rule evaluation) |
| `xmldom` | Minimal XML DOM used for structured edits |

A runnable example is included:

```bash
cargo run --example render --release -- template.docx out.docx
```

> Note: examples link against `libpython` (PyO3). If your Python is uv-managed, run with
> `LD_LIBRARY_PATH=$(python -c "import sysconfig; print(sysconfig.get_config_var('LIBDIR'))")`.
> The `extension-module` cargo feature is only enabled by maturin for the Python wheel build.

---

## 2. Usage as a Python package

### Install

```bash
# in this repo: build & install into the uv venv
uv sync

# in another uv project
uv add /path/to/docxtplrs/target/wheels/docxtplrs-0.1.0-cp313-cp313-manylinux_2_34_x86_64.whl
# or editable (local development)
uv add --editable /path/to/docxtplrs
```

The wheel is a native `cp313` build. For other Python versions either build under that
version, or set pyo3 to `abi3-py312` in `Cargo.toml` for a single portable wheel.

### Quick start

```python
from docxtplrs import DocxTemplate, R, RichText, Listing, InlineImage, Mm

tpl = DocxTemplate("template.docx")          # path / file-like / bytes

tpl.render({
    "title": "Quarterly Report",
    "rows": ["East", "South"],                # {%p for r in rows %} paragraph loop
    "items": [{"name": "A", "price": 12.5}],  # {%tr for x in items %} table row loop
    "rt": R("important", bold=True, color="#C00000"),  # write {{r rt }} in the template
    "img": InlineImage(tpl, "logo.png", width=Mm(15)), # {{ img }}
    "sub": tpl.new_subdoc("chapter.docx"),             # {{p sub }}
}, autoescape=True)

tpl.save("out.docx")                         # or io.BytesIO()
```

### Template syntax (same as docxtpl)

| Syntax | Purpose |
|---|---|
| `{{ var }}` | variables (filters, attr/item access, method calls, str methods, `'%s' % x`, Unicode identifiers) |
| `{% if %}` `{% for %}` `{% set %}` `{% macro %}` | statements (break/continue/do supported) |
| `{%tr for x in items %}` | table-row loop |
| `{%tc for x in items %}` | cell loop (auto-fixes `tblGrid` columns/widths) |
| `{%p if x %}` / `{{p var }}` | paragraph-level statement / replacement |
| `{{r var }}` | run-level replacement (with `RichText`) |
| `{% colspan x %}` `{% cellbg c %}` `{% vm %}` `{% hm %}` | colspan / cell background / vertical & horizontal merge |
| `{%- ... -%}` / `{_{ ... }_}` | paragraph text merging / literal-brace escaping |
| `{% trans %}` / `{% pluralize %}` | i18n (after `tpl.install_gettext("zh.mo")`) |
| `{% include %}` / `{% import %}` | includes (after `tpl.set_template_loader(fn)`) |
| headers / footers / footnotes / doc properties | rendered as well (images can be inserted into headers) |

### Python API summary

- `DocxTemplate(f)`: `render()`, `save()`, `get_docx_bytes()`, `get_xml()`,
  `get_undeclared_template_variables()`, `new_subdoc()`, `build_url_id()`,
  `replace_media()` / `replace_embedded()` / `replace_pic()` / `replace_zipname()`,
  `reset_replacements()`, `allow_missing_pics`, `get_pic_map()`
- Customization: `register_filter/test/function/global()`, `set_template_loader()`,
  `install_gettext()`; `render(ctx, jinja_env=...)` accepts a real jinja2
  `Environment` (reads `autoescape/filters/globals/tests/trim_blocks/lstrip_blocks/
  keep_trailing_newline/undefined`)
- Document object model (read-write): `get_docx()`, `paragraphs/tables/sections/
  styles/settings/comments/core_properties/inline_shapes`,
  `add_paragraph/add_heading/add_picture/add_table/add_page_break/add_section`,
  run/cell formatting assignment, `__getattr__` delegation
- `RichText`/`R`, `RichTextParagraph`/`RP`, `Listing`, `InlineImage`, `Subdoc`
  (file merging + programmatic `add_paragraph/add_picture/add_table`)
- Units: `Length/Emu/Inches/Cm/Mm/Pt/Twips`; jinja2.utils: `Cycler/Joiner/
  generate_lorem_ipsum`; exception: `TemplateError` (with `docx_context`)
- CLI: `python -m docxtplrs template.docx data.json output.docx [-o] [-q]`

### Tests

```bash
.venv/bin/python -m pytest tests/ -q          # 173 unit tests
.venv/bin/python tests/crosscheck.py docxtplrs > /tmp/rs.json
.venv/bin/python tests/crosscheck.py docxtpl > /tmp/ref.json   # needs crosscheck group
.venv/bin/python tests/crosscheck.py compare /tmp/ref.json /tmp/rs.json
```

---

## 3. Limitations

**Architectural**

- The engine is minijinja, not jinja2: broadly compatible, but exotic jinja2 edge
  behavior is not guaranteed (common gaps already covered with custom tests/filters,
  loader, Unicode identifiers, kwargs support, etc.).
- The document object model covers python-docx's common read/write paths but is
  **not a full rewrite**: section writes, style editing, settings and comments are
  supported; deeper APIs (precise paragraph-format objects, field-code editing,
  OLE editing, ...) are not.
- `get_docx()` returns this library's facade object, **not** a python-docx
  `Document` — it cannot be passed to third-party code expecting python-docx types.

**Behavioral differences (equal or better outcomes)**

- Malformed-XML recovery: non-namespace `<` is escaped into text (producing
  schema-valid documents), unlike lxml recover which swallows text into unknown
  elements; text like `a<b` is treated as a tag by both.
- `== true` / `== none` comparisons are rewritten before compilation into
  equivalent custom tests (exact Python equality semantics, tag-scoped, string
  literals untouched).
- `{% trans %}` is a full gettext implementation (.mo parsing + plural rules),
  but does not cover every corner of jinja2's newstyle i18n.
- On MB-scale documents, 6 of docxtpl's regexes are replaced by linear scanners
  to avoid backtracking-stack overflows; outputs are byte-compared against the
  official test suite, but extremely malformed templates may differ.
- Zip output preserves compression methods and timestamps, but is not guaranteed
  byte-identical to python-docx output.

**Not supported**

- Custom jinja delimiters via `jinja_env` (unsupported by docxtpl itself too).
- jinja2 extensions other than i18n / do / loopcontrols / debug.
- WMF images (dropped by upstream python-docx 1.2 as well).

---

## 中文版

> ⚠️ **Vibecoding 项目**：本项目由 AI（Kimi Code）与用户结对编写，未经人工逐行审查。
> 已通过 docxtpl 官方测试套件（32 个真实模板用例）与 173 个自建测试验证，
> 但仍可能存在未发现的问题，生产环境使用前请自行评估。

[python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template) 的 **Rust 实现**，
既可作为 **Rust crate** 使用，也可通过 PyO3 作为 **Python 包** 使用（API 与 docxtpl 几乎完全一致）。
模板引擎为 minijinja（Jinja2 兼容语法）；docx 的 zip/XML 处理、模板预处理、表格修复、
关系管理、文档对象模型均为 Rust 原生实现，**不依赖** python-docx / lxml / jinja2。

### Rust 用法

```toml
[dependencies]
docxtplrs = { path = "../docxtplrs" }
minijinja = "2"
```

```rust
use docxtplrs::template::TplCore;
use minijinja::Value;
use std::collections::BTreeMap;

let mut ctx = BTreeMap::new();
ctx.insert("title".to_string(), Value::from("季度报告"));
ctx.insert("items".to_string(), Value::from_serialize(&vec!["甲", "乙"]));

let mut tpl = TplCore::new(std::fs::read("template.docx")?);
tpl.render(false, &|_core, _part| Ok(Value::from_serialize(&ctx)))?;
std::fs::write("out.docx", tpl.save_bytes()?)?;
```

主要模块：`template`（TplCore 渲染管线）、`package`（docx 包/关系/内容类型）、
`patch`（patch_xml 预处理）、`richtext`、`image`（图片尺寸/DPI）、`gettext`、
`xmldom`。可运行示例：`cargo run --example render --release -- template.docx out.docx`。

### Python 用法

```bash
uv sync                                   # 本仓库内构建安装
uv add /path/to/docxtplrs-*.whl           # 其他项目装 wheel
uv add --editable /path/to/docxtplrs      # 或本地开发模式
```

```python
from docxtplrs import DocxTemplate, R, InlineImage, Mm

tpl = DocxTemplate("template.docx")
tpl.render({
    "title": "季度报告",
    "rows": ["华东", "华南"],
    "rt": R("重点", bold=True, color="#C00000"),   # 模板写 {{r rt }}
    "img": InlineImage(tpl, "logo.png", width=Mm(15)),
    "sub": tpl.new_subdoc("chapter.docx"),
}, autoescape=True)
tpl.save("out.docx")
```

模板语法与 docxtpl 一致：`{{ var }}`、`{% if/for %}`、`{%tr %}`、`{%tc %}`、
`{%p %}`、`{{r }}`/`{{p }}`、`{% colspan %}`、`{% cellbg %}`、`{% vm %}`/`{% hm %}`、
`{% trans %}`（`tpl.install_gettext("zh.mo")`）、`{% include %}`（`tpl.set_template_loader`）、
页眉/页脚/脚注/文档属性渲染。

API 要点：`register_filter/test/function/global()`（支持 kwargs）、
`render(ctx, jinja_env=真实 jinja2 Environment)`、文档对象模型（可读写：
paragraphs/tables/sections/styles/settings/comments/core_properties、
add_paragraph/add_heading/add_picture/add_table/add_page_break/add_section、
run/单元格赋值）、替换类 API（replace_media/replace_pic/replace_embedded/replace_zipname）、
CLI（`python -m docxtplrs tpl.docx data.json out.docx -o`）。

### 限制

- 引擎为 minijinja：常见差异已补齐，极端 jinja2 边角不保证一致。
- 文档对象模型覆盖常用读写路径但非完整重写；`get_docx()` 返回本库外观对象，
  与 python-docx 类型不兼容。
- 非法 XML 容错策略与 lxml 不同但结果更好（产出 schema 合法文档）。
- `== true`/`== none` 编译前转换为等义自定义测试（语义一致）。
- 大文档上 6 处正则改为线性扫描器（避免栈溢出，输出已与官方套件比对一致）。
- 不支持自定义界定符（docxtpl 亦不支持）、部分 jinja2 扩展、WMF。

## Project structure / 项目结构

```
src/            Rust sources (17 modules, see AGENTS.md)   Rust 源码
python/         Python package shell (__init__ + __main__ CLI)
tests/          173 tests + crosscheck/compare scripts      测试与交叉验证
examples/       render.rs (Rust API), patch_dbg.rs (large-doc debugging)
AGENTS.md       Notes for AI coding assistants              给 AI 助手的项目说明
```

## License / 许可证

Same as the reference implementation python-docx-template: LGPL-2.1-or-later
(please verify compliance yourself before redistribution).
