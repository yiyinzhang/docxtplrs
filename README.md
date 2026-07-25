# docxtplrs

> ⚠️ **Vibecoding project**: this codebase was written by an AI (Kimi Code) in a
> pair-programming session with the user, without line-by-line human review.
> It is verified against the official docxtpl test suite (32 real-world template
> cases) plus 188 in-house tests, but undiscovered issues may remain — evaluate
> before production use.

A **Rust implementation of [python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template)**,
usable both as a **Rust crate** and as a **Python package** (via PyO3) with an API
almost identical to docxtpl. The template engine is
[minijinja](https://docs.rs/minijinja) (Jinja2-compatible syntax); all docx zip/XML
handling, template preprocessing, table fixing, relationship management and the
document object model are native Rust — **no dependency** on
python-docx / lxml / jinja2.

---

## Why docxtplrs instead of python-docx-template?

- **Zero Python dependencies** — the whole engine (docx zip/XML handling,
  template preprocessing, table fixing, document object model) is native Rust.
  No python-docx / lxml / jinja2 / markupsafe chain: one self-contained wheel,
  no lxml build issues on exotic platforms, no dependency-version conflicts.
- **Fast** — in a microbenchmark (title + paragraph loop + `{%tr %}` table-row
  loop over 10/100/400 items, 30 renders each, CPython 3.13, release build)
  docxtplrs renders **~6-14x faster** than docxtpl (lxml + jinja2): e.g.
  0.4/0.8/2.3ms vs 5.4/6.9/13.9ms per render at 10/100/400 rows. Reproduce
  with your own templates before quoting numbers.
- **Drop-in replacement** — the Python API and the template syntax
  (`{{ var }}`, `{%tr %}`/`{%tc %}`/`{%p %}`, `RichText`, `InlineImage`,
  `Subdoc`, ...) mirror docxtpl; migrating is usually just changing the import.
  Verified against the official docxtpl test suite (32 real-world templates),
  188 in-house tests, plus automated output cross-checking against docxtpl
  itself (`tests/crosscheck.py`).
- **One engine, two languages** — the same renderer is available as a native
  Rust crate, so Rust services/CLIs can render docx templates with no Python
  involved at all.
- **Engine extras beyond stock docxtpl** — `str` methods in templates
  (`upper/replace/split/...`), `'%s' % x` formatting, `{% break %}`/`{% continue %}`,
  the `{% do %}` statement, gettext i18n (`{% trans %}`/`{% pluralize %}`),
  `{% include %}`/`{% import %}` via `set_template_loader()`, custom
  filters/tests/functions/globals with keyword arguments, and interop with a
  real jinja2 `Environment`.
- **Built-in read-write document model** — paragraphs/tables/sections/styles/
  comments/core properties plus `add_paragraph/add_heading/add_picture/
  add_table/add_page_break/add_section` cover the common python-docx paths,
  without installing python-docx.
- **Built-in CLI** — `python -m docxtplrs template.docx data.json out.docx`
  renders a template straight from the shell, no scripting needed.

---

## Performance

Engine-vs-engine (docxtplrs vs docxtpl + lxml + jinja2, same templates):
**~6-14x faster** per render (see the "Fast" bullet above).

On top of that, a dedicated profiling round removed the remaining internal
hotspots (measured with the `maturin develop` debug build, CPython 3.13 —
release builds are faster still):

- **Document object model (live proxies)** — paragraph/run/table/section
  accessors used to re-parse the whole of `word/document.xml` on *every*
  attribute read or write. The parsed DOM is now cached and only serialized
  back at render/save time. Reading the text of 400 paragraphs:
  **1253ms → 3.6ms (~350x)**; 200 × `add_paragraph`: **832ms → 7ms (~120x)**;
  200 × run-text edits: **801ms → 0.4ms (~2000x)**.
- **Render pipeline** — every XML patch pass is now gated (a zero-match pass
  no longer copies the full document), `InlineImage`/`Subdoc` placeholders are
  materialized in a single scan with `Arc`-shared blobs, and the
  table/docPr fix skips its DOM round-trip when there is nothing to fix.
  Rendering a 2MB / 20k-paragraph document: **3384ms → 1901ms (-44%)**.
- **Images & subdoc merge** — image-dedup hashes are cached (per-insert cost
  no longer re-hashes all media), `.rels` parts are parsed once per part
  instead of once per relationship, and rId/style/numId remapping runs as a
  single scan with a map lookup instead of one full-text regex per entry.

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

### Building the wheel yourself

```bash
# release wheel -> target/wheels/docxtplrs-<ver>-cp313-cp313-manylinux_2_34_x86_64.whl
uv run maturin build --release

# on machines that also have conda installed:
env -u CONDA_PREFIX uv run maturin build --release
```

Variants:

```bash
# portable abi3 wheel (one file for all CPython >= 3.12):
#   1. in Cargo.toml set  pyo3 = { version = "0.29", features = ["abi3-py312"] }
#   2. build
uv run maturin build --release        # -> docxtplrs-...-cp312-abi3-....whl

# wheel for a specific Python version: point maturin at that interpreter
uv run maturin build --release -i python3.12

# also emit the sdist (source distribution) alongside the wheel
uv run maturin build --release --sdist

# then install / distribute
uv pip install target/wheels/*.whl                     # local install
# or upload to your index, e.g.: uv publish target/wheels/*   (or: twine upload)
```

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

### Custom filters & engine extensions

Register your own jinja filters, tests, functions and globals — plain Python
callables, with positional **and keyword** arguments supported:

```python
tpl = DocxTemplate("template.docx")

# filters: {{ price|rmb }}, {{ name|shout('!') }}, {{ rows|col(name='药物名称') }}
tpl.register_filter("rmb", lambda v: f"¥{v:.2f}")
tpl.register_filter("shout", lambda v, punct="?": str(v).upper() + punct)
tpl.register_filter("col", lambda rows, name: [r[name] for r in rows])

# tests: {% if v is even %}
tpl.register_test("even", lambda v: v % 2 == 0)

# functions & globals: {{ add(1, 2) }}, {{ company }}
tpl.register_function("add", lambda a, b: a + b)
tpl.register_global("company", "ACME")

# template loader for {% include %}/{% import %}
tpl.set_template_loader(lambda name: "included {{ v }}" if name == "part" else None)

# gettext i18n for {% trans %}/{% pluralize %}
tpl.install_gettext("messages.mo")

tpl.render(context)
```

You can also pass a real **jinja2 Environment** — its `autoescape`, `filters`,
`globals`, `tests`, `trim_blocks`, `lstrip_blocks`, `keep_trailing_newline` and
`undefined` (Chainable/Strict) settings are honored (duck-typed):

```python
import jinja2
env = jinja2.Environment()
env.filters["usd"] = lambda v: f"${v:.2f}"
tpl.render(context, jinja_env=env)
```

Notes: jinja2's own builtin filters/tests/globals are **not** imported — they are
detected by identity against `jinja2.defaults` and handled by minijinja's native
implementations (so undefined-value semantics like `default`/`defined` stay
correct); only entries you added or overrode are imported, and they take priority
over same-named `register_filter/test/function/global()` registrations.

```python
import jinja2
env = jinja2.Environment()
env.filters["usd"] = lambda v: f"${v:.2f}"
tpl.render(context, jinja_env=env)
```

Engine-level extras beyond stock minijinja, so jinja2-style templates work as-is:

- **str methods**: `upper/capitalize/title/strip/replace/split/join/zfill/center/
  partition/splitlines/format(...)` etc., including format specs (`{:>8}`, `{:,}`,
  `{:.2f}`, `{:#x}`, `{:.1%}`)
- **printf operator**: `'%s-%d' % (a, 2)`
- **extra filters**: `filesizeformat`, `wordcount`, `center`, `forceescape`,
  `truncate`, `xmlattr`, `wordwrap`, `urlize`, `random`, `striptags`
- **tests**: `true/false/boolean/none/callable/mapping/sequence` (Python-faithful),
  `eq_true/eq_false` (exact `== True`/`== False` semantics)
- **Python object interop**: attribute/item access, method calls, iteration,
  `__len__` truthiness, `__eq__/__lt__` comparisons, `__html__()` protocol,
  dict insertion-order iteration, i128 big integers, `True/False/None` rendered
  Python-style, `in` containment on dicts/lists

### Tests

```bash
.venv/bin/python -m pytest tests/ -q          # 188 unit tests
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
> 已通过 docxtpl 官方测试套件（32 个真实模板用例）与 188 个自建测试验证，
> 但仍可能存在未发现的问题，生产环境使用前请自行评估。

[python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template) 的 **Rust 实现**，
既可作为 **Rust crate** 使用，也可通过 PyO3 作为 **Python 包** 使用（API 与 docxtpl 几乎完全一致）。
模板引擎为 minijinja（Jinja2 兼容语法）；docx 的 zip/XML 处理、模板预处理、表格修复、
关系管理、文档对象模型均为 Rust 原生实现，**不依赖** python-docx / lxml / jinja2。

### 相比 python-docx-template 的优势

- **零 Python 依赖**：整个引擎（docx zip/XML 处理、模板预处理、表格修复、文档对象模型）
  均为 Rust 原生实现。不再需要 python-docx / lxml / jinja2 / markupsafe 依赖链——单个自包含
  wheel，没有 lxml 在冷门平台上的编译问题，也没有依赖版本冲突的烦恼。
- **快**：微基准测试（标题 + 段落循环 + `{%tr %}` 表格行循环，10/100/400 行，
  各渲染 30 次，CPython 3.13，release 构建）下 docxtplrs 比 docxtpl（lxml + jinja2）
  **快约 6-14 倍**：10/100/400 行时每次渲染 0.4/0.8/2.3ms vs 5.4/6.9/13.9ms。
  引用数字前请用你自己的模板复测。
- **无缝替代**：Python API 与模板语法（`{{ var }}`、`{%tr %}`/`{%tc %}`/`{%p %}`、
  `RichText`、`InlineImage`、`Subdoc` 等）与 docxtpl 几乎完全一致，迁移通常只需改一行 import。
  已通过 docxtpl 官方测试套件（32 个真实模板）、188 个自建测试，并有与 docxtpl 输出自动
  交叉比对的工具（`tests/crosscheck.py`）。
- **一个引擎，两种语言**：同一渲染器同时提供原生 Rust crate，Rust 服务/CLI 可以完全不
  经过 Python 渲染 docx 模板。
- **超出原版 docxtpl 的引擎能力**：模板内 `str` 方法（`upper/replace/split/...`）、
  `'%s' % x` 格式化、`{% break %}`/`{% continue %}`、`{% do %}` 语句、gettext i18n
  （`{% trans %}`/`{% pluralize %}`）、`set_template_loader()` 支持的 `{% include %}`/
  `{% import %}`、带关键字参数的自定义 filters/tests/functions/globals，以及与真实
  jinja2 `Environment` 的互操作。
- **内置可读写文档对象模型**：paragraphs/tables/sections/styles/comments/core properties
  以及 `add_paragraph/add_heading/add_picture/add_table/add_page_break/add_section`，
  覆盖 python-docx 常用读写路径，无需再安装 python-docx。
- **内置 CLI**：`python -m docxtplrs template.docx data.json out.docx`，无需写脚本即可
  在命令行直接渲染模板。

### 性能

引擎对引擎（docxtplrs vs docxtpl + lxml + jinja2，相同模板）：每次渲染**快约 6-14 倍**
（见上方"快"一条）。

在此之上，最近一轮专项 profiling 又清掉了剩余的内部热点（数据基于 `maturin develop`
debug 构建 + CPython 3.13 实测，release 构建会更快）：

- **文档对象模型（实时代理）**：段落/表格/节等代理此前**每次**属性读写都重新全量解析
  `word/document.xml`；现在解析结果带脏标记缓存，仅在 render/save 时写回。读取 400 个
  段落文本：**1253ms → 3.6ms（约 350 倍）**；200 次 `add_paragraph`：**832ms → 7ms
  （约 120 倍）**；200 次 run 文本改写：**801ms → 0.4ms（约 2000 倍）**。
- **渲染管线**：所有 XML patch pass 加了 gate（零匹配不再全量拷贝文档），
  `InlineImage`/`Subdoc` 占位符单遍扫描实体化且 blob 用 `Arc` 共享，表格/docPr 修复
  在无事可修时跳过整个 DOM 往返。渲染 2MB / 2 万段落文档：**3384ms → 1901ms（-44%）**。
- **图片与 subdoc 合并**：图片去重哈希带缓存（插一张图不再重算全部媒体的 sha1），
  `.rels` 每个 part 只解析一次（而非每条关系一次），rId/style/numId 重映射从
  "每条目一次全文正则"改为单遍扫描 + map 查表。

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

### 构建 wheel

```bash
# release 构建（产物在 target/wheels/，含 manylinux 分发包）
uv run maturin build --release

# 本机装有 conda 时
env -u CONDA_PREFIX uv run maturin build --release

# abi3 通用 wheel（一份通吃 CPython ≥ 3.12）：
#   先把 Cargo.toml 的 pyo3 改为 features = ["abi3-py312"]，再构建

# 指定 Python 版本构建
uv run maturin build --release -i python3.12

# 同时产出 sdist
uv run maturin build --release --sdist

# 安装 / 分发
uv pip install target/wheels/*.whl        # 本地安装
# 或上传到自有 PyPI 索引后 uv add docxtplrs --index-url ...
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
`render(ctx, jinja_env=真实 jinja2 Environment)`（jinja2 内建 filters/tests/globals 由 minijinja 原生实现接管，仅导入用户新增/覆盖的条目，且同名优先级高于 register_*）、文档对象模型（可读写：
paragraphs/tables/sections/styles/settings/comments/core_properties、
add_paragraph/add_heading/add_picture/add_table/add_page_break/add_section、
run/单元格赋值）、替换类 API（replace_media/replace_pic/replace_embedded/replace_zipname）、
CLI（`python -m docxtplrs tpl.docx data.json out.docx -o`）。

### 自定义过滤器与引擎扩展

```python
tpl = DocxTemplate("template.docx")

# 过滤器（支持位置参数和关键字参数）
tpl.register_filter("rmb", lambda v: f"¥{v:.2f}")                    # {{ price|rmb }}
tpl.register_filter("shout", lambda v, punct="?": str(v).upper() + punct)  # {{ name|shout('!') }}
tpl.register_filter("col", lambda rows, name: [r[name] for r in rows])     # {{ rows|col(name='药物名称') }}

# 测试器 / 函数 / 全局值
tpl.register_test("even", lambda v: v % 2 == 0)          # {% if v is even %}
tpl.register_function("add", lambda a, b: a + b)         # {{ add(1, 2) }}
tpl.register_global("company", "ACME")                   # {{ company }}

# 模板 loader（{% include %}/{% import %}）与 gettext i18n（{% trans %}/{% pluralize %}）
tpl.set_template_loader(lambda name: "included {{ v }}" if name == "part" else None)
tpl.install_gettext("messages.mo")

tpl.render(context)
```

也可直接传真正的 **jinja2 Environment**（自动读取其 `autoescape` / `filters` /
`globals` / `tests` / `trim_blocks` / `lstrip_blocks` / `keep_trailing_newline` /
`undefined`（Chainable/Strict）设置）：

```python
import jinja2
env = jinja2.Environment()
env.filters["usd"] = lambda v: f"${v:.2f}"
tpl.render(context, jinja_env=env)
```

引擎层对 jinja2 的补齐（模板无需改动即可使用）：

- **str 方法**：`upper/capitalize/title/strip/replace/split/join/zfill/center/
  partition/splitlines/format(...)` 等，含格式规格（`{:>8}`、`{:,}`、`{:.2f}`、
  `{:#x}`、`{:.1%}`）
- **printf 运算符**：`'%s-%d' % (a, 2)`
- **增补过滤器**：`filesizeformat`、`wordcount`、`center`、`forceescape`、
  `truncate`、`xmlattr`、`wordwrap`、`urlize`、`random`、`striptags`
- **增补测试**：`true/false/boolean/none/callable/mapping/sequence`（Python 语义）、
  `eq_true/eq_false`（精确的 `== True`/`== False` 语义）
- **Python 对象互操作**：属性/下标访问、方法调用、迭代、`__len__` 真值、
  `__eq__/__lt__` 比较、`__html__()` 协议、dict 插入序迭代、i128 大整数、
  `True/False/None` 按 Python 风格渲染、`in` 包含运算

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
tests/          188 tests + crosscheck/compare scripts      测试与交叉验证
examples/       render.rs (Rust API), patch_dbg.rs (large-doc debugging)
AGENTS.md       Notes for AI coding assistants              给 AI 助手的项目说明
```

## License / 许可证

Same as the reference implementation python-docx-template: LGPL-2.1-or-later
(please verify compliance yourself before redistribution).
