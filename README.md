# docxtplrs

[![crates.io](https://img.shields.io/crates/v/docxtplrs.svg)](https://crates.io/crates/docxtplrs)
[![docs.rs](https://docs.rs/docxtplrs/badge.svg)](https://docs.rs/docxtplrs)
[![license](https://img.shields.io/badge/license-LGPL--2.1--or--later-blue.svg)](https://github.com/yiyinzhang/docxtplrs)
[![tests](https://img.shields.io/badge/tests-445%20passed-brightgreen.svg)](https://github.com/yiyinzhang/docxtplrs)
[![coverage](https://img.shields.io/badge/coverage-86%25-brightgreen.svg)](https://github.com/yiyinzhang/docxtplrs)
[![crosscheck](https://img.shields.io/badge/crosscheck-ALL%20MATCH-brightgreen.svg)](https://github.com/yiyinzhang/docxtplrs)

**[English](#english) · [中文版](#中文版)**

A **Rust implementation of [python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template)** —
render docx templates with Jinja2 syntax ([minijinja](https://docs.rs/minijinja)),
usable as a **Rust crate** and as a **Python package** (PyO3). All docx zip/XML
handling, template preprocessing, table fixing, relationship management and the
document object model are native Rust: **no dependency** on python-docx / lxml / jinja2.

[python-docx-template (docxtpl)](https://github.com/elapouya/python-docx-template) 的 **Rust 实现**——
用 Jinja2 语法（minijinja）渲染 docx 模板，既可作为 **Rust crate**，也可作为 **Python 包**（PyO3）。
docx 的 zip/XML 处理、模板预处理、表格修复、关系管理、文档对象模型均为 Rust 原生实现，
**不依赖** python-docx / lxml / jinja2。

> ⚠️ Vibecoding project / AI 结对编写项目：written by an AI (Kimi Code) with the user,
> without line-by-line human review / 未经人工逐行审查。Verified by the official docxtpl
> test suite (32 real-world templates) + 445 in-house tests / 已通过官方套件与 445 个自建测试，
> but evaluate before production use / 生产使用前请自行评估。

## Quick start / 快速开始

**Rust** — `cargo add docxtplrs` ([crates.io](https://crates.io/crates/docxtplrs)):

```rust
use docxtplrs::template::TplCore;
use minijinja::Value;
use std::collections::BTreeMap;

let mut ctx = BTreeMap::new();
ctx.insert("title".to_string(), Value::from("Quarterly Report"));
ctx.insert("items".to_string(), Value::from_serialize(&vec!["alpha", "beta"]));

let mut tpl = TplCore::new(std::fs::read("template.docx")?);
tpl.render(false, &|_core, _part| Ok(Value::from_serialize(&ctx)))?;
std::fs::write("out.docx", tpl.save_bytes()?)?;
```

**Python** — build the wheel with maturin / 用 maturin 构建 wheel:

```bash
uv sync                      # in this repo / 本仓库内
uv add --editable .          # or into another project / 或装进其他项目
```

```python
from docxtplrs import DocxTemplate

tpl = DocxTemplate("template.docx")
tpl.render({"title": "Quarterly Report", "rows": ["East", "South"]})
tpl.save("out.docx")
```

---

## English

**Contents** — [Why docxtplrs?](#why-docxtplrs) · [Performance](#performance) ·
[Rust crate](#1-usage-as-a-rust-crate) · [Python package](#2-usage-as-a-python-package) ·
[Limitations](#3-limitations) · [中文版](#中文版)

### Why docxtplrs?

- **Zero Python dependencies** — the whole engine (docx zip/XML handling,
  template preprocessing, table fixing, document object model) is native Rust.
  No python-docx / lxml / jinja2 / markupsafe chain: one self-contained wheel,
  no lxml build issues on exotic platforms, no dependency-version conflicts.
- **Fast** — **~6-12x faster** than docxtpl (lxml + jinja2) on small templates
  (title + paragraph loop + `{%tr %}` table-row loop over 10/100/400 items,
  30 renders each: 0.3-0.5/0.8-1.1/3.2-4.0ms vs 4.3-4.8/7.7-9.3/20-25ms per
  render), and **~60x** on an 8.9MB / 20k-paragraph document (0.3s vs ~20s
  for a full load-render-save cycle). CPython 3.13, release build; reproduce with
  `tests/benchmark.py [--big]` — and with your own templates before quoting
  numbers; see [Performance](#performance) for the internals.
- **Drop-in replacement** — the Python API and the template syntax
  (`{{ var }}`, `{%tr %}`/`{%tc %}`/`{%p %}`, `RichText`, `InlineImage`,
  `Subdoc`, ...) mirror docxtpl; migrating is usually just changing the import.
  Verified against the official docxtpl test suite (32 real-world templates),
  445 in-house tests, plus automated output cross-checking against docxtpl
  itself (`tests/crosscheck.py`).
- **One engine, two languages** — the same renderer is available as a native
  Rust crate, so Rust services/CLIs can render docx templates with no Python
  involved at all.
- **Document composition** — `Composer` concatenates whole docx files
  (docxcompose-style: page break between documents, style conflicts renamed,
  list numbering restarted, media deduped), from both Python and pure Rust.
  Going beyond docxcompose, appended documents keep their **section
  properties** (page size/orientation/margins and header/footer references,
  intermediate sections of multi-section documents included), **comments**
  (with w15 threading state) are merged, and styles/numbering referenced
  from inside headers/footers are merged consistently with the body.
- **Engine extras beyond stock docxtpl** — `str` methods in templates
  (`upper/replace/split/...`), `'%s' % x` formatting, `{% break %}`/`{% continue %}`,
  the `{% do %}` statement, gettext i18n (`{% trans %}`/`{% pluralize %}`),
  `{% include %}`/`{% import %}` via `set_template_loader()`, custom
  filters/tests/functions/globals with keyword arguments, and interop with a
  real jinja2 `Environment`.
- **Built-in read-write document model** — paragraphs/tables/sections/styles/
  comments/core properties plus `add_paragraph/add_heading/add_picture/
  add_table/add_page_break/add_section` cover the common python-docx paths,
  without installing python-docx — and the same model is available to pure
  Rust consumers via the `doc` module (the Python bindings are thin wrappers
  over it). Full `Font` (29 properties with python-docx
  tri-state semantics), `ParagraphFormat` (indents/spacing/line rules/keep
  flags/tab stops), table columns/alignment/autofit, row heights, cell
  vertical alignment, section start types and field codes (`add_field`,
  `update_fields_on_open`) are all supported.
- **Built-in CLI** — `python -m docxtplrs template.docx data.json out.docx`
  renders a template straight from the shell, no scripting needed.

### Performance

Engine-vs-engine (docxtplrs vs docxtpl + lxml + jinja2, same templates):
**~6-12x faster** per render on small templates and **~60x** on a large
(8.9MB) document (see the "Fast" bullet above).

On top of that, dedicated profiling rounds removed the remaining internal
hotspots (measured with the `maturin develop` debug build unless noted,
CPython 3.13 — release builds are faster still):

- **Document object model (live proxies)** — paragraph/run/table/section
  accessors used to re-parse the whole of `word/document.xml` on *every*
  attribute read or write. The parsed DOM is now cached and only serialized
  back at render/save time. Reading the text of 400 paragraphs:
  **1253ms → 3.6ms (~350x)**; 200 × `add_paragraph`: **832ms → 7ms (~120x)**;
  200 × run-text edits: **801ms → 0.4ms (~2000x)**.
- **Per-part DOM cache** — `styles.xml` / `settings.xml` / `comments.xml` /
  `core.xml` (and every `XmlElement` proxy part) used to be re-parsed, and
  mutations re-serialized, on *every* access. They now share the same
  parse-once/flush-on-save cache as `document.xml`. 20 × `styles[name]`
  lookups over 200 styles: **21.0ms → 0.41ms (~50x)**; 20 × style-font
  writes: **8.6ms → 0.02ms (~430x)** (release build); `resolve_style_id` no
  longer deep-clones every style node per call.
- **One env + one context per render** — the minijinja `Environment` (40+
  filter/test registrations) and the Python-context → `Value` conversion used
  to be rebuilt for *every* part (body, each header/footer, footnotes, core
  properties). Both are now built once per `render()` call. Rendering a
  7-part document with a 300-variable context: **5.0ms → 1.9ms (-61%)**
  (release build).
- **Render pipeline** — every XML patch pass is now gated (a zero-match pass
  no longer copies the full document), `InlineImage`/`Subdoc` placeholders are
  materialized in a single scan with `Arc`-shared blobs, and the
  table/docPr fix skips its DOM round-trip when there is nothing to fix.
  Rendering a 2MB / 20k-paragraph document: **3384ms → 1901ms (-44%)**.
- **Large-document render path** — the two dominant `patch_xml` regexes
  (split-tag stripping, tag cleaning) and the paragraph-newline handling are
  hand-rolled linear scanners, all pipeline passes return `Cow` (zero-copy
  when a pass changes nothing), and xmldom got single-allocation text nodes,
  bulk escaping and preallocated serialization; subdoc merges renumber
  `pic:cNvPr` ids inside the same DOM round-trip instead of a second
  parse+serialize. On the 8.9MB / 20k-paragraph document this cut render
  time by a further **~60%** (Rust, release build) — the same input takes
  docxtpl ~20s.
- **Images & subdoc merge** — image-dedup hashes are cached (per-insert cost
  no longer re-hashes all media), `.rels` parts are parsed once per part
  instead of once per relationship, and rId/style/numId remapping runs as a
  single scan with a map lookup instead of one full-text regex per entry.
- **Zip & template caching round** — media payloads (png/jpg/...) are stored
  without re-deflating on save, `render()` reloads clone the pristine package
  instead of re-inflating the whole zip, compiled jinja templates are cached
  per part across renders (content-hash keyed), the three always-firing
  `patch_xml` passes run as one fused scan, and the table/docPr fix skips its
  DOM round-trip for table-free documents. On a 4.2MB template with 4×1MB
  embedded images: save **204ms → 7ms (~29x)**, repeated render
  **2.5ms → 0.9ms (~3x)**, full cycle **177ms → 11ms (~16x)** (release build;
  `tests/bench_media.py`).
- **Bridge & listing round** — `resolve_listing` (newline/tab expansion) no
  longer runs per-paragraph fancy-regexes (one linear scan instead),
  zip-level media replacements no longer clone the whole package, media
  blobs are `Arc`-shared and hashed once, and the Python bridge converts
  scalars through a fast path (no TLS lookup, no failed casts, dict misses
  without exceptions). Listing-heavy render (300 multi-line cells):
  **8.9ms → 4.8ms (~1.9x)** vs the pre-P0 baseline (`tests/bench_listing.py`).
- **Concurrency & model round** — `render()`/`save()` now run detached from
  the GIL (re-acquired on demand for Python callbacks), docmodel proxies use
  generation-validated access cursors instead of rescanning from the body
  head (reading 2000 paragraph texts sequentially: **17.4ms → 0.7ms
  (~25x)**), `.rels` lookups are `Rc`-shared, style-id resolution and
  subdoc bookmark renumbering are cached, and the render pipeline skips
  footnotes/core-property work when there are no template tags.
- **Scanner & allocation round** — the remaining hot scanners are
  SIMD-vectorized (`memchr` jumps between candidate bytes instead of
  per-byte loops), the row/paragraph/run element strip locates the rare
  jinja markers first instead of probing every element open, and
  `resolve_listing` rewrites only paragraphs that actually contain listing
  characters, so multi-MB documents are bulk-copied between hits. xmldom
  element/attribute names are interned (`Rc<str>` over the few dozen
  distinct names in a docx: parse allocations **740k → 445k, -40%**), and
  the fused jinja emitter reuses its per-tag buffers (`patch_xml`
  allocations **20.5k → 31**). Zip entries keep `Arc`-shared payloads — a
  per-render reload is a refcount bump instead of a full copy — and
  already-compressed media stay as the raw deflate stream end-to-end
  (inflated lazily on first access, raw-copied back at save); compression
  runs on the faster zlib-ng backend, and `get_docx_bytes()`/`get_xml()`
  also run detached from the GIL. The pure per-part preprocessing chain
  (entity decode + patch_xml + rewrites) is cached keyed by raw source
  hash, so a repeated `render()` skips it entirely. On the 8.9MB document
  (Rust, release build): patch_xml **148ms → 57ms**, table/docPr fix
  **234ms → 118ms**, save **103ms → ~35ms**, e2e render
  **673ms → 348ms (-48%)**; repeated render on the 4.2MB media template
  **0.9ms → 0.4ms**.
- **Structural scan & bridge round** — the table/docPr fix first runs a
  read-only string scan that decides whether any table grid actually needs
  changes (exactly replicating `fix_one_table`'s decision logic, nested
  tables and gridSpan sums included) and skips the full DOM parse+walk when
  grids are consistent — which they almost always are after rendering
  (**118ms → 16ms** on the 8.9MB document). Element-tag markers
  (`{%tr %}`/`{%p %}`/…) are collected in a single aho-corasick sweep
  instead of 11 memmem gate scans (patch_xml **57ms → 35ms**), the fused
  jinja emitter locates exact `<w:t>` tags with memmem instead of checking
  every `<`, unmodified package entries are raw-copied at save (not just
  media), and part sources are decoded zero-copy on the UTF-8 fast path.
  On the Python side, dict attribute access is materialized once per dict
  (one GIL attach) instead of once per lookup: re-rendering the 8.9MB
  document **185ms → 50ms**, full cycle **367ms → 226ms**, and the
  small-template speedup vs docxtpl is now **~15-28x**. Overall e2e render
  on the 8.9MB benchmark: **348ms → ~190ms** (Rust, release build).

### 1. Usage as a Rust crate

Add to your `Cargo.toml`:

```toml
[dependencies]
docxtplrs = { version = "0.2.2", default-features = false }  # pure Rust: no PyO3/libpython
minijinja = "2"
```

Cargo features:

| Feature | Default | Purpose |
|---|---|---|
| `python` | ✓ | PyO3 bindings and Python interop (Python callables as custom filters/tests/functions/globals, `set_template_loader`). Disable it (`default-features = false`) for a pure-Rust build with no `pyo3`/`libpython` dependency. |
| `extension-module` | – | Only enabled by maturin when building the Python wheel (implies `python`); never enable it for a Rust binary/library. |

Render a template (see [Quick start](#quick-start--快速开始) for the code).

Edit a document with the pure-Rust object model (no Python involved):

```rust
use docxtplrs::{doc, template::TplCore};

let mut tpl = TplCore::new(std::fs::read("in.docx")?);
let p = doc::paragraphs(&mut tpl)[0];
p.set_alignment(&mut tpl, Some(1))?;            // center
let run = p.add_run(&mut tpl, "important")?;
run.font().set_bold(&mut tpl, Some(true))?;      // tri-state like python-docx
doc::add_table(&mut tpl, 2, 2)?;
std::fs::write("out.docx", tpl.save_bytes()?)?;
```

Useful Rust API surface (all in `docxtplrs::`):

| Module | Highlights |
|---|---|
| `template` | `TplCore`: `new/render/save_bytes/get_xml/undeclared_variables/build_url_id`, `Deferred`, replacement maps (`crc_to_new_media`, `pics_to_replace`, …) |
| `package` | `Package` (docx zip entries, `.rels`, content types), `Rels`, CRC32/SHA1 helpers |
| `patch` | `patch_xml`, `resolve_listing`, `decode_text_entities` (docxtpl's XML preprocessing) |
| `composer` | `Composer`: docxcompose-style whole-document concatenation (`new`/`append`/`save_bytes`) |
| `doc` | Pure-Rust document object model: `Paragraph/Run/Table/TableRow/Cell/Section/Style` index handles over `TplCore` (`text`/`style`/`alignment` read-write, `Font`/`ParagraphFormat` handles, `add_paragraph/add_table/...` free fns) — the same model the Python bindings wrap |
| `richtext` | `richtext_run`, `richtext_paragraph`, `listing_xml`, `TextProps` |
| `image` | `ImageInfo` (PNG/JPEG/GIF/BMP/TIFF size & DPI), EMU conversions |
| `gettext` | `Catalog` (.mo parsing, plural-rule evaluation) |
| `xmldom` | Minimal XML DOM used for structured edits |

A runnable example is included:

```bash
cargo run --example render --release -- template.docx out.docx
```

> Note: with default features the examples link against `libpython` (PyO3). If your Python is uv-managed, run with
> `LD_LIBRARY_PATH=$(python -c "import sysconfig; print(sysconfig.get_config_var('LIBDIR'))")`.
> Built with `--no-default-features` (e.g. `cargo run --no-default-features --example render --release -- ...`)
> the examples are pure Rust and need no Python at all.

### 2. Usage as a Python package

#### Install

```bash
# in this repo: build & install into the uv venv
uv sync

# in another uv project
uv add /path/to/docxtplrs/target/wheels/docxtplrs-0.2.2-cp313-cp313-manylinux_2_17_x86_64.manylinux2014_x86_64.whl
# or editable (local development)
uv add --editable /path/to/docxtplrs
```

The wheel is a native `cp313` build. For other Python versions either build under that
version, or set pyo3 to `abi3-py312` in `Cargo.toml` for a single portable wheel.

#### Building the wheel yourself

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

#### Template syntax (same as docxtpl)

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

Full render example:

```python
from docxtplrs import DocxTemplate, R, RichText, Listing, InlineImage, Mm

tpl = DocxTemplate("template.docx")          # path / file-like / bytes

tpl.render({
    "title": "Quarterly Report",
    "rows": ["East", "South"],                # {%p for r in rows %} paragraph loop
    "items": [{"name": "A", "price": 12.5}],  # {%tr for x in items %} table row loop
    "rt": R("important", bold=True, color="#C00000"),  # write {{r rt }} in the template
    "img": InlineImage(tpl, "logo.png", width=Mm(15)), # {{ img }}
    "sub": tpl.new_subdoc("chapter.docx"),             # {{p sub }}; keep_sections=True
                                                     # preserves the subdoc's page setup
}, autoescape=True)

tpl.save("out.docx")                         # or io.BytesIO()
```

#### Python API summary

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
  full `Font` (29 tri-state properties) and `ParagraphFormat` (indents,
  spacing, line rules, keep flags, tab stops) on runs/styles/paragraphs,
  table `columns/add_column/alignment/autofit`, row `height/height_rule`,
  cell `paragraphs/width/vertical_alignment`, section `start_type/distances`,
  style flags (`hidden/locked/priority/quick_style/builtin/...`),
  `insert_paragraph_before`/`clear`/`add_break`/`add_tab`/`add_text`,
  `hyperlinks`, `iter_inner_content` (document/section/paragraph/run/cell),
  `contains_page_break`/`rendered_page_breaks`, `part` facade,
  field codes (`paragraph.fields`/`add_field`, `instr`/cached-result read-write,
  `settings.update_fields_on_open`),
  `run.add_picture`/`mark_comment_range`, `table_direction`,
  `grid_span`/`grid_cols_before/after`, nested tables in cells,
  `__getattr__` delegation
- Raw-XML escape hatch: `settings/document/styles.element` return a live
  `XmlElement` proxy (`tag/text/attrib/get/set/remove_attr/find/findall/
  children/append/insert/remove/xml`), covering python-docx's `.element` use
  cases (e.g. `settings.element` for `w:zoom`)
- `Table.style`, `Cell.merge()` (gridSpan/vMerge, python-docx semantics),
  `Section.even_page_header/footer`, `Section.first_page_header/footer`
- docxtpl helpers: `HEADER_URI`/`FOOTER_URI`, `get_file_crc()` (path/bytes/
  file-like), `get_headers_footers(uri)` yielding live `(rid, XmlElement)` pairs
- `RichText`/`R`, `RichTextParagraph`/`RP`, `Listing`, `InlineImage`, `Subdoc`
  (file merging + programmatic `add_paragraph/add_picture/add_table`;
  `new_subdoc(path, keep_sections=True)` additionally preserves the subdoc's
  section properties — page setup and header/footer references — as its own
  section; subdoc **comments are merged** in all cases)
- `Composer`: docxcompose-style concatenation of whole documents —
  `Composer(master).append(doc).save(out)`, with a page break between
  documents and style/numbering/media merging:

  ```python
  from docxtplrs import Composer

  c = Composer("master.docx")   # path / file-like / bytes
  c.append("chapter1.docx")     # same input kinds
  c.append("chapter2.docx")
  c.save("out.docx")            # or io.BytesIO()
  ```
- Units: `Length/Emu/Inches/Cm/Mm/Pt/Twips`; jinja2.utils: `Cycler/Joiner/
  generate_lorem_ipsum`; exception: `TemplateError` (with `docx_context`)
- CLI: `python -m docxtplrs template.docx data.json output.docx [-o] [-q]`

#### Custom filters & engine extensions

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

#### Tests

```bash
.venv/bin/python -m pytest tests/ -q          # 445 unit tests
.venv/bin/python tests/crosscheck.py docxtplrs > /tmp/rs.json
.venv/bin/python tests/crosscheck.py docxtpl > /tmp/ref.json   # needs crosscheck group
.venv/bin/python tests/crosscheck.py compare /tmp/ref.json /tmp/rs.json
.venv/bin/python tests/benchmark.py --big     # engine-vs-engine benchmark
```

Coverage badge: line coverage of `src/` measured by rebuilding with
`RUSTFLAGS="-Cinstrument-coverage"`, running the pytest suite and reporting
with `llvm-cov`.

### 3. Limitations

**Architectural**

- The engine is minijinja, not jinja2: broadly compatible, but exotic jinja2 edge
  behavior is not guaranteed (common gaps already covered with custom tests/filters,
  loader, Unicode identifiers, kwargs support, etc.).
- The document object model covers python-docx's common read/write paths
  (including full `Font`/`ParagraphFormat` objects, table/row/cell geometry,
  style flags, `rendered_page_breaks` markers, `iter_inner_content`, a
  minimal `part` facade and field codes) but is **not a full rewrite**:
  OLE editing is not supported.
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
- On MB-scale documents, 8 of docxtpl's regexes are replaced by linear scanners
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

**目录** — [相比 docxtpl 的优势](#相比-python-docx-template-的优势) · [性能](#性能) ·
[Rust 用法](#1-rust-用法) · [Python 用法](#2-python-用法) · [限制](#3-限制) ·
[English](#english)

### 相比 python-docx-template 的优势

- **零 Python 依赖**：整个引擎（docx zip/XML 处理、模板预处理、表格修复、文档对象模型）
  均为 Rust 原生实现。不再需要 python-docx / lxml / jinja2 / markupsafe 依赖链——单个自包含
  wheel，没有 lxml 在冷门平台上的编译问题，也没有依赖版本冲突的烦恼。
- **快**：小模板（标题 + 段落循环 + `{%tr %}` 表格行循环，10/100/400 行，各渲染 30 次）
  下比 docxtpl（lxml + jinja2）**快约 6-12 倍**：每次渲染 0.3-0.5/0.8-1.1/3.2-4.0ms vs
  4.3-4.8/7.7-9.3/20-25ms；8.9MB / 2 万段落大文档上**快约 60 倍**（完整加载-渲染-保存
  周期 0.3s vs 约 20s）。CPython 3.13，release 构建；可用 `tests/benchmark.py [--big]`
  复现——引用数字前请用你自己的模板复测；引擎内部优化见[性能](#性能)。
- **无缝替代**：Python API 与模板语法（`{{ var }}`、`{%tr %}`/`{%tc %}`/`{%p %}`、
  `RichText`、`InlineImage`、`Subdoc` 等）与 docxtpl 几乎完全一致，迁移通常只需改一行 import。
  已通过 docxtpl 官方测试套件（32 个真实模板）、445 个自建测试，并有与 docxtpl 输出自动
  交叉比对的工具（`tests/crosscheck.py`）。
- **一个引擎，两种语言**：同一渲染器同时提供原生 Rust crate，Rust 服务/CLI 可以完全不
  经过 Python 渲染 docx 模板。
- **文档拼接**：`Composer` 把多个完整 docx 首尾拼接为一个文档（docxcompose 风格：
  文档间分页、样式冲突改名、列表编号重启、媒体去重），Python 与纯 Rust 两侧均可用。
  比 docxcompose 更进一步：被拼接文档的**节属性**（页面尺寸/方向/页边距及页眉页脚
  引用，多节文档的中间节同样保留）会被保留，**批注**（含 w15 线程状态）会被合并，
  页眉页脚内部引用的样式与编号也会与正文一致地合并。
- **超出原版 docxtpl 的引擎能力**：模板内 `str` 方法（`upper/replace/split/...`）、
  `'%s' % x` 格式化、`{% break %}`/`{% continue %}`、`{% do %}` 语句、gettext i18n
  （`{% trans %}`/`{% pluralize %}`）、`set_template_loader()` 支持的 `{% include %}`/
  `{% import %}`、带关键字参数的自定义 filters/tests/functions/globals，以及与真实
  jinja2 `Environment` 的互操作。
- **内置可读写文档对象模型**：paragraphs/tables/sections/styles/comments/core properties
  以及 `add_paragraph/add_heading/add_picture/add_table/add_page_break/add_section`，
  覆盖 python-docx 常用读写路径，无需再安装 python-docx——同一模型也通过 `doc`
  模块面向纯 Rust 用户（Python 绑定只是它的薄包装）。完整 `Font`（29 个属性，
  python-docx 三态语义）、`ParagraphFormat`（缩进/间距/行距规则/keep 标志/制表位）、
  表格列/对齐/自动适应、行高、单元格垂直对齐、节起始类型、域代码（`add_field`、
  `update_fields_on_open`）均已支持。
- **内置 CLI**：`python -m docxtplrs template.docx data.json out.docx`，无需写脚本即可
  在命令行直接渲染模板。

### 性能

引擎对引擎（docxtplrs vs docxtpl + lxml + jinja2，相同模板）：小模板每次渲染
**快约 6-12 倍**，8.9MB 大文档**快约 60 倍**（见上方"快"一条）。

在此之上，专项 profiling 又清掉了剩余的内部热点（数据基于 `maturin develop`
debug 构建 + CPython 3.13 实测（除特别注明外），release 构建会更快）：

- **文档对象模型（实时代理）**：段落/表格/节等代理此前**每次**属性读写都重新全量解析
  `word/document.xml`；现在解析结果带脏标记缓存，仅在 render/save 时写回。读取 400 个
  段落文本：**1253ms → 3.6ms（约 350 倍）**；200 次 `add_paragraph`：**832ms → 7ms
  （约 120 倍）**；200 次 run 文本改写：**801ms → 0.4ms（约 2000 倍）**。
- **per-part DOM 缓存**：`styles.xml` / `settings.xml` / `comments.xml` / `core.xml`
  （以及所有 `XmlElement` 代理的 part）此前**每次**访问都重新解析、每次修改都重新
  序列化；现在与 `document.xml` 一样解析一次、render/save 时统一写回。200 个 style 下
  连续 20 次 `styles[name]` 查找：**21.0ms → 0.41ms（约 50 倍）**；20 次 style 字体
  写入：**8.6ms → 0.02ms（约 430 倍）**（release 构建）；`resolve_style_id` 不再每次
  调用都深克隆全部 style 节点。
- **每次 render 只构建一次 env 与 context**：minijinja `Environment`（40+ 个
  filter/test 注册）和 Python context → `Value` 的转换此前**每个 part**（body、每个
  页眉/页脚、footnotes、core 属性）都重建一遍；现在每次 `render()` 调用只构建一次。
  渲染 7 个 part + 300 个变量的文档：**5.0ms → 1.9ms（-61%）**（release 构建）。
- **渲染管线**：所有 XML patch pass 加了 gate（零匹配不再全量拷贝文档），
  `InlineImage`/`Subdoc` 占位符单遍扫描实体化且 blob 用 `Arc` 共享，表格/docPr 修复
  在无事可修时跳过整个 DOM 往返。渲染 2MB / 2 万段落文档：**3384ms → 1901ms（-44%）**。
- **大文档渲染路径**：`patch_xml` 耗时最高的两条正则（拆分标签剥离、标签清理）与
  段落换行处理改为手写线性扫描器，管线各 pass 返回 `Cow`（无修改时零拷贝），
  xmldom 文本节点单次分配、批量转义、序列化预分配；subdoc 合并时 `pic:cNvPr`
  重编号并入同一趟 DOM 往返，不再二次解析+序列化。8.9MB / 2 万段落文档上渲染时间
  再降 **约 60%**（Rust，release 构建）——同样输入 docxtpl 需要约 20s。
- **图片与 subdoc 合并**：图片去重哈希带缓存（插一张图不再重算全部媒体的 sha1），
  `.rels` 每个 part 只解析一次（而非每条关系一次），rId/style/numId 重映射从
  "每条目一次全文正则"改为单遍扫描 + map 查表。
- **zip 与模板缓存轮**：png/jpg 等已压缩媒体在 save 时不再重新 deflate（直接 Stored
  写出），`render()` 重载时克隆原始 Package 而非重新解压整个 zip，编译后的 jinja
  模板按 part 缓存（内容哈希为 key）跨 render 复用，三个必然触发的 `patch_xml`
  pass 融合为单趟扫描，无表格文档的表格/docPr 修复跳过整个 DOM 往返。含 4×1MB
  图片的 4.2MB 模板：save **204ms → 7ms（约 29 倍）**，重复 render
  **2.5ms → 0.9ms（约 3 倍）**，完整周期 **177ms → 11ms（约 16 倍）**（release
  构建；`tests/bench_media.py`）。
- **桥接与 listing 轮**：`resolve_listing`（换行/制表符展开）不再逐段落跑
  fancy-regex（改为一趟线性扫描），zip 级媒体替换不再整体克隆包，媒体 blob
  改为 `Arc` 共享且只哈希一次，Python 桥接层的标量转换走快路径（免 TLS 查询、
  免失败 cast、dict miss 不再产生异常）。listing 密集渲染（300 个多行单元格）：
  **8.9ms → 4.8ms（约 1.9 倍）**（对比 P0 前基线；`tests/bench_listing.py`）。
- **并发与模型轮**：`render()`/`save()` 现在脱离 GIL 执行（Python 回调按需重新
  获取），docmodel 代理改用带代际校验的访问游标（不再每次从 body 头部重扫——
  顺序读取 2000 个段落文本：**17.4ms → 0.7ms（约 25 倍）**），`.rels` 查找
  改为 `Rc` 共享，styleId 解析与 subdoc bookmark 重编号带缓存，渲染管线在
  footnotes/core 属性无模板标签时整体跳过。
- **扫描器与分配轮**：剩余热点扫描器全部 SIMD 化（`memchr` 在候选字节间跳跃，
  替代逐字节循环），行/段落/run 元素剥离改为先定位稀少的 jinja 标记再反查包围
  元素（不再探测每个元素开标签），`resolve_listing` 只重写真正含 listing 字符的
  段落（多 MB 文档在命中点之间整块拷贝）。xmldom 元素名/属性名改为 intern 共享
  （`Rc<str>`，docx 中不同名字仅几十种：parse 分配 **74 万 → 44.5 万，-40%**），
  融合的 jinja 发射器复用每标签缓冲（`patch_xml` 分配 **2.05 万 → 31**）。zip
  条目负载改为 `Arc` 共享——render 重载从全量拷贝变为引用计数自增——已压缩媒体
  全程保留原始 deflate 流（首次访问才解压，save 时原样拷贝回写）；压缩切换到
  更快的 zlib-ng 后端，`get_docx_bytes()`/`get_xml()` 同样脱离 GIL 执行。纯函数
  的 per-part 预处理链（实体解码 + patch_xml + 各重写 pass）按原始源串哈希缓存，
  重复 `render()` 时整链跳过。8.9MB 文档（Rust，release 构建）：patch_xml
  **148ms → 57ms**，表格/docPr 修复 **234ms → 118ms**，save
  **103ms → 约 35ms**，端到端渲染 **673ms → 348ms（-48%）**；4.2MB 媒体模板
  重复 render **0.9ms → 0.4ms**。
- **结构扫描与桥接轮**：表格/docPr 修复先跑一趟只读字符串扫描，精确复刻
  `fix_one_table` 的判定逻辑（含嵌套表与 gridSpan 求和），判断是否有表格真的
  需要修复——渲染后网格几乎总是完好的，此时整个 DOM 解析+遍历直接跳过
  （8.9MB 文档上 **118ms → 16ms**）。元素标记（`{%tr %}`/`{%p %}` 等）由一趟
  aho-corasick 多模式扫描统一收集，替代 11 趟 memmem 门控（patch_xml
  **57ms → 35ms**）；融合的 jinja 发射器改用 memmem 精确定位 `<w:t>`，不再逐个
  检查 `<`；未修改的包条目在 save 时一律原样拷贝（不再仅限媒体）；part 源串在
  UTF-8 快路径零拷贝解码。Python 侧 dict 属性访问改为每个 dict 一次 GIL 获取
  批量物化（不再每个属性一次）：8.9MB 文档重复 render **185ms → 50ms**，完整
  周期 **367ms → 226ms**，小模板对 docxtpl 提速比升至 **约 15-28 倍**。8.9MB
  基准端到端渲染：**348ms → 约 190ms**（Rust，release 构建）。

### 1. Rust 用法

在 `Cargo.toml` 中添加：

```toml
[dependencies]
docxtplrs = { version = "0.2.2", default-features = false }  # 纯 Rust：不依赖 PyO3/libpython
minijinja = "2"
```

Cargo features：

| Feature | 默认 | 作用 |
|---|---|---|
| `python` | ✓ | PyO3 绑定与 Python 互操作（Python 回调作为自定义 filters/tests/functions/globals、`set_template_loader`）。纯 Rust 场景用 `default-features = false` 关闭，完全不依赖 `pyo3`/`libpython`。 |
| `extension-module` | – | 仅供 maturin 构建 Python wheel 时启用（隐含 `python`）；Rust 二进制/库切勿启用。 |

渲染模板的代码见[快速开始](#quick-start--快速开始)（中文上下文把 `title`/`items`
换成中文即可）。

用纯 Rust 文档对象模型编辑文档（无需 Python）：

```rust
use docxtplrs::{doc, template::TplCore};

let mut tpl = TplCore::new(std::fs::read("in.docx")?);
let p = doc::paragraphs(&mut tpl)[0];
p.set_alignment(&mut tpl, Some(1))?;            // 居中
let run = p.add_run(&mut tpl, "重点")?;
run.font().set_bold(&mut tpl, Some(true))?;      // 与 python-docx 同三态语义
doc::add_table(&mut tpl, 2, 2)?;
std::fs::write("out.docx", tpl.save_bytes()?)?;
```

主要 Rust API（均在 `docxtplrs::` 下）：

| 模块 | 要点 |
|---|---|
| `template` | `TplCore`：`new/render/save_bytes/get_xml/undeclared_variables/build_url_id`、`Deferred`、替换映射（`crc_to_new_media`、`pics_to_replace` 等） |
| `package` | `Package`（docx zip 条目、`.rels`、内容类型）、`Rels`、CRC32/SHA1 工具 |
| `patch` | `patch_xml`、`resolve_listing`、`decode_text_entities`（docxtpl 的 XML 预处理） |
| `composer` | `Composer`：docxcompose 风格的整文档拼接（`new`/`append`/`save_bytes`） |
| `doc` | 纯 Rust 文档对象模型：`Paragraph/Run/Table/TableRow/Cell/Section/Style` 索引句柄（`text`/`style`/`alignment` 读写、`Font`/`ParagraphFormat` 句柄、`add_paragraph/add_table` 等自由函数）——Python 绑定即此模型的薄包装 |
| `richtext` | `richtext_run`、`richtext_paragraph`、`listing_xml`、`TextProps` |
| `image` | `ImageInfo`（PNG/JPEG/GIF/BMP/TIFF 尺寸与 DPI）、EMU 换算 |
| `gettext` | `Catalog`（.mo 解析、复数规则求值） |
| `xmldom` | 用于结构化修改的极简 XML DOM |

内置可运行示例：

```bash
cargo run --example render --release -- template.docx out.docx
```

> 注意：默认 features 下示例会链接 `libpython`（PyO3）。若 Python 由 uv 管理，请用
> `LD_LIBRARY_PATH=$(python -c "import sysconfig; print(sysconfig.get_config_var('LIBDIR'))")` 运行。
> 使用 `--no-default-features` 构建时（如 `cargo run --no-default-features --example render --release -- ...`）
> 示例为纯 Rust，完全不需要 Python。

### 2. Python 用法

#### 安装

```bash
# 本仓库内：构建并安装到 uv venv
uv sync

# 在其他 uv 项目中
uv add /path/to/docxtplrs/target/wheels/docxtplrs-0.2.2-cp313-cp313-manylinux_2_17_x86_64.manylinux2014_x86_64.whl
# 或本地开发模式（editable）
uv add --editable /path/to/docxtplrs
```

wheel 为原生 `cp313` 构建。其他 Python 版本需在对应版本下构建，或把 `Cargo.toml`
中 pyo3 改为 `abi3-py312` 以产出单个通用 wheel。

#### 自行构建 wheel

```bash
# release wheel -> target/wheels/docxtplrs-<ver>-cp313-cp313-manylinux_2_34_x86_64.whl
uv run maturin build --release

# 本机同时装有 conda 时：
env -u CONDA_PREFIX uv run maturin build --release
```

其他变体：

```bash
# abi3 通用 wheel（一份通吃 CPython >= 3.12）：
#   1. 把 Cargo.toml 的 pyo3 改为 features = ["abi3-py312"]
#   2. 构建
uv run maturin build --release        # -> docxtplrs-...-cp312-abi3-....whl

# 指定 Python 版本构建：把 maturin 指向对应解释器
uv run maturin build --release -i python3.12

# 同时产出 sdist（源码分发包）
uv run maturin build --release --sdist

# 安装 / 分发
uv pip install target/wheels/*.whl                     # 本地安装
# 或上传到自有索引，例如：uv publish target/wheels/*   （或 twine upload）
```

#### 模板语法（与 docxtpl 一致）

| 语法 | 用途 |
|---|---|
| `{{ var }}` | 变量（过滤器、属性/下标访问、方法调用、str 方法、`'%s' % x`、Unicode 标识符） |
| `{% if %}` `{% for %}` `{% set %}` `{% macro %}` | 语句（支持 break/continue/do） |
| `{%tr for x in items %}` | 表格行循环 |
| `{%tc for x in items %}` | 单元格循环（自动修复 `tblGrid` 列数/宽度） |
| `{%p if x %}` / `{{p var }}` | 段落级语句 / 替换 |
| `{{r var }}` | run 级替换（配合 `RichText`） |
| `{% colspan x %}` `{% cellbg c %}` `{% vm %}` `{% hm %}` | 跨列 / 单元格底色 / 垂直与水平合并 |
| `{%- ... -%}` / `{_{ ... }_}` | 段落文本合并 / 字面花括号转义 |
| `{% trans %}` / `{% pluralize %}` | i18n（先 `tpl.install_gettext("zh.mo")`） |
| `{% include %}` / `{% import %}` | 包含（先 `tpl.set_template_loader(fn)`） |
| 页眉 / 页脚 / 脚注 / 文档属性 | 同样会被渲染（图片可插入页眉） |

完整渲染示例：

```python
from docxtplrs import DocxTemplate, R, RichText, Listing, InlineImage, Mm

tpl = DocxTemplate("template.docx")          # 路径 / 文件对象 / bytes

tpl.render({
    "title": "季度报告",
    "rows": ["华东", "华南"],                 # {%p for r in rows %} 段落循环
    "items": [{"name": "甲", "price": 12.5}],  # {%tr for x in items %} 表格行循环
    "rt": R("重点", bold=True, color="#C00000"),  # 模板里写 {{r rt }}
    "img": InlineImage(tpl, "logo.png", width=Mm(15)),  # {{ img }}
    "sub": tpl.new_subdoc("chapter.docx"),                # {{p sub }}；keep_sections=True
                                                            # 可保留子文档页面设置
}, autoescape=True)

tpl.save("out.docx")                         # 或 io.BytesIO()
```

#### Python API 概览

- `DocxTemplate(f)`：`render()`、`save()`、`get_docx_bytes()`、`get_xml()`、
  `get_undeclared_template_variables()`、`new_subdoc()`、`build_url_id()`、
  `replace_media()` / `replace_embedded()` / `replace_pic()` / `replace_zipname()`、
  `reset_replacements()`、`allow_missing_pics`、`get_pic_map()`
- 自定义：`register_filter/test/function/global()`、`set_template_loader()`、
  `install_gettext()`；`render(ctx, jinja_env=...)` 接受真实 jinja2
  `Environment`（读取 `autoescape/filters/globals/tests/trim_blocks/lstrip_blocks/
  keep_trailing_newline/undefined`）
- 文档对象模型（可读写）：`get_docx()`、`paragraphs/tables/sections/
  styles/settings/comments/core_properties/inline_shapes`、
  `add_paragraph/add_heading/add_picture/add_table/add_page_break/add_section`、
  完整 `Font`（29 个三态属性）与 `ParagraphFormat`（缩进/间距/行距规则/
  keep 标志/制表位，段落与样式均可）、表格 `columns/add_column/alignment/autofit`、
  行 `height/height_rule`、单元格 `paragraphs/width/vertical_alignment`、
  节 `start_type/页眉页脚距离`、样式标志位（`hidden/locked/priority/quick_style/builtin/...`）、
  `insert_paragraph_before`/`clear`/`add_break`/`add_tab`/`add_text`、
  `hyperlinks`、`iter_inner_content`（文档/节/段落/run/单元格）、
  `contains_page_break`/`rendered_page_breaks`、`part` 门面、
  域代码（`paragraph.fields`/`add_field`、`instr`/缓存结果读写、
  `settings.update_fields_on_open`）、
  `run.add_picture`/`mark_comment_range`、`table_direction`、
  `grid_span`/`grid_cols_before/after`、单元格内嵌套表格、
  `__getattr__` 委托
- 原始 XML 逃生口：`settings/document/styles.element` 返回实时的 `XmlElement`
  代理（`tag/text/attrib/get/set/remove_attr/find/findall/children/append/
  insert/remove/xml`），覆盖 python-docx `.element` 的常见用法（如用
  `settings.element` 设置 `w:zoom`）
- `Table.style`、`Cell.merge()`（gridSpan/vMerge，python-docx 语义）、
  `Section.even_page_header/footer`、`Section.first_page_header/footer`
- docxtpl 辅助 API：`HEADER_URI`/`FOOTER_URI`、`get_file_crc()`（路径/bytes/
  文件对象）、`get_headers_footers(uri)` 产出实时的 `(rid, XmlElement)` 对
- `RichText`/`R`、`RichTextParagraph`/`RP`、`Listing`、`InlineImage`、`Subdoc`
  （文件合并 + 编程式 `add_paragraph/add_picture/add_table`；
  `new_subdoc(path, keep_sections=True)` 还会把子文档的节属性——页面设置与
  页眉页脚引用——保留为独立一节；子文档的**批注**在任何模式下都会被合并）
- `Composer`：docxcompose 风格的整文档拼接——`Composer(master).append(doc).save(out)`，
  文档间自动分页，样式/编号/媒体自动合并：

  ```python
  from docxtplrs import Composer

  c = Composer("master.docx")   # 路径 / 文件对象 / bytes
  c.append("chapter1.docx")     # 同样支持这三种输入
  c.append("chapter2.docx")
  c.save("out.docx")            # 或 io.BytesIO()
  ```
- 单位：`Length/Emu/Inches/Cm/Mm/Pt/Twips`；jinja2.utils：`Cycler/Joiner/
  generate_lorem_ipsum`；异常：`TemplateError`（含 `docx_context`）
- CLI：`python -m docxtplrs template.docx data.json output.docx [-o] [-q]`

#### 自定义过滤器与引擎扩展

注册自己的 jinja 过滤器、测试器、函数与全局值——普通 Python 可调用对象，
位置参数**和关键字参数**均支持：

```python
tpl = DocxTemplate("template.docx")

# 过滤器：{{ price|rmb }}、{{ name|shout('!') }}、{{ rows|col(name='药物名称') }}
tpl.register_filter("rmb", lambda v: f"¥{v:.2f}")
tpl.register_filter("shout", lambda v, punct="?": str(v).upper() + punct)
tpl.register_filter("col", lambda rows, name: [r[name] for r in rows])

# 测试器：{% if v is even %}
tpl.register_test("even", lambda v: v % 2 == 0)

# 函数与全局值：{{ add(1, 2) }}、{{ company }}
tpl.register_function("add", lambda a, b: a + b)
tpl.register_global("company", "ACME")

# 模板 loader（{% include %}/{% import %}）
tpl.set_template_loader(lambda name: "included {{ v }}" if name == "part" else None)

# gettext i18n（{% trans %}/{% pluralize %}）
tpl.install_gettext("messages.mo")

tpl.render(context)
```

也可以直接传真正的 **jinja2 Environment**——其 `autoescape`、`filters`、
`globals`、`tests`、`trim_blocks`、`lstrip_blocks`、`keep_trailing_newline` 与
`undefined`（Chainable/Strict）设置均会被读取（鸭子类型）：

```python
import jinja2
env = jinja2.Environment()
env.filters["usd"] = lambda v: f"${v:.2f}"
tpl.render(context, jinja_env=env)
```

注意：jinja2 自带的内建 filters/tests/globals **不会**被导入——它们通过与
`jinja2.defaults` 的同一性判断识别，并交由 minijinja 原生实现处理（因此
`default`/`defined` 等 undefined 语义保持正确）；只有你新增或覆盖的条目会被导入，
且同名时优先于 `register_filter/test/function/global()` 的注册。

引擎层相对原生 minijinja 的补齐，jinja2 风格模板开箱即用：

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

#### 测试

```bash
.venv/bin/python -m pytest tests/ -q          # 445 个单元测试
.venv/bin/python tests/crosscheck.py docxtplrs > /tmp/rs.json
.venv/bin/python tests/crosscheck.py docxtpl > /tmp/ref.json   # 需要 crosscheck 依赖组
.venv/bin/python tests/crosscheck.py compare /tmp/ref.json /tmp/rs.json
.venv/bin/python tests/benchmark.py --big     # 引擎对引擎基准
```

覆盖率徽章：`src/` 的行覆盖率，测量方式为 `RUSTFLAGS="-Cinstrument-coverage"` 重新
构建后跑 pytest 套件，再用 `llvm-cov` 出报告。

### 3. 限制

**架构层面**

- 引擎为 minijinja 而非 jinja2：总体兼容，但不保证极端 jinja2 边角行为一致
  （常见差异已通过自定义测试/过滤器、loader、Unicode 标识符、kwargs 支持等补齐）。
- 文档对象模型覆盖 python-docx 常用读写路径（含完整 `Font`/`ParagraphFormat` 对象、
  表格/行/单元格几何属性、样式标志位、`rendered_page_breaks` 标记、`iter_inner_content`
  与最小 `part` 门面、域代码读写），但**不是完整重写**：OLE 编辑未覆盖。
- `get_docx()` 返回本库的外观对象，**不是** python-docx 的 `Document`——不能传给
  期望 python-docx 类型的第三方代码。

**行为差异（结果相当或更好）**

- 非法 XML 容错：非命名空间的 `<` 会被转义为文本（产出 schema 合法的文档），而 lxml
  recover 会把文本吞进未知元素；`a<b` 这类文本两边都按标签处理。
- `== true` / `== none` 比较在编译前改写为等义的自定义测试（精确的 Python 相等语义，
  仅限标签内，字符串字面量不受影响）。
- `{% trans %}` 是完整的 gettext 实现（.mo 解析 + 复数规则），但不覆盖 jinja2 newstyle
  i18n 的所有角落。
- MB 级大文档上，docxtpl 的 8 处正则被替换为线性扫描器以避免回溯栈溢出；输出已与官方
  测试套件逐字节比对，但极端畸形的模板可能有差异。
- zip 输出保留压缩方式与时间戳，但不保证与 python-docx 输出逐字节一致。

**不支持**

- 通过 `jinja_env` 自定义 jinja 界定符（docxtpl 本身也不支持）。
- i18n / do / loopcontrols / debug 之外的 jinja2 扩展。
- WMF 图片（上游 python-docx 1.2 同样已放弃）。

---

## Project structure / 项目结构

```
src/            Rust sources (17 modules, see AGENTS.md)   Rust 源码
python/         Python package shell (__init__ + __main__ CLI)
tests/          445 tests + crosscheck/compare/benchmark    测试、交叉验证与基准
examples/       render.rs (Rust API), patch_dbg.rs (large-doc debugging), bench.rs (stage timings)
AGENTS.md       Notes for AI coding assistants              给 AI 助手的项目说明
```

## License / 许可证

Same as the reference implementation python-docx-template: LGPL-2.1-or-later
(please verify compliance yourself before redistribution).

与参考实现 python-docx-template 相同：LGPL-2.1-or-later（再分发前请自行确认合规）。
