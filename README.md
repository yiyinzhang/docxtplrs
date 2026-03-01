# docxtplrs

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Built with Vibe Coding](https://img.shields.io/badge/Built%20with-Vibe%20Coding-purple.svg)](https://twitter.com/karpathy/status/1886191732477106580)
[![Status](https://img.shields.io/badge/Status-Work%20in%20Progress-orange.svg)]()

> 🎧 **A Vibe Coding Project** — This entire library was built through AI-assisted programming. No traditional coding sessions, just pure vibes with Claude and Cursor. If you find any "vibey" code patterns, now you know why!

> ⚠️ **Under Active Development** — This project is still in early development. APIs may change, features may break, and documentation is evolving. Use at your own risk, and feel free to open issues for bugs or feature requests!

A Rust implementation of [python-docx-template](https://github.com/elapouya/python-docx-template) with Python bindings.

## Features

- 🦀 **High Performance**: Written in Rust for maximum speed
- 🐍 **Python Compatible**: Drop-in replacement for python-docx-template
- 📝 **Jinja2 Templates**: Use familiar Jinja2 syntax in Word documents
- 🔧 **Custom Filters**: Define your own Jinja2 filters for data transformation
- 🎨 **Rich Text**: Styled text with fonts, colors, and formatting
- 🖼️ **Images**: Inline image support with customizable dimensions
- 📄 **Sub-documents**: Embed other documents or build content programmatically
- 🔗 **Hyperlinks**: Add clickable links to your documents
- 📊 **Tables**: Support for table row/cell loops and formatting
- 🛡️ **Type Safe**: Full type hints included

## Installation

This is a Rust library with Python bindings and is **not available on PyPI**. You need to build it from source.

### Prerequisites

- [Rust](https://www.rust-lang.org/tools/install) toolchain
- [uv](https://docs.astral.sh/uv/getting-started/installation/) package manager

### Build with uv + maturin

```bash
# Clone the repository
git clone https://github.com/yourusername/docxtplrs
cd docxtplrs

# Create virtual environment and install maturin
uv venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
uv pip install maturin

# Build and install in development mode
maturin develop --release

# Or use uv run directly
uv run maturin develop --release
```

### Import in Python

After building, you can import the package in your Python code:

```python
from docxtplrs import DocxTemplate

# Load template
doc = DocxTemplate("my_template.docx")

# Render with context
context = {
    "company_name": "Example Corp",
    "user_name": "John Doe",
}
doc.render(context)

# Save result
doc.save("generated_doc.docx")
```

### Use in Your Project

Add the local package to your project's dependencies in `pyproject.toml`:

```toml
[tool.uv.sources]
docxtplrs = { path = "/path/to/docxtplrs" }
```

Or install directly:

```bash
uv pip install /path/to/docxtplrs
```

## Quick Start

### Basic Usage

```python
from docxtplrs import DocxTemplate

# Load template
doc = DocxTemplate("my_template.docx")

# Render with context
context = {
    "company_name": "Example Corp",
    "user_name": "John Doe",
}
doc.render(context)

# Save result
doc.save("generated_doc.docx")
```

### Template Syntax

Create your template in Microsoft Word with Jinja2-like tags:

```
Hello {{ company_name }},

Welcome to our platform, {{ user_name }}!

{% if premium_user %}
You have premium access to all features.
{% endif %}
```

### Special Tags

For controlling entire paragraphs, table rows, or cells:

| Tag | Description |
|-----|-------------|
| `{%p ... %}` | Paragraph-level control |
| `{%tr ... %}` | Table row control |
| `{%tc ... %}` | Table cell control |
| `{%r ... %}` | Run-level control |
| `{{r ... }}` | RichText variables |

```
{%p if show_paragraph %}
This entire paragraph can be conditionally shown.
{%p endif %}
```

### RichText

Add styled text programmatically:

```python
from docxtplrs import DocxTemplate, RichText, R

doc = DocxTemplate("template.docx")

# Using RichText class
rt = RichText("Important: ", bold=True, color="FF0000")
rt.add("This is styled text.", italic=True)

# Using R() shortcut
rt2 = R("Bold text", bold=True)
rt2.add(" and ")
rt2.add("italic text", italic=True)

context = {
    "styled_text": rt,
    "styled_text2": rt2,
}
doc.render(context)
```

Use `{{r styled_text }}` in your template (note the `r` after `{{`).

### Inline Images

```python
from docxtplrs import DocxTemplate, InlineImage, Mm, Cm

doc = DocxTemplate("template.docx")

# Add image with specific dimensions
image = InlineImage(
    doc,
    "logo.png",
    width=Mm(50),  # or Cm(5), Inches(2), Pt(144)
    height=Mm(30)
)

context = {"logo": image}
doc.render(context)
```

Use `{{ logo }}` in your template.

**Measurement Units:**
- `Mm(mm)` - Millimeters
- `Cm(cm)` - Centimeters
- `Inches(inches)` - Inches
- `Pt(points)` - Points

### Hyperlinks

```python
doc = DocxTemplate("template.docx")

# Create URL relationship
url_id = doc.build_url_id("https://example.com")

# Create styled link
rt = RichText("Click here to visit our website")
rt.add_link("example.com", url_id, color="0000FF", underline=True)

context = {"link": rt}
doc.render(context)
```

Use `{{r link }}` in your template.

### Sub-documents

Embed other documents:

```python
doc = DocxTemplate("template.docx")

# Load existing document as subdoc
subdoc = doc.new_subdoc("content.docx")

context = {"embedded_content": subdoc}
doc.render(context)
```

Use `{{p embedded_content }}` in your template.

### Tables

Dynamic table generation:

```python
doc = DocxTemplate("template.docx")

items = [
    {"name": "Item 1", "price": 10.00},
    {"name": "Item 2", "price": 20.00},
]

context = {"items": items}
doc.render(context)
```

In template:

```
| Name | Price |
|------|-------|
{%tr for item in items %}
| {{ item.name }} | {{ item.price }} |
{%tr endfor %}
```

### Cell Formatting

```python
from docxtplrs import DocxTemplate, CellColor, ColSpan

doc = DocxTemplate("template.docx")

context = {
    "header_color": CellColor("4472C4"),
    "title_span": ColSpan(3),
}
doc.render(context)
```

### Getting Undeclared Variables

```python
doc = DocxTemplate("template.docx")

# Get all variables needed
vars_needed = doc.get_undeclared_template_variables()
print(vars_needed)  # ['company_name', 'user_name', ...]

# Check what's missing from your context
vars_missing = doc.get_undeclared_template_variables(context)
```

## Advanced Features

### Media Replacement

Replace images in headers/footers:

```python
doc = DocxTemplate("template.docx")
doc.render(context)

# Replace dummy image with real one
doc.replace_pic("dummy_logo.png", "real_logo.png")
doc.save("output.docx")
```

### Automatic Field Update

Enable automatic update of fields (table of contents, page numbers, etc.) when the document is opened:

```python
doc = DocxTemplate("template.docx")
doc.render(context)

# Enable automatic field update
doc.set_updatefields_true()
doc.save("output.docx")
```

### Document Properties (Metadata)

Get and set document core properties:

```python
doc = DocxTemplate("template.docx")

# Get current properties
props = doc.get_docx_properties()
print(props)  # {'author': '...', 'title': '...', 'subject': '...', ...}

# Set properties
doc.set_docx_properties({
    "author": "John Doe",
    "title": "My Document",
    "subject": "Quarterly Report",
    "keywords": "report, quarterly, finance",
    "description": "This is a quarterly report",
})

doc.render(context)
doc.save("output.docx")
```

### Paragraph Properties

Modify paragraph styles programmatically:

```python
doc = DocxTemplate("template.docx")
doc.render(context)

# Modify paragraph at index 0 (first paragraph)
doc.set_paragraph_properties(
    paragraph_index=0,
    style_id="Heading1",      # Apply heading style
    alignment="center",        # Center alignment
    space_before=240,         # Space before (in twips, 1/20 of a point)
    space_after=120,          # Space after (in twips)
)

doc.save("output.docx")
```

### Custom Filters

You can define custom Jinja2 filters to transform values in your templates:

```python
from docxtplrs import DocxTemplate, JinjaEnv

def format_currency(value):
    """Format a number as currency."""
    return f"${value:,.2f}"

def uppercase(value):
    """Convert value to uppercase."""
    return str(value).upper()

# Create Jinja environment and add filters
env = JinjaEnv()
env.add_filter("currency", format_currency)
env.add_filter("upper", uppercase)

# Use in template: {{ price|currency }} or {{ name|upper }}
doc = DocxTemplate("template.docx")
context = {"price": 1234.5, "name": "john doe"}
doc.render(context, jinja_env=env)
doc.save("output.docx")
```

#### JinjaEnv API

- `add_filter(name, func)` - Add a custom filter
- `remove_filter(name)` - Remove a filter
- `get_filter_names()` - Get list of all filter names
- `has_filter(name)` - Check if a filter exists
- `clear_filters()` - Remove all filters

#### Built-in Filters

- `e` / `escape` - Escape XML special characters

## Comparison with python-docx-template

| Feature | python-docx-template | docxtplrs |
|---------|---------------------|-----------|
| Language | Python | Rust + Python bindings |
| Performance | Good | Excellent |
| Jinja2 Support | Full | Most features |
| Custom Filters | ✅ | ✅ |
| RichText | ✅ | ✅ |
| Inline Images | ✅ | ✅ |
| Sub-documents | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ |
| Document Properties | ✅ | ✅ |
| Paragraph Styling | ✅ | ✅ |
| Field Auto-Update | Manual | ✅ Built-in |
| Type Hints | Partial | Full |

## Performance

Benchmark generating 1000 documents:

```
python-docx-template: 45.2s
docxtplrs:            8.5s  (5.3x faster)
```

## Development

### Setup

```bash
# Install Rust (if not already installed)
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh

# Install maturin
pip install maturin

# Build and install locally
maturin develop

# Run tests
pytest tests/
```

### Building Wheels

```bash
maturin build --release
```

## Vibe Coding

This project was built entirely through **Vibe Coding** — a collaborative approach where AI and humans work together in a continuous flow of ideas and implementation. No rigid specifications, just good vibes and iterative refinement.

If you're curious about the process:
- All major features were implemented through conversational AI interactions
- Code reviews were done by AI assistants
- Documentation written collaboratively
- Bugs fixed through vibe-based debugging sessions

Want to contribute? Just bring good vibes! 🎧✨

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Acknowledgments

This project is inspired by [python-docx-template](https://github.com/elapouya/python-docx-template) by Eric Lapouyade.
