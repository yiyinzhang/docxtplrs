"""Basic tests for docxtplrs."""

import os
import tempfile
import zipfile
from pathlib import Path

import pytest

try:
    from docxtplrs import (
        DocxTemplate,
        RichText,
        InlineImage,
        Subdoc,
        Mm,
        Inches,
        Pt,
        R,
        Listing,
        escape_xml,
        unescape_xml,
    )
except ImportError:
    pytest.skip("docxtplrs not built", allow_module_level=True)


def create_simple_template() -> str:
    """Create a simple .docx template for testing."""
    # Create a minimal .docx file
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        template_path = f.name

    with zipfile.ZipFile(template_path, "w") as zf:
        # [Content_Types].xml
        zf.writestr(
            "[Content_Types].xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''',
        )

        # _rels/.rels
        zf.writestr(
            "_rels/.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''',
        )

        # word/_rels/document.xml.rels
        zf.writestr(
            "word/_rels/document.xml.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''',
        )

        # word/document.xml
        zf.writestr(
            "word/document.xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Hello {{ name }}!</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Welcome to {{ company }}.</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>''',
        )

    return template_path


def create_conditional_template() -> str:
    """Create a template with conditional content."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        template_path = f.name

    with zipfile.ZipFile(template_path, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''',
        )

        zf.writestr(
            "_rels/.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''',
        )

        zf.writestr(
            "word/_rels/document.xml.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''',
        )

        zf.writestr(
            "word/document.xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{% if show_message %}Hello World!{% endif %}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>{% for item in items %}{{ item }} {% endfor %}</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>''',
        )

    return template_path


class TestDocxTemplate:
    """Test DocxTemplate class."""

    def test_load_template(self):
        """Test loading a template."""
        template_path = create_simple_template()
        try:
            doc = DocxTemplate(template_path)
            assert doc is not None
            xml = doc.get_xml()
            assert "Hello" in xml
        finally:
            os.unlink(template_path)

    def test_render_simple(self):
        """Test basic rendering."""
        template_path = create_simple_template()
        output_path = template_path.replace(".docx", "_output.docx")

        try:
            doc = DocxTemplate(template_path)
            context = {"name": "World", "company": "Example Corp"}
            doc.render(context)
            doc.save(output_path)

            # Verify output file exists
            assert os.path.exists(output_path)

            # Check content
            with zipfile.ZipFile(output_path, "r") as zf:
                content = zf.read("word/document.xml").decode("utf-8")
                assert "Hello World!" in content
                assert "Welcome to Example Corp." in content
        finally:
            if os.path.exists(template_path):
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_render_conditional(self):
        """Test conditional rendering."""
        template_path = create_conditional_template()
        output_path = template_path.replace(".docx", "_output.docx")

        try:
            doc = DocxTemplate(template_path)
            context = {"show_message": True, "items": ["a", "b", "c"]}
            doc.render(context)
            doc.save(output_path)

            assert os.path.exists(output_path)

            with zipfile.ZipFile(output_path, "r") as zf:
                content = zf.read("word/document.xml").decode("utf-8")
                assert "Hello World!" in content
                assert "a b c" in content
        finally:
            if os.path.exists(template_path):
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_get_undeclared_variables(self):
        """Test getting undeclared variables."""
        template_path = create_simple_template()

        try:
            doc = DocxTemplate(template_path)
            vars_needed = doc.get_undeclared_template_variables()

            assert "name" in vars_needed
            assert "company" in vars_needed

            context = {"name": "Test"}
            vars_missing = doc.get_undeclared_template_variables(context)
            assert "name" not in vars_missing
            assert "company" in vars_missing
        finally:
            os.unlink(template_path)


class TestRichText:
    """Test RichText class."""

    def test_create_richtext(self):
        """Test creating RichText."""
        rt = RichText("Hello", bold=True)
        assert rt is not None
        assert str(rt) == "Hello"

    def test_add_text(self):
        """Test adding text to RichText."""
        rt = RichText()
        rt.add("Hello ")
        rt.add("World", bold=True, color="FF0000")
        assert str(rt) == "Hello World"

    def test_add_newline(self):
        """Test adding newline."""
        rt = RichText("Line 1")
        rt.add_newline()
        rt.add("Line 2")
        assert "Line 1" in str(rt)
        assert "Line 2" in str(rt)

    def test_r_shortcut(self):
        """Test R() shortcut."""
        rt = R("Test", bold=True, italic=True)
        assert rt is not None
        assert str(rt) == "Test"


class TestMeasurements:
    """Test measurement classes."""

    def test_mm(self):
        """Test millimeters."""
        mm = Mm(50)
        assert float(mm) == 50.0
        assert "50" in repr(mm)

    def test_inches(self):
        """Test inches."""
        inches = Inches(2)
        assert float(inches) == 2.0
        assert "2" in repr(inches)

    def test_pt(self):
        """Test points."""
        pt = Pt(12)
        assert float(pt) == 12.0
        assert "12" in repr(pt)


class TestUtilityFunctions:
    """Test utility functions."""

    def test_escape_xml(self):
        """Test XML escaping."""
        assert escape_xml("<test>") == "&lt;test&gt;"
        assert escape_xml("&") == "&amp;"
        assert escape_xml('"test"') == "&quot;test&quot;"

    def test_unescape_xml(self):
        """Test XML unescaping."""
        assert unescape_xml("&lt;test&gt;") == "<test>"
        assert unescape_xml("&amp;") == "&"
        assert unescape_xml("&quot;test&quot;") == '"test"'


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
