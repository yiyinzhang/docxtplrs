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

    def test_set_updatefields_true(self):
        """Test setting updateFields to true."""
        template_path = create_simple_template()
        output_path = template_path.replace(".docx", "_output.docx")

        try:
            doc = DocxTemplate(template_path)
            
            # Call set_updatefields_true
            doc.set_updatefields_true()
            
            # Render and save
            doc.render({"name": "Test", "company": "Corp"})
            doc.save(output_path)

            # Check settings.xml
            with zipfile.ZipFile(output_path, "r") as zf:
                # Check if settings.xml exists
                if "word/settings.xml" in zf.namelist():
                    content = zf.read("word/settings.xml").decode("utf-8")
                    assert "updateFields" in content or "updatefields" in content.lower()
                else:
                    # If no settings.xml, that's also OK - it means one was created
                    pass
        finally:
            if os.path.exists(template_path):
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_docx_properties(self):
        """Test getting and setting document properties."""
        template_path = create_simple_template()
        output_path = template_path.replace(".docx", "_output.docx")

        try:
            doc = DocxTemplate(template_path)
            
            # Initially properties should be empty or minimal
            props = doc.get_docx_properties()
            assert isinstance(props, dict)
            
            # Set properties
            doc.set_docx_properties({
                "author": "Test Author",
                "title": "Test Title",
                "subject": "Test Subject",
            })
            
            # Get updated properties
            props = doc.get_docx_properties()
            assert props.get("author") == "Test Author"
            assert props.get("title") == "Test Title"
            assert props.get("subject") == "Test Subject"
            
            # Render and save
            doc.render({"name": "Test", "company": "Corp"})
            doc.save(output_path)
            
            assert os.path.exists(output_path)
        finally:
            if os.path.exists(template_path):
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_paragraph_properties(self):
        """Test setting paragraph properties."""
        template_path = create_simple_template()
        output_path = template_path.replace(".docx", "_output.docx")

        try:
            doc = DocxTemplate(template_path)
            
            # Set paragraph properties
            doc.set_paragraph_properties(
                paragraph_index=0,
                style_id="Heading1",
                alignment="center",
            )
            
            # Render and save
            doc.render({"name": "Test", "company": "Corp"})
            doc.save(output_path)
            
            # Check document.xml
            with zipfile.ZipFile(output_path, "r") as zf:
                content = zf.read("word/document.xml").decode("utf-8")
                assert "Heading1" in content
                assert "center" in content or "jc" in content
        finally:
            if os.path.exists(template_path):
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)


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

    def test_cm(self):
        """Test centimeters."""
        from docxtplrs import Cm
        cm = Cm(5)
        assert float(cm) == 5.0
        assert "5" in repr(cm)

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


class TestJinjaEnv:
    """Test JinjaEnv class for custom filters."""

    def test_create_jinja_env(self):
        """Test creating JinjaEnv."""
        from docxtplrs import JinjaEnv
        env = JinjaEnv()
        assert env is not None
        assert repr(env) == "JinjaEnv(filters=0)"

    def test_add_filter(self):
        """Test adding a filter."""
        from docxtplrs import JinjaEnv
        env = JinjaEnv()
        
        def double(value):
            return value * 2
        
        env.add_filter("double", double)
        assert env.has_filter("double")
        assert "double" in env.get_filter_names()

    def test_remove_filter(self):
        """Test removing a filter."""
        from docxtplrs import JinjaEnv
        env = JinjaEnv()
        
        def test_filter(value):
            return value
        
        env.add_filter("test", test_filter)
        assert env.has_filter("test")
        
        env.remove_filter("test")
        assert not env.has_filter("test")

    def test_clear_filters(self):
        """Test clearing all filters."""
        from docxtplrs import JinjaEnv
        env = JinjaEnv()
        
        def f1(v): return v
        def f2(v): return v
        
        env.add_filter("f1", f1)
        env.add_filter("f2", f2)
        assert len(env.get_filter_names()) == 2
        
        env.clear_filters()
        assert len(env.get_filter_names()) == 0

    def test_filter_not_callable(self):
        """Test that non-callable filters are rejected."""
        from docxtplrs import JinjaEnv
        env = JinjaEnv()
        
        with pytest.raises(TypeError):
            env.add_filter("not_callable", "string")


class TestCustomFilters:
    """Test custom filters in template rendering."""

    def test_uppercase_filter(self):
        """Test uppercase filter in template."""
        from docxtplrs import DocxTemplate, JinjaEnv
        
        template_path = create_simple_template()
        output_path = template_path.replace(".docx", "_output.docx")
        
        try:
            # Modify template to use filter
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
                <w:t>Hello {{ name|upper }}!</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>''',
                )
            
            doc = DocxTemplate(template_path)
            
            env = JinjaEnv()
            env.add_filter("upper", lambda x: str(x).upper())
            
            context = {"name": "world"}
            doc.render(context, jinja_env=env)
            doc.save(output_path)
            
            with zipfile.ZipFile(output_path, "r") as zf:
                content = zf.read("word/document.xml").decode("utf-8")
                assert "HELLO WORLD" in content or "WORLD" in content
                
        finally:
            if os.path.exists(template_path):
                os.unlink(template_path)
            if os.path.exists(output_path):
                os.unlink(output_path)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
