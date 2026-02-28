#!/usr/bin/env python3
"""
Basic example of using docxtplrs to generate Word documents from templates.

This example demonstrates:
- Loading a template
- Rendering with context variables
- Using RichText for styled content
- Saving the result
"""

import os
import tempfile
import zipfile
from pathlib import Path


def create_sample_template(template_path: str):
    """Create a sample .docx template for demonstration."""

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

        # word/document.xml - Template content
        zf.writestr(
            "word/document.xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Title"/>
            </w:pPr>
            <w:r>
                <w:t>Welcome to {{ company_name }}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Dear {{ customer_name }},</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Thank you for choosing our services. We are excited to work with you!</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Your account details:</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>  - Plan: {{ plan_type }}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>  - Start Date: {{ start_date }}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>{% if premium_features %}You have access to premium features!{% endif %}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Best regards,</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>The {{ company_name }} Team</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>''',
        )


def main():
    """Run the example."""
    print("=" * 60)
    print("docxtplrs Basic Example")
    print("=" * 60)

    try:
        from docxtplrs import DocxTemplate, RichText, R, escape_xml
    except ImportError as e:
        print(f"Error: Could not import docxtplrs. Please build the package first.")
        print(f"Run: maturin develop --uv")
        print(f"Details: {e}")
        return 1

    # Create temporary files
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        template_path = f.name

    output_path = template_path.replace(".docx", "_output.docx")

    try:
        # Create sample template
        print("\n1. Creating sample template...")
        create_sample_template(template_path)
        print(f"   Template created: {template_path}")

        # Load template
        print("\n2. Loading template...")
        doc = DocxTemplate(template_path)
        print("   Template loaded successfully")

        # Show template variables
        print("\n3. Checking required variables...")
        vars_needed = doc.get_undeclared_template_variables()
        print(f"   Required variables: {vars_needed}")

        # Prepare context
        print("\n4. Preparing context...")

        # Create styled text
        company_header = RichText("Premium Services Inc.", bold=True, color="0066CC", font_size=24)

        context = {
            "company_name": company_header,
            "customer_name": "John Smith",
            "plan_type": "Professional",
            "start_date": "2024-01-15",
            "premium_features": True,
        }
        print(f"   Context prepared with {len(context)} variables")

        # Render
        print("\n5. Rendering document...")
        doc.render(context)
        print("   Document rendered successfully")

        # Save
        print(f"\n6. Saving to: {output_path}")
        doc.save(output_path)
        print("   Document saved successfully")

        # Verify output
        print("\n7. Verifying output...")
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"   Output file size: {file_size} bytes")

            # Read and show some content
            with zipfile.ZipFile(output_path, "r") as zf:
                content = zf.read("word/document.xml").decode("utf-8")
                if "John Smith" in content:
                    print("   ✓ Customer name rendered correctly")
                if "Professional" in content:
                    print("   ✓ Plan type rendered correctly")
                if "You have access to premium features" in content:
                    print("   ✓ Premium features message shown")

        print("\n" + "=" * 60)
        print("Example completed successfully!")
        print(f"Output file: {output_path}")
        print("=" * 60)

        return 0

    except Exception as e:
        print(f"\nError: {e}")
        import traceback

        traceback.print_exc()
        return 1

    finally:
        # Cleanup
        if os.path.exists(template_path):
            os.unlink(template_path)
        # Keep output file for user to inspect
        if os.path.exists(output_path):
            print(f"\nNote: Output file preserved at: {output_path}")


if __name__ == "__main__":
    exit(main())
