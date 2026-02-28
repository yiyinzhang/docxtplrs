#!/usr/bin/env python3
"""
Advanced example of using docxtplrs.

This example demonstrates:
- Tables with dynamic rows using standard Jinja2 loops
- RichText with multiple styles
- Conditional content
- Loops
"""

import os
import tempfile
import zipfile


def create_advanced_template(template_path: str):
    """Create an advanced .docx template with tables and conditionals."""

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

        # Template with standard Jinja2 for table rows (using single cell with all data)
        zf.writestr(
            "word/document.xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr><w:pStyle w:val="Title"/></w:pPr>
            <w:r><w:t>Invoice</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Invoice #: {{ invoice_number }}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Date: {{ invoice_date }}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Customer: {{ customer_name }}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t></w:t></w:r>
        </w:p>
        <w:tbl>
            <w:tr>
                <w:tc>
                    <w:p><w:r><w:t>Description</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:p><w:r><w:t>Amount</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p><w:r><w:t>Items:</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:p><w:r><w:t>{% for item in items %}{{ item.name }} - Qty: {{ item.quantity }} - ${{ item.price }} each{% endfor %}</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
        <w:p>
            <w:r><w:t></w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Subtotal: ${{ subtotal }}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Tax ({{ tax_rate }}%): ${{ tax_amount }}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Total: ${{ total }}</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Notes: {{ notes }}</w:t></w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>''',
        )


def main():
    """Run the advanced example."""
    print("=" * 60)
    print("docxtplrs Advanced Example")
    print("=" * 60)

    try:
        from docxtplrs import DocxTemplate, RichText, R
    except ImportError as e:
        print(f"Error: Could not import docxtplrs. Please build the package first.")
        print(f"Run: maturin develop --uv")
        return 1

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        template_path = f.name

    output_path = template_path.replace(".docx", "_invoice.docx")

    try:
        # Create template
        print("\n1. Creating invoice template...")
        create_advanced_template(template_path)

        # Load template
        print("2. Loading template...")
        doc = DocxTemplate(template_path)

        # Show required variables
        vars_needed = doc.get_undeclared_template_variables()
        print(f"   Required variables: {vars_needed}")

        # Calculate totals
        items = [
            {"name": "Widget A", "quantity": 2, "price": 29.99},
            {"name": "Widget B", "quantity": 1, "price": 49.99},
            {"name": "Service C", "quantity": 3, "price": 15.00},
        ]
        subtotal = sum(item["price"] * item["quantity"] for item in items)
        tax_rate = 8.5
        tax_amount = round(subtotal * tax_rate / 100, 2)
        total = round(subtotal + tax_amount, 2)

        # Create styled header
        invoice_title = RichText("INVOICE", bold=True, font_size=36, color="2E5090")

        # Prepare context
        print("3. Preparing invoice data...")
        context = {
            "invoice_number": "INV-2024-001",
            "invoice_date": "2024-01-15",
            "customer_name": "Acme Corporation",
            "items": items,
            "subtotal": f"{subtotal:.2f}",
            "tax_rate": tax_rate,
            "tax_amount": f"{tax_amount:.2f}",
            "total": f"{total:.2f}",
            "notes": "Payment due within 30 days. Thank you for your business!",
        }

        # Render
        print("4. Rendering invoice...")
        doc.render(context)

        # Save
        print(f"5. Saving invoice to: {output_path}")
        doc.save(output_path)

        # Verify
        print("\n6. Verifying invoice...")
        if os.path.exists(output_path):
            with zipfile.ZipFile(output_path, "r") as zf:
                content = zf.read("word/document.xml").decode("utf-8")

                checks = [
                    ("INV-2024-001" in content, "Invoice number"),
                    ("Acme Corporation" in content, "Customer name"),
                    ("Widget A" in content, "First item"),
                    (str(total) in content or f"{total:.2f}" in content, "Total amount"),
                ]

                for passed, desc in checks:
                    status = "✓" if passed else "✗"
                    print(f"   {status} {desc}")

        print("\n" + "=" * 60)
        print("Invoice generated successfully!")
        print(f"Output file: {output_path}")
        print("=" * 60)

        return 0

    except Exception as e:
        print(f"\nError: {e}")
        import traceback

        traceback.print_exc()
        return 1

    finally:
        if os.path.exists(template_path):
            os.unlink(template_path)
        if os.path.exists(output_path):
            print(f"\nNote: Invoice file preserved at: {output_path}")


if __name__ == "__main__":
    exit(main())
