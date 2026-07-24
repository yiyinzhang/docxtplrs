"""docxtplrs demo: render a small docx template entirely in Rust."""

import io
import zipfile

from docxtplrs import DocxTemplate, RichText, Listing, Mm, InlineImage
from tests.helpers import make_docx, tp, tbl, tr, cell, read_docx_part, text_of


def main():
    body = (
        tp("Hello {{ name }}!")
        + tp("{%p for item in items %}")
        + tp("- {{ item }}")
        + tp("{%p endfor %}")
        + tp("{{r rt }}")
        + tp("{{ listing }}")
        + tbl(
            [
                tr(cell(tp("{%tr for x in rows %}")), cell(tp(""))),
                tr(cell(tp("{{ x }}")), cell(tp("cell"))),
                tr(cell(tp("{%tr endfor %}")), cell(tp(""))),
            ]
        )
    )

    tpl = DocxTemplate(io.BytesIO(make_docx(body)))
    tpl.render(
        {
            "name": "docxtplrs",
            "items": ["rust", "jinja", "docx"],
            "rows": [1, 2, 3],
            "rt": RichText("styled text", bold=True, color="#C00000"),
            "listing": Listing("line1\nline2"),
        }
    )
    tpl.save("demo_out.docx")

    with zipfile.ZipFile("demo_out.docx") as z:
        xml = z.read("word/document.xml").decode()
    print(text_of(xml))
    print("saved demo_out.docx")


if __name__ == "__main__":
    main()
