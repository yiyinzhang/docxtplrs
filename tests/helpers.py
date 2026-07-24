"""Helpers to build minimal .docx files for testing docxtplrs."""

import io
import struct
import zipfile
import zlib

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

NSDECL = (
    f'xmlns:w="{W}" xmlns:r="{R}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}"'
)


def document_xml(body: str) -> str:
    return (
        XML_DECL
        + f"<w:document {NSDECL}><w:body>{body}<w:sectPr/></w:body></w:document>"
    )


def hdrftr_xml(tag: str, body: str) -> str:
    return XML_DECL + f"<w:{tag} {NSDECL}>{body}</w:{tag}>"


def run(text: str, rpr: str = "") -> str:
    """A run; text is used raw (must be xml-escaped already if needed)."""
    rpr_xml = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f"<w:r>{rpr_xml}<w:t>{text}</w:t></w:r>"


def p(*parts: str, ppr: str = "") -> str:
    """A paragraph from runs (or raw xml)."""
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{''.join(parts)}</w:p>"


def tp(*parts: str) -> str:
    """A paragraph of plain text runs."""
    return p(*[run(x) for x in parts])


def cell(*paras: str, tcpr: str = "") -> str:
    tcpr_xml = f"<w:tcPr>{tcpr}</w:tcPr>" if tcpr else "<w:tcPr/>"
    return f"<w:tc>{tcpr_xml}{''.join(paras)}</w:tc>"


def tr(*cells: str) -> str:
    return f"<w:tr>{''.join(cells)}</w:tr>"


def tbl(rows, widths=(2000, 2000)):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return f"<w:tbl><w:tblGrid>{grid}</w:tblGrid>{''.join(rows)}</w:tbl>"


CT_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

RELS_TMPL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"""

HEADER_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FOOTER_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
FOOTNOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

HDR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FTR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
FOOTNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

CORE_CT = "application/vnd.openxmlformats-package.core-properties+xml"
CORE_RT = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>{title}</dc:title><dc:creator>{author}</dc:creator><dc:subject></dc:subject>
</cp:coreProperties>"""


def make_docx(
    body: str,
    headers: dict = None,  # {"header1.xml": content}
    footers: dict = None,
    footnotes: str = None,
    core: str = None,
    media: dict = None,  # {"image1.png": bytes}
    extra_files: dict = None,
    styles: str = None,  # contents of <w:styles> element
    numbering: str = None,  # full numbering.xml
):
    """Build a minimal docx (as bytes) with a body and optional parts."""
    buf = io.BytesIO()
    z = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)

    ct = CT_BASE
    doc_rels = []
    rid = 1

    for name in (headers or {}):
        ct = ct.replace(
            "</Types>",
            f'<Override PartName="/word/{name}" ContentType="{HDR_CT}"/></Types>',
        )
        doc_rels.append((f"rId{rid}", HEADER_RT, name))
        rid += 1
    for name in (footers or {}):
        ct = ct.replace(
            "</Types>",
            f'<Override PartName="/word/{name}" ContentType="{FTR_CT}"/></Types>',
        )
        doc_rels.append((f"rId{rid}", FOOTER_RT, name))
        rid += 1
    if footnotes is not None:
        ct = ct.replace(
            "</Types>",
            f'<Override PartName="/word/footnotes.xml" ContentType="{FOOTNOTES_CT}"/></Types>',
        )
        doc_rels.append((f"rId{rid}", FOOTNOTES_RT, "footnotes.xml"))
        rid += 1
    if styles is not None:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        )
        doc_rels.append((f"rId{rid}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"))
        rid += 1
    if numbering is not None:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
        )
        doc_rels.append((f"rId{rid}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml"))
        rid += 1

    if core is not None:
        ct = ct.replace(
            "</Types>",
            f'<Override PartName="/docProps/core.xml" ContentType="{CORE_CT}"/></Types>',
        )

    if media:
        for name, blob in media.items():
            ext = name.rsplit(".", 1)[-1]
            mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(
                ext, "application/octet-stream"
            )
            ct = ct.replace(
                "</Types>",
                f'<Default Extension="{ext}" ContentType="{mime}"/></Types>',
            )
            z.writestr(f"word/media/{name}", blob)

    z.writestr("[Content_Types].xml", ct)

    root_rels = ROOT_RELS
    if core is not None:
        root_rels = root_rels.replace(
            "</Relationships>",
            f'<Relationship Id="rId99" Type="{CORE_RT}" Target="docProps/core.xml"/></Relationships>',
        )
    z.writestr("_rels/.rels", root_rels)

    z.writestr("word/document.xml", document_xml(body))

    for name, content in (headers or {}).items():
        z.writestr(f"word/{name}", hdrftr_xml("hdr", content))
    for name, content in (footers or {}).items():
        z.writestr(f"word/{name}", hdrftr_xml("ftr", content))
    if footnotes is not None:
        z.writestr(
            "word/footnotes.xml",
            XML_DECL + f'<w:footnotes {NSDECL}>{footnotes}</w:footnotes>',
        )
    if styles is not None:
        z.writestr("word/styles.xml", XML_DECL + f'<w:styles {NSDECL}>{styles}</w:styles>')
    if numbering is not None:
        z.writestr("word/numbering.xml", numbering)
    if core is not None:
        z.writestr("docProps/core.xml", core)

    if doc_rels:
        rels_xml = RELS_TMPL.format(
            "".join(
                f'<Relationship Id="{i}" Type="{t}" Target="{tg}"/>'
                for i, t, tg in doc_rels
            )
        )
        z.writestr("word/_rels/document.xml.rels", rels_xml)

    for name, blob in (extra_files or {}).items():
        z.writestr(name, blob)

    z.close()
    return buf.getvalue()


def read_docx_part(docx_bytes: bytes, name: str) -> str:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        return z.read(name).decode("utf-8")


def docx_names(docx_bytes: bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        return z.namelist()


def text_of(xml: str) -> str:
    """Extract concatenated text of <w:t> elements (crude)."""
    import re

    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def make_png(w: int = 4, h: int = 3) -> bytes:
    """A minimal valid RGB PNG (no pHYs -> 72 dpi default)."""

    def chunk(t, d):
        return (
            struct.pack(">I", len(d))
            + t
            + d
            + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * w for _ in range(h))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
