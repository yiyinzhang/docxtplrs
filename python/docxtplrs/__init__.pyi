"""
docxtplrs - A Rust implementation of python-docx-template with Python bindings
"""

from typing import Any, Dict, List, Optional, Set, Union

__version__: str

class DocxTemplate:
    """Load and render Word document templates using Jinja2 syntax."""

    def __init__(self, template_path: str) -> None:
        """
        Load a .docx template file.

        Args:
            template_path: Path to the .docx template file
        """
        ...

    def render(
        self,
        context: Dict[str, Any],
        jinja_env: Any = None,
        autoescape: bool = False,
    ) -> None:
        """
        Render the template with context variables.

        Args:
            context: Dictionary of variables for Jinja2 templating
            jinja_env: Optional custom Jinja2 environment
            autoescape: Whether to autoescape special characters (default: False)
        """
        ...

    def save(self, output_path: str) -> None:
        """
        Save the rendered document to a file.

        Args:
            output_path: Path where the generated document will be saved
        """
        ...

    def get_undeclared_template_variables(
        self, context: Optional[Dict[str, Any]] = None
    ) -> List[str]:
        """
        Get the set of undeclared template variables.

        Args:
            context: Optional context dict. If provided, returns variables not in context.

        Returns:
            List of variable names that need to be defined
        """
        ...

    def new_subdoc(self, docx_path: Optional[str] = None) -> Subdoc:
        """
        Create a new subdoc for embedding.

        Args:
            docx_path: Optional path to an existing .docx file to load as subdoc

        Returns:
            A new Subdoc instance
        """
        ...

    def build_url_id(self, url: str) -> str:
        """
        Build URL ID for hyperlinks.

        Args:
            url: The URL to link to

        Returns:
            A relationship ID for use with RichText.add_link()
        """
        ...

    def replace_pic(self, dummy_pic_name: str, new_pic_path: str) -> None:
        """
        Replace a picture in the document.

        Args:
            dummy_pic_name: Name of the dummy picture in the template
            new_pic_path: Path to the replacement picture
        """
        ...

    def replace_media(self, dummy_media: str, new_media: str) -> None:
        """
        Replace media in the document.

        Args:
            dummy_media: Name of the dummy media in the template
            new_media: Path to the replacement media
        """
        ...

    def replace_embedded(self, dummy_embedded: str, new_embedded: str) -> None:
        """
        Replace embedded objects in the document.

        Args:
            dummy_embedded: Name of the dummy embedded object
            new_embedded: Path to the replacement embedded object
        """
        ...

    def replace_zipname(self, zip_name: str, new_file: str) -> None:
        """
        Replace content by zip name.

        Args:
            zip_name: The internal zip path to replace
            new_file: Path to the replacement file
        """
        ...

    def reset_replacements(self) -> None:
        """Reset all replacements (for multiple renderings)."""
        ...

    def get_xml(self) -> str:
        """Get a preview of the document XML (for debugging)."""
        ...

class RichText:
    """Create styled text content for Word documents."""

    def __init__(
        self,
        text: Optional[str] = None,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        strike: bool = False,
        font: Optional[str] = None,
        font_size: Optional[int] = None,
        color: Optional[str] = None,
        highlight: Optional[str] = None,
        caps: bool = False,
        small_caps: bool = False,
    ) -> None:
        """
        Create a new RichText object.

        Args:
            text: Initial text content
            **kwargs: Style options (bold, italic, underline, strike, font,
                      font_size, color, highlight, caps, small_caps)
        """
        ...

    def add(
        self,
        text: str,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        strike: bool = False,
        font: Optional[str] = None,
        font_size: Optional[int] = None,
        color: Optional[str] = None,
        highlight: Optional[str] = None,
        caps: bool = False,
        small_caps: bool = False,
    ) -> None:
        """
        Add text with optional styling.

        Args:
            text: Text to add
            **kwargs: Style options for this text fragment
        """
        ...

    def add_link(self, text: str, url_id: str, **kwargs: Any) -> None:
        """
        Add a hyperlink.

        Args:
            text: Link text
            url_id: The URL ID (obtained from DocxTemplate.build_url_id())
            **kwargs: Style options for this link
        """
        ...

    def add_newline(self) -> None:
        """Add a newline (line break within paragraph)."""
        ...

    def add_paragraph(self) -> None:
        """Add a new paragraph."""
        ...

    def add_tab(self) -> None:
        """Add a tab character."""
        ...

    def add_page_break(self) -> None:
        """Add a page break."""
        ...

    def __str__(self) -> str:
        ...

    def __repr__(self) -> str:
        ...

class RichTextParagraph:
    """Create styled paragraphs for Word documents."""

    def __init__(self) -> None:
        """Create a new RichTextParagraph object."""
        ...

    def add_rt(self, rt: RichText) -> None:
        """Add RichText to this paragraph."""
        ...

    @property
    def style(self) -> Optional[str]:
        """Get/set paragraph style."""
        ...

    @style.setter
    def style(self, style: str) -> None:
        ...

    @property
    def alignment(self) -> Optional[str]:
        """Get/set paragraph alignment."""
        ...

    @alignment.setter
    def alignment(self, alignment: str) -> None:
        ...

class InlineImage:
    """Insert images into Word documents."""

    def __init__(
        self,
        template: DocxTemplate,
        image_descriptor: str,
        width: Optional[Union[Mm, Inches, Pt, float]] = None,
        height: Optional[Union[Mm, Inches, Pt, float]] = None,
    ) -> None:
        """
        Create a new InlineImage object.

        Args:
            template: The DocxTemplate object (for relationship management)
            image_descriptor: Path to the image file
            width: Optional width (use Mm, Inches, or Pt classes)
            height: Optional height (use Mm, Inches, or Pt classes)
        """
        ...

    def __repr__(self) -> str:
        ...

class Subdoc:
    """Embed sub-documents within documents."""

    def __init__(self) -> None:
        """Create a new empty Subdoc."""
        ...

    def __repr__(self) -> str:
        ...

class DocumentBuilder:
    """Build document content programmatically."""

    def __init__(self) -> None:
        """Create a new DocumentBuilder."""
        ...

    def add_paragraph(self, text: str) -> None:
        """Add a paragraph."""
        ...

    def add_heading(self, text: str, level: int) -> None:
        """Add a heading."""
        ...

    def add_run(self, text: str, bold: bool = False, italic: bool = False) -> None:
        """Add text to current paragraph."""
        ...

    def build(self) -> Subdoc:
        """Build and return the subdoc."""
        ...

    def __repr__(self) -> str:
        ...

class Listing:
    """Insert escaped text with formatting."""

    def __init__(self, text: str) -> None:
        """
        Create a new Listing object.

        Args:
            text: The text content with special characters
        """
        ...

    def __str__(self) -> str:
        ...

class CellColor:
    """Specify cell background color for tables."""

    def __init__(self, color: str) -> None:
        """
        Create a new CellColor object.

        Args:
            color: Hex color code (with or without #)
        """
        ...

class ColSpan:
    """Specify column span for table cells."""

    def __init__(self, span: int) -> None:
        """
        Create a new ColSpan object.

        Args:
            span: Number of columns to span
        """
        ...

class VerticalMerge:
    """Specify vertical merge for table cells."""

    def __init__(self, merge: bool) -> None:
        """
        Create a new VerticalMerge object.

        Args:
            merge: True to continue merge, False to start new merge
        """
        ...

class Mm:
    """Millimeters measurement."""

    def __init__(self, value: float) -> None:
        """Create a millimeters measurement."""
        ...

    def __float__(self) -> float:
        ...

    def __repr__(self) -> str:
        ...

class Inches:
    """Inches measurement."""

    def __init__(self, value: float) -> None:
        """Create an inches measurement."""
        ...

    def __float__(self) -> float:
        ...

    def __repr__(self) -> str:
        ...

class Pt:
    """Points measurement."""

    def __init__(self, value: float) -> None:
        """Create a points measurement."""
        ...

    def __float__(self) -> float:
        ...

    def __repr__(self) -> str:
        ...

def R(
    text: Optional[str] = None,
    *,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    strike: bool = False,
    font: Optional[str] = None,
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    highlight: Optional[str] = None,
    caps: bool = False,
    small_caps: bool = False,
) -> RichText:
    """Shortcut for RichText."""
    ...

def RP() -> RichTextParagraph:
    """Shortcut for RichTextParagraph."""
    ...

def escape_xml(text: str) -> str:
    """Escape XML special characters."""
    ...

def unescape_xml(text: str) -> str:
    """Unescape XML entities."""
    ...

def version() -> str:
    """Get the version of the docxtplrs library."""
    ...
