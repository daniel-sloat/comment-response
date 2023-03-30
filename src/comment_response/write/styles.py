"""Write comment-response section."""

from docx.document import Document
from docx.enum.base import EnumValue
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt
from docx.styles.style import _ParagraphStyle
from docx.styles.styles import Styles

PARAGRAPH_STYLE: EnumValue = WD_STYLE_TYPE.PARAGRAPH  # pylint: disable=no-member


def create_style(
    doc: Document,
    name: str,
    base_style: str = "Normal",
    left_indent: int | float = 0,
    space_before: int = 12,
    space_after: int = 12,
    next_style: str = "",
    keep_with_next: bool = False,
) -> None:
    """Create style in document."""
    styles: Styles = doc.styles
    style: _ParagraphStyle = styles.add_style(name, PARAGRAPH_STYLE)
    style.base_style = styles[base_style]
    style.paragraph_format.left_indent = Inches(left_indent)
    style.paragraph_format.space_before = Pt(space_before)
    style.paragraph_format.space_after = Pt(space_after)
    if next_style:
        style.next_paragraph_style = styles[next_style]
    if keep_with_next:
        style.paragraph_format.keep_with_next = keep_with_next
