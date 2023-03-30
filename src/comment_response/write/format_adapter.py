"""Adapt formats from XLSX to DOCX"""

from docx.enum.base import EnumValue
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_UNDERLINE
from docx.text.run import Run

PARAGRAPH_STYLE: EnumValue = WD_STYLE_TYPE.PARAGRAPH  # pylint: disable=no-member
DOUBLE_UNDERLINE_STYLE: EnumValue = WD_UNDERLINE.DOUBLE  # pylint: disable=no-member


def toggle(value_dict: dict[str:str]) -> bool:
    """Decode toggled property."""
    match value_dict:
        case {}:
            return True
        case {"val": value}:
            try:
                if str(value).casefold() == "true":
                    return True
                return False
            except TypeError:
                return bool(int(value))
        case _:
            return False


def format_adapter(tag: dict | None, run: Run) -> None:
    """Adapt format properties from XLSX to DOCX."""
    if "b" in tag:
        run.font.bold = toggle(tag["b"])
    if "i" in tag:
        run.font.italic = toggle(tag["i"])
    if "u" in tag:
        match tag["u"]:
            case {"val": _type}:
                string = str(_type).casefold()
                if string == "double" or string == "wavydouble":
                    run.font.underline = DOUBLE_UNDERLINE_STYLE
                elif string == "none":
                    run.font.underline = False
                else:
                    run.font.underline = True
            case _:
                run.font.underline = toggle(tag["u"])
    if "strike" in tag:
        if tag.get("color", {}).get("rgb") == "FFFF0000":
            run.font.double_strike = toggle(tag["strike"])
        else:
            run.font.strike = toggle(tag["strike"])
    if "vertAlign" in tag:
        match tag["vertAlign"]:
            case {"val": "superscript"}:
                run.font.superscript = True
            case {"val": "subscript"}:
                run.font.subscript = True
