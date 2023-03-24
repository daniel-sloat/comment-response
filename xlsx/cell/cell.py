"""Cell-level class."""

from reprlib import Repr
from typing import TYPE_CHECKING

from lxml.etree import _Element

from xlsx.cell.richtext import RichText
from xlsx.helpers.xl_position import xl_position
from xlsx.ooxml_ns import ns

if TYPE_CHECKING:
    from xlsx.sheets.sheet import Sheet
    from xlsx.sheets.newdatasheet import NewSheet


class Cell:
    """Representation of cell in OOXML."""

    def __init__(self, element: _Element, sheet: "NewSheet"):
        self._element = element
        self._sheet = sheet
        self._book = self._sheet.workbook
        self._sharedstrings = self._book.sharedstrings
        self._styles = self._book.styles

    def __repr__(self):
        return (
            f"{self.__class__.__name__}("
            f"'{self.reference}',"
            f"pos={self.position},"
            f"value={Repr().repr(str(self.value))})"
        )

    def __str__(self):
        if self.value:
            return str(self.value)
        return ""

    def __int__(self):
        if self.value:
            return int(self.value)
        return 0

    def __lt__(self, other):
        return self.value < other.value

    @property
    def reference(self):
        return self._element.xpath("string(@r)")

    @property
    def position(self):
        row, col = xl_position(self.reference)
        return int(row), int(col)

    @property
    def formula(self):
        return self._element.xpath("string(w:f)", **ns)

    @property
    def style(self):
        style_num = self._element.xpath("string(@s)")
        return self._styles[int(style_num)]

    @property
    def value(self):
        value_xml = self._element.xpath("string(w:v)", **ns)
        value_type = self._element.xpath("string(@t)")
        match value_type:
            case "b":  # Boolean (0 or 1)
                return bool(value_xml)
            case "inlineStr":
                return RichText(self._element.xpath("w:is", **ns))
            case "s":
                return self._sharedstrings[int(value_xml)]
            case "e":
                return None
            case _:
                if not value_xml:
                    return None
                elif value_xml.isdigit():
                    return int(value_xml)
                else:
                    try:
                        return float(value_xml)
                    except ValueError:
                        msg = f"Cannot detect cell value: {self.reference}"
                        raise TypeError(msg)  # pylint: disable=raise-missing-from
