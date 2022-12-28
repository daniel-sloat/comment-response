from reprlib import Repr

from xlsx.ooxml_ns import ns
from xlsx.cell.richtext import RichText
from xlsx.helpers.xl_position import xl_position


class Cell:
    def __init__(self, element, sheet):
        self._element = element
        self._sheet = sheet
        self._book = self._sheet._parent._book
        self._sharedstrings = self._book.sharedstrings
        self._styles = self._book.styles
        self._type = self._element.xpath("string(@t)")
        self._value = self._element.xpath("string(w:v)", **ns)
        self.reference = self._element.xpath("string(@r)")
        self._style = self._element.xpath("string(@s)")

    def __repr__(self):
        return (
            f"{self.__class__.__name__}("
            f"'{self.reference}',"
            f"pos={self.position},"
            f"value={Repr().repr(str(self.value))})"
        )

    @property
    def position(self):
        row = self._element.xpath("string(parent::w:row/@r)", **ns)
        row2, col = xl_position(self.reference)
        assert row == row2, "Reference position and row/columns do not match."
        return int(row), int(col)

    @property
    def formula(self):
        return self._element.xpath("string(w:f)", **ns)
    
    @property
    def style(self):
        return self._styles[int(self._style)]

    @property
    def value(self):
        match self._type:
            case "b":  # Boolean (0 or 1)
                return bool(self._value)
            case "inlineStr":
                return RichText(self._element.xpath("w:is", **ns), self._book)
            case "s":
                return self._sharedstrings[int(self._value)]
            case "e":
                return None
            case _:
                if not self._value:
                    return None
                elif self._value.isdigit():
                    return int(self._value)
                else:
                    try:
                        return float(self._value)
                    except ValueError:
                        return self._value


class DataCell(Cell):
    def __repr__(self):
        value_repr = str(self.value) if isinstance(self.value, RichText) else self.value
        return (
            f"{self.__class__.__name__}("
            f"'{self.reference}',"
            f"pos={self.position},"
            f"col={Repr().repr(self.column)},"
            f"value={Repr().repr(value_repr)})"
        )

    @property
    def column(self):
        return self._sheet.header.get(str(self.position[1]))
