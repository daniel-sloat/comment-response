"""DataSheet"""

from functools import cached_property
from typing import TYPE_CHECKING

from xlsx.cell.cell import Cell
from xlsx.cell.datacell import DataCell
from xlsx.helpers.xl_position import xl_position_reverse
from xlsx.ooxml_ns import ns
from xlsx.sheets.record import Record
from xlsx.sheets.sheet import Sheet

if TYPE_CHECKING:
    from xlsx.sheets.sheets import Sheets


class DataSheet(Sheet):
    """Representation of worksheet as table (structured), with header."""

    def __init__(self, _name: str, sheets: "Sheets", header_row: int | str = 1):
        super().__init__(_name, sheets)
        self.header_row = int(header_row)

    def __getitem__(self, key):
        return self.records[key]

    def __iter__(self):
        return iter(self.records)

    def __len__(self):
        return len(self.records)

    @cached_property
    def header(self) -> dict[str:str]:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        elements = self.xml.xpath(xpath, _r=self.header_row, **ns)
        return {
            str(cell.position[1]): str(cell.value)
            for cell in [Cell(el, self) for el in elements]
        }

    def cell(self, row: str, col: str) -> DataCell | None:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c[@r=$_c]"
        if len(cell := self.xml.xpath(xpath, _r=row, _c=col, **ns)):
            return DataCell(cell[0], self)
        return None

    def row(self, row: str) -> dict[tuple[int, int] : DataCell]:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        elements = self.xml.xpath(xpath, _r=row, **ns)
        return {cell.position: cell for cell in [DataCell(el, self) for el in elements]}

    def col(self, col: str) -> list[DataCell]:
        xl_col = xl_position_reverse(int(col))
        xpath = r"w:sheetData/w:row[@r>$_r]/w:c[re:test(@r,concat('^',$_c,'\d*$'))]"
        xpathvars = {"_r": self.header_row, "_c": xl_col}
        return [DataCell(el, self) for el in self.xml.xpath(xpath, **xpathvars, **ns)]

    @property
    def data(self) -> list[DataCell]:
        xpath = "w:sheetData/w:row[@r>$_r]/w:c"
        return [
            DataCell(el, self) for el in self.xml.xpath(xpath, _r=self.header_row, **ns)
        ]

    @property
    def records(self) -> dict[int:Record]:
        xpath = "w:sheetData/w:row[@r>$_r]"
        return {
            int(el.xpath("string(@r)")): Record(el, self)
            for el in self.xml.xpath(xpath, _r=self.header_row, **ns)
        }
