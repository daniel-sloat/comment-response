"""New DataSheet"""

from functools import cached_property
from typing import TYPE_CHECKING

from xlsx.cell.cell import Cell
from xlsx.cell.datacell import DataCell
from xlsx.helpers.xl_position import xl_position_reverse
from xlsx.ooxml_ns import ns
from xlsx.sheets.record import Record

if TYPE_CHECKING:
    from xlsx.workbook import Workbook


class NewSheet:
    """Create Sheet, basic access to worksheet."""

    def __init__(self, sheetname: str, workbook: "Workbook"):
        self.sheetname = sheetname
        self.workbook = workbook
        self.sheet_xml = self.workbook.xlsx.sheet(self.sheetname)
        if not self._id:
            raise ValueError("Worksheet not found in book.")

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(name='{self.sheetname}')"

    @property
    def _id(self) -> str:
        xpath = "string(w:sheets/w:sheet[@name=$_name]/@sheetId)"
        return self.workbook.xlsx.workbook.xpath(xpath, _name=self.sheetname, **ns)


class NewDataSheet(NewSheet):
    """Create DataSheet, with access for each record by header column."""

    def __init__(
        self,
        sheetname: str,
        workbook: "Workbook",
        header_row: int = 1,
    ):
        super().__init__(sheetname, workbook)
        self.header_row = header_row

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(name='{self.sheetname}', header_row={self.header_row})"

    @cached_property
    def header(self) -> dict[str, str]:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        elements = self.sheet_xml.xpath(xpath, _r=self.header_row, **ns)
        return {
            str(cell.position[1]): str(cell.value)
            for cell in [Cell(el, self) for el in elements]
        }

    def cell(self, row: str, col: str) -> DataCell | None:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c[@r=$_c]"
        if len(cell := self.sheet_xml.xpath(xpath, _r=row, _c=col, **ns)):
            return DataCell(cell[0], self)
        return None

    def row(self, row: str) -> dict[tuple[int, int], DataCell]:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        elements = self.sheet_xml.xpath(xpath, _r=row, **ns)
        return {cell.position: cell for cell in [DataCell(el, self) for el in elements]}

    def col(self, col: str) -> list[DataCell]:
        xl_col = xl_position_reverse(int(col))
        xpath = r"w:sheetData/w:row[@r>$_r]/w:c[re:test(@r,concat('^',$_c,'\d*$'))]"
        xpathvars = {"_r": self.header_row, "_c": xl_col}
        return [
            DataCell(el, self) for el in self.sheet_xml.xpath(xpath, **xpathvars, **ns)
        ]

    @property
    def data(self) -> list[DataCell]:
        xpath = "w:sheetData/w:row[@r>$_r]/w:c"
        return [
            DataCell(el, self)
            for el in self.sheet_xml.xpath(xpath, _r=self.header_row, **ns)
        ]

    @property
    def records(self) -> dict[int, Record]:
        xpath = "w:sheetData/w:row[@r>$_r]"
        return {
            int(el.xpath("string(@r)")): Record(el, self)
            for el in self.sheet_xml.xpath(xpath, _r=self.header_row, **ns)
        }
