"""Sheet"""

from xlsx.cell.cell import Cell
from xlsx.ooxml_ns import ns
from xlsx.xml import XLSXXML


class Sheet:
    """Representation of sheet.xml"""

    def __init__(self, _name: str, xlsx: XLSXXML) -> None:
        self._name = _name
        self.xlsx = xlsx
        self.xml = self.xlsx.sheet(self._name)

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(name='{self._name}')"

    @property
    def _id(self) -> str:
        xpath = "string(w:sheets/w:sheet[@name=$_name]/@sheetId)"
        return self.xlsx.workbook.xpath(xpath, _name=self._name, **ns)

    def cell(self, row: int | str, col: int | str) -> Cell | None:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c[@r=$_c]"
        if len(cell := self.xml.xpath(xpath, _r=row, _c=col, **ns)):
            return Cell(cell[0], self)
        return None

    def row(self, row: int | str) -> list[Cell]:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        return [Cell(el, self) for el in self.xml.xpath(xpath, _r=row, **ns)]

    def col(self, col: int | str) -> list[Cell]:
        xpath = "w:sheetData/w:row/w:c[@r=$_c]"
        return [Cell(el, self) for el in self.xml.xpath(xpath, _c=col, **ns)]
