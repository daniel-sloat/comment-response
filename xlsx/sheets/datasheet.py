from functools import cached_property

from xlsx.cell.cell import Cell
from xlsx.cell.datacell import DataCell
from xlsx.helpers.xl_position import xl_position_reverse
from xlsx.ooxml_ns import ns
from xlsx.sheets.record import Record
from xlsx.sheets.sheet import Sheet


class DataSheet(Sheet):
    def __init__(self, _name, sheets, header=1):
        super().__init__(_name, sheets)
        self._hrow = int(header)

    def __getitem__(self, key):
        return self.records[key]

    def __iter__(self):
        return iter(self.records)

    def __len__(self):
        return len(self.records)

    @cached_property
    def header(self):
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        elements = self.xml.xpath(xpath, _r=self._hrow, **ns)
        return {
            str(cell.position[1]): str(cell.value)
            for cell in [Cell(el, self) for el in elements]
        }

    def cell(self, row, col):
        xpath = "w:sheetData/w:row[@r=$_r]/w:c[@r=$_c]"
        if len(cell := self.xml.xpath(xpath, _r=row, _c=col, **ns)):
            return DataCell(cell[0], self)
        return None

    def row(self, row):
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        elements = self.xml.xpath(xpath, _r=row, **ns)
        return {cell.position: cell for cell in [DataCell(el, self) for el in elements]}

    def col(self, col):
        xl_col = xl_position_reverse(int(col))
        xpath = r"w:sheetData/w:row[@r>$_r]/w:c[re:test(@r,concat('^',$_c,'\d*$'))]"
        xpathvars = {"_r": self._hrow, "_c": xl_col}
        return [DataCell(el, self) for el in self.xml.xpath(xpath, **xpathvars, **ns)]

    @property
    def data(self):
        xpath = "w:sheetData/w:row[@r>$_r]/w:c"
        return [DataCell(el, self) for el in self.xml.xpath(xpath, _r=self._hrow, **ns)]

    @property
    def records(self):
        xpath = "w:sheetData/w:row[@r>$_r]"
        return {
            int(el.xpath("string(@r)")): Record(el, self)
            for el in self.xml.xpath(xpath, _r=self._hrow, **ns)
        }
