from functools import cached_property

from xlsx.cell.cell import DataCell
from xlsx.ooxml_ns import ns
from xlsx.helpers.xl_position import xl_position_reverse


class Record:
    def __init__(self, element, sheet):
        self.element = element
        self._sheet = sheet
        self.row_num = self.element.xpath("string(@r)", **ns)
        self.record_num = str(int(self.row_num) - int(self._sheet._hrow) - 1)
        self.header = self._sheet.header

    def __repr__(self):
        return f"{self.__class__.__name__}(num={self.record_num},row={self.row_num})"

    def __getitem__(self, key):
        return self.cells[key]

    def __iter__(self):
        return iter(self.cells.items())

    def __len__(self):
        return int(self.element.xpath("string(@spans)").split(":")[1])

    @cached_property
    def cells(self):
        cell = []
        for i in range(len(self)):
            ref = xl_position_reverse(i) + self.row_num
            cell.extend(self.element.xpath("w:c[@r=$_r]", _r=ref, **ns))
        return {
            cell.column: cell
            for cell in [DataCell(c, self._sheet) for c in cell]
            if cell.column
        }
