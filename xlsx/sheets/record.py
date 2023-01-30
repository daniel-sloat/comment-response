from functools import cached_property

from xlsx.cell.datacell import DataCell
from xlsx.ooxml_ns import ns


class Record:
    def __init__(self, element, sheet):
        self.element = element
        self._sheet = sheet
        self.row_num = self.element.xpath("string(@r)", **ns)
        self.record_num = str(int(self.row_num) - int(self._sheet._hrow) - 1)

    def __repr__(self):
        return f"{self.__class__.__name__}(row={self.row_num})"

    def __getitem__(self, key):
        return self.col.get(key)

    def __iter__(self):
        return iter(self.col)

    def __len__(self):
        return len(self.col)

    def __lt__(self, other):
        return self.record_num < other.record_num

    @cached_property
    def col(self):
        return {
            cell.column: cell
            for cell in [DataCell(c, self._sheet) for c in self.element]
        }
