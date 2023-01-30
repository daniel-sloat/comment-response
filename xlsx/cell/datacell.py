"""DataCell subclass of Cell."""

from reprlib import Repr

from xlsx.cell.cell import Cell
from xlsx.cell.richtext import RichText


class DataCell(Cell):
    """DataCell for DataSheet object."""

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
