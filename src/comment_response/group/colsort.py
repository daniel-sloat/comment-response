from dataclasses import dataclass
from itertools import groupby

from xlsx_rich_text.sheets.record import Record


@dataclass(frozen=True, order=True)
class ColSort:
    """For column sorting. Displays textual information in repr, but stores custom sort
    information as descriptor."""

    num: int
    title: str

    def __bool__(self):
        return any((bool(self.num), bool(self.title)))


class CommentSort:
    def __init__(self, record, sort_cols):
        self.record: Record = record
        self._sort_cols = sort_cols

    def __repr__(self):
        return f"{self.__class__.__name__}(record={self.record})"

    def __lt__(self, other):
        return self.sort_values < other.sort_values

    @property
    def sort_values(self):
        new = []
        for cols in self._sort_cols:
            number_col, title_col = cols
            title = str(self.record.col.get(title_col, ""))
            num = int(self.record.col.get(number_col, 0))
            if num or title:
                new.append((num, title))
        return new

    @property
    def sort_cols(self):
        return self._sort_cols[0 : len(self.sort_values)]


class GroupedRecordsSort:
    def __init__(self, records, sort_values):
        self.records = tuple(records)
        self.sort_values = sort_values
        self.current_cols, *self.remaining_cols = self.sort_values

    def __repr__(self):
        return (
            f"{self.__class__.__name__}("
            f"count={len(self.records)},"
            f"sort_values={self.sort_values})"
        )

    @staticmethod
    def key_sort(record, col_sort) -> ColSort:
        """Sorts by single key column. (0, "Text")"""
        number_col, title_col = col_sort
        title = str(record.col.get(title_col, ""))
        num = int(record.col.get(number_col, 0))
        return ColSort(num, title)

    @property
    def recurse(self):
        keyfunc = lambda record: self.key_sort(record, self.current_cols)
        for col_sort, records in groupby(self.records, key=keyfunc):
            if self.remaining_cols:
                return {
                    col_sort: GroupedRecordsSort(records, self.remaining_cols).recurse
                }
            else:
                return self.records
