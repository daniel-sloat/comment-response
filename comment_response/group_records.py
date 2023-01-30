"""Grouping for datasheet records."""

from dataclasses import dataclass
from itertools import groupby
from typing import TypeAlias

ColumnTuples: TypeAlias = tuple[tuple[str, str], ...]


@dataclass(frozen=True, order=True)
class ColSort:
    """For column sorting. Displays textual information in repr, but stores custom sort
    information as descriptor."""

    num: int
    title: str


def key_sort(record, col_sort) -> tuple[int, str]:
    number_col, title_col = col_sort
    title = str(record.col.get(title_col, ""))
    num = int(record.col.get(number_col, 0))
    return ColSort(num, title)


def group_records(records, sort_cols: ColumnTuples) -> dict:
    """Recursive sorting and grouping of records using specified columns."""
    current_cols, *remaining_cols = sort_cols
    key_func = lambda record: key_sort(record, current_cols)
    records = sorted(records, key=key_func)

    new = {}
    for key, group in groupby(records, key=key_func):
        if remaining_cols:
            new[key] = group_records(group, remaining_cols)
        else:
            new[key] = tuple(group)
    return new


class GroupRecords:
    """Groups records within datasheet based on column values. Two sort columns are used:
    one for alphabetic sorting, and another that can be used for custom sorting."""

    def __init__(self, datasheet, **config):
        self._sheet = datasheet
        self._config = config
        self.title_sort = config.get("columns", {}).get("sort", [])
        self.ordered_sort = config.get("columns", {}).get("numbered_sort", [])
        # self.sort_by_comment_count = True

    def __repr__(self):
        return f"{self.__class__.__name__}(records={len(self._sheet)},sort={self.sort})"

    @property
    def sort(self):
        return tuple(zip(self.ordered_sort, self.title_sort))

    def group(self):
        return group_records(self._sheet.records.values(), self.sort)
