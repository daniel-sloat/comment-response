from comment_response.group.colsort import ColSort
from itertools import groupby
from typing import TypeAlias

ColumnTuples: TypeAlias = tuple[tuple[str, str], ...]


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
