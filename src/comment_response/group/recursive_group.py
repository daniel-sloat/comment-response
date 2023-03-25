"""
Recursive function with key sort to sort and group records.
"""

from itertools import groupby
from numbers import Number
from typing import TypeAlias

from comment_response.group.colsort import ColSort
from xlsx_rich_text.sheets.record import Record

ColumnSort: TypeAlias = tuple[int, str]
ColumnTuples: TypeAlias = tuple[tuple[str, str], ...]


def col_sort(record: Record, columns: tuple[int, str]) -> ColSort:
    """Sorts by single key column."""
    number_col, title_col = columns
    title = str(record.col.get(title_col, ""))
    num = int(record.col.get(number_col, 0))
    return ColSort(num, title)


def comment_count_sort(sort_level) -> Number:
    if "data" in sort_level:
        end = len(sort_level["data"]) == 1
        records = sort_level["data"][0].get("records")
        if records and end:
            return -len(records)
        return 0
    else:
        return float("-inf")


# RECURSIVE LIST/DICTIONARY FUNCTION
def group_records(records, sort_cols, count_sort=False) -> list[dict]:
    """Recursive sorting and grouping of records using specified columns."""
    lst = []

    current_cols, *remaining_cols = sort_cols
    keysort = lambda record: col_sort(record, current_cols)
    records = sorted(records, key=keysort)

    for sort_col_value, records in groupby(records, key=keysort):
        new = {}
        if sort_col_value:
            new["heading"] = sort_col_value
            if remaining_cols:
                grouped_records = group_records(records, remaining_cols, count_sort)
                new["data"] = (
                    sorted(grouped_records, key=comment_count_sort)
                    if count_sort
                    else grouped_records
                )
            else:
                new["data"] = [{"records": tuple(records)}]
        else:
            new["records"] = tuple(records)
        lst.append(new)
    return lst
