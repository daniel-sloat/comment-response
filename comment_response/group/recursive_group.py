"""
Recursive function with key sort to sort and group records.
"""

from itertools import groupby
from typing import TypeAlias

from comment_response.group.colsort import ColSort

ColumnSort: TypeAlias = tuple[int, str]
ColumnTuples: TypeAlias = tuple[tuple[str, str], ...]


def col_sort(record, col_sort) -> ColSort:
    """Sorts by single key column. (0, "Text")"""
    number_col, title_col = col_sort
    title = str(record.col.get(title_col, ""))
    num = int(record.col.get(number_col, 0))
    return ColSort(num, title)


def comment_count_sort(x):
    if "data" in x:
        end = len(x["data"]) == 1
        records = x["data"][0].get("records")
        if records and end:
            return -len(records)
        return 0
    else:
        return float("-inf")


# RECURSIVE LIST/DICTIONARY FUNCTION
def group_records(records, sort_cols, count_sort=False):
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
