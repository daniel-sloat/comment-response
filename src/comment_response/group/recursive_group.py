"""
Recursive function with key sort to sort and group records.
"""

from itertools import groupby
from numbers import Number

from xlsx_rich_text.sheets.record import Record

from comment_response.group.sort_records import Heading


def column_sort(record: Record, columns: tuple[int, str]) -> Heading:
    """Key function to sort by single key column."""
    number_col, title_col = columns
    title = str(record.col.get(title_col, ""))
    num = int(record.col.get(number_col, 0))
    return Heading(num, title)


def comment_count_sort(sort_level: dict[str, list[dict]]) -> Number:
    """Key function to sort by comment count."""
    if "data" in sort_level:
        end = len(sort_level["data"]) == 1
        records = sort_level["data"][0].get("records")
        if records and end:
            return -len(records)
        return 0
    else:
        return float("-inf")


def group_records(
    records: list[Record], sort_cols: list[tuple[int, str]], count_sort: bool = False
) -> list[dict]:
    """Recursive sorting and grouping of records using specified columns."""
    group = []

    current_cols, *remaining_cols = sort_cols
    keysort = lambda record: column_sort(record, current_cols)
    records = sorted(records, key=keysort)

    for sort_col_value, records in groupby(records, key=keysort):
        info = {}
        if sort_col_value:
            info["heading"] = sort_col_value
            if remaining_cols:
                grouped_records = group_records(records, remaining_cols, count_sort)
                info["data"] = (
                    sorted(grouped_records, key=comment_count_sort)
                    if count_sort
                    else grouped_records
                )
            else:
                info["data"] = [{"records": tuple(records)}]
        else:
            info["records"] = tuple(records)
        group.append(info)
    return group
