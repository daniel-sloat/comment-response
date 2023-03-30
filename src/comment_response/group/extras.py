"""Extra functions implemented while trying to get a good recursive data structure."""

from itertools import groupby

from comment_response.group.recursive_group import Heading


def key_sort(record, col_sort) -> Heading:
    """Sorts by single key column. (0, "Text")"""
    number_col, title_col = col_sort
    title = str(record.col.get(title_col, ""))
    num = int(record.col.get(number_col, 0))
    return Heading(num, title)


# BASIC RECURSIVE FUNCTION, INCLUDES ALL EMPTY TITLES
def group_records_basic(records, sort_cols):
    """Recursive sorting and grouping of records using specified columns."""
    new = {}
    current_cols, *remaining_cols = sort_cols
    keysort = lambda record: key_sort(record, current_cols)
    records = sorted(records, key=keysort)

    for sort_col_value, records in groupby(records, key=keysort):
        records = tuple(records)
        if remaining_cols:
            new[sort_col_value] = group_records_basic(records, remaining_cols)
        else:
            new[sort_col_value] = tuple(records)

    return new


# RECURSIVE LIST/DICTIONARY FUNCTION, USING SORT VALUE AS KEY
# MUST TEST FOR OBJECT TYPE WHILE UNPACKING
def group_records_type(records, sort_cols):
    """Recursive sorting and grouping of records using specified columns."""
    lst = []

    current_cols, *remaining_cols = sort_cols
    keysort = lambda record: key_sort(record, current_cols)
    records = sorted(records, key=keysort)

    for sort_col_value, records in groupby(records, key=keysort):
        new = {}
        if sort_col_value:
            if remaining_cols:
                new[sort_col_value] = group_records_type(records, remaining_cols)
            else:
                new[sort_col_value] = [tuple(records)]
        else:
            new = tuple(records)
        lst.append(new)
    return lst
