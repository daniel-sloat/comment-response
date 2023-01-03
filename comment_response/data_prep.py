from itertools import groupby


def keysort(record, sort_col):
    if record.cells.get(sort_col) and record[sort_col].value:
        return str(record[sort_col].value).strip()
    else:
        return ""


def group_records(records, col):
    key = lambda record: keysort(record, col)
    sorted_records = sorted(records, key=key)
    return groupby(sorted_records, key=key)


# def recurse_group_all(records, sort_cols):
#     for key, group in group_records(records, sort_cols[0]):
#         if len(sort_cols) > 1:
#             yield key, recursion(group, sort_cols[1:])
#         else:
#             yield key, group


def recurse_group(records, sort_cols):
    for key, group in group_records(records, sort_cols[0]):
        if len(sort_cols) > 1:
            yield key, tuple(recurse_group(group, sort_cols[1:]))
        elif key:
            yield key, tuple(group)
        else:
            yield tuple(group)


class PrepData:
    def __init__(self, datasheet, **config):
        self.sheet = datasheet
        self._config = config
        self.sort = config.get("columns", {}).get("sort", [])

    def __repr__(self):
        return f"{self.__class__.__name__}(records={len(self.sheet)},sort={self.sort})"

    def __iter__(self):
        return iter(self.grouped_records())

    def grouped_records(self):
        return recurse_group(self.sheet.records, self.sort)
