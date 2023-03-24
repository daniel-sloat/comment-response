"""Grouping for datasheet records."""

from comment_response.group.recursive_group import group_records
from xlsx_rich_text.sheets.newdatasheet import NewDataSheet


class GroupRecords:
    """Groups records within datasheet based on column values. Two sort columns are used:
    one for alphabetic sorting, and another that can be used for custom sorting."""

    def __init__(self, datasheet: "NewDataSheet", **config):
        self._sheet = datasheet
        self._config = config
        self.title_sort = config.get("columns", {}).get("sort", [])
        self.ordered_sort = config.get("columns", {}).get("numbered_sort", [])
        self.sort_by_comment_count = config.get("sort", {}).get(
            "by_comment_count", True
        )

    def __repr__(self):
        return f"{self.__class__.__name__}(sort={self.sort})"

    @property
    def sort(self):
        return tuple(zip(self.ordered_sort, self.title_sort))

    def group(self):
        return group_records(
            self._sheet.records.values(),
            self.sort,
            count_sort=self.sort_by_comment_count,
        )
