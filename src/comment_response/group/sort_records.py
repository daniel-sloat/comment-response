"""Grouping for datasheet records."""


class SortRecords:
    """Provides sort data information. Two sort columns are used:
    (1) one for alphabetic sorting, and
    (2) another that can be used for custom sorting."""

    def __init__(self, config: dict):
        self.title = config["title"]
        self.ordered = config["ordered"]
        self.by_count = config["by_count"]

    def key(self):
        """Sorting tuple for grouping records."""
        return tuple(zip(self.ordered, self.title))
