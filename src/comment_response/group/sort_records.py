"""Grouping for datasheet records."""


from dataclasses import dataclass
from itertools import zip_longest


class SortRecords:
    """Provides sort data information. Two sort columns are used:
    (1) one for alphabetic sorting, and
    (2) another that can be used for custom sorting."""

    def __init__(self, config: dict):
        self.title = config["title"]
        self.ordered = config["ordered"]
        self.by_count = config["by_count"]

        if len(self.title) < len(self.ordered):
            raise ValueError(
                "The number of order columns must be less "
                "than the number of title columns."
            )

    def key(self):
        """Sorting tuple for grouping records."""
        return tuple(zip_longest(self.ordered, self.title, fillvalue=""))


@dataclass(frozen=True, order=True)
class Heading:
    """For column sorting. Displays textual information in repr, but stores custom sort
    information as descriptor."""

    num: int
    title: str

    def __bool__(self):
        return any((bool(self.num), bool(self.title)))

    def __repr__(self):
        return self.title