from dataclasses import dataclass


@dataclass(frozen=True, order=True)
class ColSort:
    """For column sorting. Displays textual information in repr, but stores custom sort
    information as descriptor."""

    num: int
    title: str
