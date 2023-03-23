"""Shared Strings"""

from typing import TYPE_CHECKING
from lxml.etree import _ElementTree

from xlsx.cell.richtext import RichText
from xlsx.ooxml_ns import ns

if TYPE_CHECKING:
    from xlsx.workbook import Workbook


class SharedStrings:
    """Access to xl/sharedStrings.xml"""

    def __init__(self, book: "Workbook"):
        self._book = book
        self._xml: _ElementTree = self._book.xml["xl/sharedStrings.xml"]
        self.strings = [RichText(el) for el in self._xml.xpath("w:si", **ns)]

    def __repr__(self):
        return f"SharedStrings(count={len(self.strings)})"

    def __getitem__(self, key):
        return self.strings[key]

    def __iter__(self):
        return iter(self.strings)

    def __len__(self):
        return len(self.strings)

    def index(self, key: int | str) -> RichText:
        rich_text = self._xml.xpath("w:si[$_index]", _index=int(key), **ns)[0]
        return RichText(rich_text)
