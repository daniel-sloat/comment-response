"""Sheets"""

from typing import TYPE_CHECKING

from lxml.etree import _ElementTree

from xlsx.ooxml_ns import ns
from xlsx.sheets.datasheet import DataSheet
from xlsx.sheets.sheet import Sheet

if TYPE_CHECKING:
    from xlsx.workbook import Workbook


class Sheets:
    """Representation of worksheets as basic (unstructured) sheet."""

    def __init__(self, workbook: "Workbook"):
        self._book = workbook
        self._book_xml: _ElementTree = self._book.xml["xl/workbook.xml"]
        self._book_rels: _ElementTree = self._book.xml["xl/_rels/workbook.xml.rels"]
        self.sheet_names: list = self._book_xml.xpath("w:sheets/w:sheet/@name", **ns)
        self.sheets = {_name: Sheet(_name, self) for _name in self.sheet_names}

    def __repr__(self):
        return f"{self.__class__.__name__}(names={tuple(self.sheets.keys())})"

    def __getitem__(self, key):
        return self.sheets[key]

    def __iter__(self):
        return iter(self.sheets.items())

    def __len__(self):
        return len(self.sheets)


class DataSheets(Sheets):
    """Representation of worksheets as table (structured), with header."""

    def __init__(self, workbook: "Workbook"):
        super().__init__(workbook)
        self.sheets = {_name: DataSheet(_name, self) for _name in self.sheet_names}

    def __getitem__(self, key):
        return self.sheets[key]
