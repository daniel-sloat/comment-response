"""Class Workbook"""

from functools import cached_property
from zipfile import ZipFile

from lxml import etree
from lxml.etree import _ElementTree

from xlsx.ooxml_ns import ns
from xlsx.sheets.newdatasheet import NewDataSheet, NewSheet
from xlsx.sheets.shared_strings import SharedStrings
from xlsx.sheets.sheets import DataSheets, Sheets
from xlsx.styles.styles import Styles
from xlsx.xml import XLSXXML


# @log_filename
class Workbook:
    """Opens xlsx workbook and creates XML file tree"""

    def __init__(self, filename: str):
        self.file = filename
        self.xlsx = XLSXXML(self.file)
        self.styles = Styles(self)
        self.sharedstrings = SharedStrings(self)
        self.sheets = Sheets(self)
        self.datasheets = DataSheets(self)

    def __repr__(self):
        return f"Workbook(file='{self.file}')"

    @property
    def sheetnames(self) -> list[str]:
        return self.xlsx.workbook.xpath("w:sheets/w:sheet/@name", **ns)

    @cached_property
    def xml(self) -> dict[str, _ElementTree]:
        with ZipFile(self.file, "r") as xlsx:
            return {
                filename: etree.fromstring(xlsx.read(filename))
                for filename in xlsx.namelist()
                if ".xml" in filename
            }

    def sheet(self, sheetname, header_row=None) -> NewSheet | NewDataSheet:
        if header_row:
            return NewDataSheet(sheetname, self, header_row)
        else:
            return NewSheet(sheetname, self)
