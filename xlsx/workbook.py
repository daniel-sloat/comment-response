"""Class Workbook"""

from functools import cached_property
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from lxml.etree import _ElementTree
from xlsx.sheets.newdatasheet import NewDataSheet

from xlsx.sheets.shared_strings import SharedStrings
from xlsx.sheets.sheets import DataSheets, Sheets
from xlsx.styles.styles import Styles
from xlsx.xml import XLSXXML


# @log_filename
class Workbook:
    """Opens xlsx workbook and creates XML file tree"""

    def __init__(self, filename):
        self.file = filename
        self.xlsx = XLSXXML(self.file)
        self.styles = Styles(self)
        self.sharedstrings = SharedStrings(self)
        self.sheets = Sheets(self)
        self.datasheets = DataSheets(self)

    def __repr__(self):
        return f"Workbook(file='{self.file}')"

    @cached_property
    def xml(self) -> dict[str:_ElementTree]:
        with ZipFile(self.file, "r") as xlsx:
            return {
                filename: etree.fromstring(xlsx.read(filename))
                for filename in xlsx.namelist()
                if ".xml" in filename
            }

    # def datasheet(self, sheetname, header_row):
    #     sheet_xml = self.xml[sheetname]
    #     return NewDataSheet(sheetname, sheet_xml, header_row)
