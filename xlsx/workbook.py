"""Class Workbook"""

from functools import cached_property
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from xlsx.sheets.shared_strings import SharedStrings
from xlsx.sheets.sheets import DataSheets, Sheets
from xlsx.styles.styles import Styles


# @log_filename
class Workbook:
    """Opens xlsx workbook and creates XML file tree"""

    def __init__(self, filename):
        self.file = Path(filename)
        self.styles = Styles(self)
        self.sharedstrings = SharedStrings(self)
        self.sheets = Sheets(self)
        self.datasheets = DataSheets(self)

    def __repr__(self):
        return f"Workbook(file='{self.file}')"

    @cached_property
    def xml(self):
        with ZipFile(self.file.absolute(), "r") as z:
            return {
                filename: etree.fromstring(z.read(filename))
                for filename in z.namelist()
                if ".xml" in filename
            }
