from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from lxml.etree import _ElementTree

from xlsx.ooxml_ns import ns


class XLSXXML:
    def __init__(self, filename: str):
        self.file = filename
        with ZipFile(self.file, "r") as xlsx:
            self.data = {
                filename: etree.fromstring(xlsx.read(filename))
                for filename in xlsx.namelist()
                if ".xml" in filename
            }

    @property
    def filelist(self) -> list[str]:
        return sorted(list(self.data))

    @property
    def workbook(self) -> _ElementTree:
        return self.data["xl/workbook.xml"]

    @property
    def rels(self) -> _ElementTree:
        return self.data["xl/_rels/workbook.xml.rels"]

    @property
    def sharedstrings(self) -> _ElementTree:
        return self.data["xl/sharedStrings.xml"]

    def sheet(self, sheetname: str) -> _ElementTree:
        rid_xpath = "string(w:sheets/w:sheet[@name=$_name]/@r:id)"
        rid = self.workbook.xpath(rid_xpath, _name=sheetname, **ns)
        if rid:
            filename_xpath = "concat('xl/',r1:Relationship[@Id=$_rid]/@Target)"
            filename = self.rels.xpath(filename_xpath, _rid=rid, **ns)
            return self.data[filename]
        raise ValueError("Sheetname not found.")
