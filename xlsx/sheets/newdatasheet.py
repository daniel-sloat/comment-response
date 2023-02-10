"""New DataSheet"""

from lxml.etree import _ElementTree

from xlsx.ooxml_ns import ns


class NewSheet:
    """Create Sheet, basic access to worksheet."""

    def __init__(
        self,
        sheetname: str,
        workbook_xml: _ElementTree,
        workbook_rels: _ElementTree,
        sheet_xml: _ElementTree,
    ):
        self.sheetname = sheetname
        self.workbook_xml = workbook_xml
        self.rels = workbook_rels
        self.xml = sheet_xml

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(name='{self.sheetname}')"

    @property
    def _id(self) -> str:
        xpath = "string(w:sheets/w:sheet[@name=$_name]/@sheetId)"
        return self.workbook_xml.xpath(xpath, _name=self.sheetname, **ns)


class NewDataSheet(NewSheet):
    """Create DataSheet, with access for each record by header column."""

    def __init__(
        self,
        sheetname: str,
        workbook_xml: _ElementTree,
        workbook_rels: _ElementTree,
        sheet_xml: _ElementTree,
        header_row: int = 1,
    ):
        super().__init__(sheetname, workbook_xml, workbook_rels, sheet_xml)
        self.header_row = header_row
