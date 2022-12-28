from xlsx.cell.cell import Cell
from xlsx.ooxml_ns import ns


class Sheet:
    """Representation of sheet.xml"""

    def __init__(self, _name, sheets):
        self._name = _name
        self._parent = sheets

    def __repr__(self):
        return f"{self.__class__.__name__}(name='{self._name}')"

    @property
    def _id(self):
        xpath = "string(w:sheets/w:sheet[@name=$_name]/@sheetId)"
        return self._parent._book_xml.xpath(xpath, _name=self._name, **ns)

    @property
    def xml(self):
        rid_xpath = "string(w:sheets/w:sheet[@name=$_name]/@r:id)"
        rid = self._parent._book_xml.xpath(rid_xpath, _name=self._name, **ns)
        filename_xpath = "string(r1:Relationship[@Id=$_rid]/@Target)"
        filename = "xl/" + self._parent._book_rels.xpath(filename_xpath, _rid=rid, **ns)
        return self._parent._book.xml[filename]

    def cell(self, row, col):
        xpath = "w:sheetData/w:row[@r=$_r]/w:c[@r=$_c]"
        if len(cell := self.xml.xpath(xpath, _r=row, _c=col, **ns)):
            return Cell(cell[0], self)
        return None

    def row(self, row):
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        return [Cell(el, self) for el in self.xml.xpath(xpath, _r=row, **ns)]

    def col(self, col):
        xpath = "w:sheetData/w:row/w:c[@r=$_c]"
        return [Cell(el, self) for el in self.xml.xpath(xpath, _c=col, **ns)]
