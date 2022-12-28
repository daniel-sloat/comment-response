from xlsx.ooxml_ns import ns
from xlsx.styles.style import Style


class Styles:
    """Representation of styles.xml"""

    def __init__(self, book):
        self._book = book
        self._styles_xml = self._book.xml["xl/styles.xml"]
        self.styles = [
            Style(xfs, num, self) for num, xfs in enumerate(self._styles_xml.xpath("w:cellXfs/w:xf", **ns))
        ]

    def __repr__(self):
        return f"{self.__class__.__name__}(file='{self._book.file.name}')"

    def __getitem__(self, key):
        return self.styles[key]

    def __iter__(self):
        return iter(self.styles)

    def __len__(self):
        return len(self.styles)
