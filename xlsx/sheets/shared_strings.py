from xlsx.ooxml_ns import ns
from xlsx.cell.richtext import RichText


class SharedStrings:
    def __init__(self, book):
        self._book = book
        self._xml = self._book.xml["xl/sharedStrings.xml"]
        self.strings = [
            RichText(el, self._book) for el in self._xml.xpath("w:si", **ns)
        ]

    def __repr__(self):
        return f"SharedStrings(count={len(self.strings)})"

    def __getitem__(self, key):
        return self.strings[key]

    def __iter__(self):
        return iter(self.strings)

    def __len__(self):
        return len(self.strings)
