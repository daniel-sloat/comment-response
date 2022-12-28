from reprlib import Repr

from xlsx.ooxml_ns import ns
from xlsx.cell.run import Run


class RichText:
    def __init__(self, element, book):
        self.element = element
        self._book = book
        self.runs = [
            Run(el, self._book) for el in self.element.xpath("w:t|w:r/w:t", **ns)
        ]
        
    def __repr__(self):
        return f"RichText({Repr().repr(self.text)})"

    def __getitem__(self, key):
        return self.runs[key]

    def __iter__(self):
        return iter(self.runs)

    def __len__(self):
        return len(self.runs)

    def __str__(self):
        return self.text

    @property
    def text(self):
        return "".join(run.text for run in self.runs)
