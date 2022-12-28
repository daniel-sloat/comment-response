from reprlib import Repr

from xlsx.helpers.attrib import get_attrib
from xlsx.ooxml_ns import ns


class Run:
    def __init__(self, element, book):
        self._element = element
        self._book = book
        self.text = self._element.xpath("string(.)", **ns)
        self.props = get_attrib(self._element.xpath("parent::r/rPr"))

    def __repr__(self):
        return f"{self.__class__.__name__}(text={Repr().repr(self.text)})"

    def __str__(self):
        return self.text
