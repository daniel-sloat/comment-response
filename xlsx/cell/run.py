"""Text formatting run for rich text formatting."""

from reprlib import Repr

from xlsx.helpers.attrib import get_attrib
from xlsx.ooxml_ns import ns


class Run:
    """Rich text run formatting."""

    def __init__(self, element, book):
        self._element = element
        self._book = book

    def __repr__(self):
        return f"{self.__class__.__name__}(text={Repr().repr(self.text)})"

    def __str__(self):
        return self.text

    @property
    def text(self):
        return self._element.xpath("string(.)", **ns)

    @property
    def props(self):
        prop = self._element.xpath("parent::w:r/w:rPr", **ns)
        if len(prop):
            return get_attrib(prop[0])
