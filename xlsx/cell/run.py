"""Text formatting run for rich text formatting."""

from reprlib import Repr

from lxml.etree import _Element

from xlsx.helpers.attrib import get_attrib
from xlsx.ooxml_ns import ns

run_repr = Repr()
run_repr.maxdict = 2
run_repr.maxlevel = 1


class Run:
    """Rich text run formatting."""

    def __init__(self, text, props):
        self.text = text
        self.props = props

    def __repr__(self):
        return (
            f"{self.__class__.__name__}("
            f"text={Repr().repr(self.text)}, "
            f"props={run_repr.repr(self.props)})"
        )

    def __str__(self):
        return self.text

    @classmethod
    def from_element(cls, element: _Element):
        text = element.xpath("string(.)", **ns)
        _props = element.xpath("parent::w:r/w:rPr", **ns)
        if len(_props):
            props = get_attrib(_props[0])["rPr"]
        else:
            props = None
        return cls(text, props)
