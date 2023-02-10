"""Text formatting run for rich text formatting."""

from reprlib import Repr

from xlsx.helpers.attrib import get_attrib
from xlsx.ooxml_ns import ns


class Run:
    """Rich text run formatting."""

    def __init__(self, text, props):
        self.text = text
        self.props = props

    def __repr__(self):
        return (
            f"{self.__class__.__name__}("
            f"text={Repr().repr(self.text)}, "
            f"props={Repr().repr(self.props)})"
        )

    def __str__(self):
        return self.text

    @classmethod
    def from_element(cls, element):
        text = element.xpath("string(.)", **ns)
        _props = element.xpath("parent::w:r/w:rPr", **ns)
        if len(_props):
            props = get_attrib(_props[0])["rPr"]
        else:
            props = None
        return cls(text, props)
