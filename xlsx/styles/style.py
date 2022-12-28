from xlsx.helpers.attrib import get_attrib
from xlsx.ooxml_ns import ns
from xlsx.styles.implied_numfmts import numfmts


class Style:
    pieces = (
        ("numFmts", "numFmt", "numFmtId", "NumberFormat"),
        ("fonts", "font", "fontId", "Font"),
        ("fills", "fill", "fillId", "Fill"),
        ("borders", "border", "borderId", "Border"),
    )

    def __init__(self, element, _id, styles):
        self.element = element
        self._id = _id
        self._styles = styles

    def __repr__(self):
        return f"{self.__class__.__name__}(_id={self._id})"

    def __getitem__(self, key):
        return self.props[key]

    def __iter__(self):
        return iter(self.props)

    def __len__(self):
        return len(self.props)

    @property
    def props(self):
        _props = {}
        for parent, child, _id_el, apply in self.pieces:
            _id = self.element.xpath(f"number(@{_id_el}) + 1", **ns)

            apply_style = self.element.xpath(f"boolean(@apply{apply})", **ns)
            if apply_style:
                style_id = self.element.xpath("number(@xfId) + 1", **ns)
                xpath = f"number(w:cellStyleXfs/w:xf[position()=$_id]/@{_id_el}) + 1"
                new_id = self._styles._styles_xml.xpath(xpath, _id=style_id, **ns)
                if new_id:
                    _id = new_id

            if parent == "numFmts":
                if int(_id) in numfmts:
                    _props |= {"numFmt": numfmts[int(_id)]}

            _props |= get_attrib(
                self._styles._styles_xml.xpath(
                    f"w:{parent}/w:{child}[position()=$_id]", _id=_id, **ns
                )[0]
            )
        return _props

    @property
    def font(self):
        _id = self.element.xpath(f"number(@fontId) + 1", **ns)
        apply = (
            self.element.xpath(f"boolean(@applyFont)", **ns) and
            self.element.xpath(f"@applyFont='1' or @applyFont='true' or @applyFont='on'", **ns)
        )
        _props = get_attrib(
            self._styles._styles_xml.xpath(
                f"w:fonts/w:font[position()=$_id]", _id=_id, **ns
            )[0]
        )
        return _props

    # @property
    # def props(self):
    #     """Apply properties if flag is set, and lookup property information in xml. Return dictionary."""
    #     _props = {}
    #     if self.element.xpath("number(@xfId) > 0", **ns):
    #         pass
    #     for parent, child, _id_el, apply in self.pieces:
    #         _id = self.element.xpath(f"number(@{_id_el}) + 1", **ns)
    #         _apply = self.element.xpath(
    #             f"@apply{apply}='1' or @apply{apply}='true' or @apply{apply}='on'", **ns
    #         )
    #         _do_not_apply = self.element.xpath(
    #             f"@apply{apply}='0' or @apply{apply}='false' or @apply{apply}='off'",
    #             **ns,
    #         )

    #     return _props
