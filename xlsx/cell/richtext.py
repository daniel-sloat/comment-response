"""Handles rich text formatting in cell."""

import re
from itertools import groupby
from reprlib import Repr

from lxml.etree import _Element

from xlsx.cell.paragraph import Paragraph, Paragraphs
from xlsx.cell.run import Run
from xlsx.ooxml_ns import ns
# from xlsx.workbook import Workbook


class RichText:
    """Provides rich text formatting for cell."""

    def __init__(self, element: _Element, book):
        self.element = element
        self._book = book
        self.runs = [
            Run.from_element(el) for el in self.element.xpath("w:t|w:r/w:t", **ns)
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
    def text(self) -> str:
        return "".join(run.text for run in self.runs)

    @property
    def paragraphs(self) -> Paragraphs:
        _feed = (
            Run(txt, run.props)
            for run in self.runs
            for txt in re.split("(\n)", run.text)
            if txt
        )
        _paragraphs = [
            Paragraph(run_group)
            for key, run_group in groupby(_feed, key=lambda run: run.text != "\n")
            if key
        ]
        return Paragraphs(_paragraphs)
