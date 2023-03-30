"""Prepare comments"""

import re
from dataclasses import dataclass
from itertools import groupby

from xlsx_rich_text.cell.richtext import RichText
from xlsx_rich_text.cell.run import Run
from xlsx_rich_text.sheets.record import Record

from comment_response.parts.paragraph import Paragraph


class Comment:
    """Prepare comment for writing to docx."""

    def __init__(self, record: Record, config: dict):
        self.record = record
        self.config = config
        self._tag_col: str = config["columns"]["other"]["tag"]
        self.comment_col: str = config["columns"]["commentresponse"]["comment"]

    def __bool__(self):
        """Returns true if comment has text or has a tag."""
        text = False
        if self._rich_text:
            text = bool(self._rich_text.text)
        return text or bool(self.tag)

    @property
    def _rich_text(self) -> RichText | None:
        cell = self.record.col.get(self.comment_col).value
        if cell:
            return cell

    @property
    def runs(self) -> list[Run]:
        if self._rich_text:
            for run in self._rich_text.runs:
                run_pieces = re.split("(\n)", run.text)
                for txt in run_pieces:
                    if txt:
                        yield Run(txt, run.props)
        else:
            return []

    @property
    def tag(self) -> str:
        cell = self.record.col.get(self._tag_col)
        if cell:
            if cell.value:
                return str(cell.value)

    def paragraphs(self, trim: bool = True, clean: bool = True) -> list[Paragraph]:
        """Group comment runs into paragraphs."""
        paras = []
        keyfunc = lambda run: run.text != "\n"
        for key, runs in groupby(self.runs, key=keyfunc):
            if key:
                paras.append(Paragraph(list(runs), trim=trim, clean=clean))
        return paras
