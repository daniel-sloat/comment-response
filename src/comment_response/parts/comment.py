"""Prepare comments"""

import re
from itertools import groupby

from xlsx_rich_text.cell.richtext import RichText
from xlsx_rich_text.cell.run import Run
from xlsx_rich_text.sheets.record import Record

from comment_response.parts.paragraph import Paragraph


class Comment:
    """Prepare comment for writing to docx."""

    def __init__(
        self, record: Record, column: str, tag_column: str, clean_config: dict
    ):
        self.record = record
        self.column: str = column
        self.tag_column: str = tag_column
        self.clean_config = clean_config

    def __bool__(self):
        """Returns true if comment has text or has a tag."""
        text = False
        if self._rich_text:
            text = bool(self._rich_text.text)
        return text or bool(self.tag)

    @property
    def _rich_text(self) -> RichText | None:
        try:
            text = self.record.col.get(self.column).value
            if text:
                return text
        except AttributeError as exc:
            raise ValueError(f"Column name '{self.column}' not found.") from exc

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
        try:
            text = self.record.col.get(self.tag_column).value
            if text:
                return str(text)
        except AttributeError as exc:
            raise ValueError(f"Column name '{self.tag_column}' not found.") from exc

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Group comment runs into paragraphs."""
        paras = []
        keyfunc = lambda run: run.text != "\n"
        for key, runs in groupby(self.runs, key=keyfunc):
            if key:
                paras.append(Paragraph(list(runs), **self.clean_config))
        return paras
