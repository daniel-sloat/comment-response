"""Prepare comments"""

import re
from itertools import groupby

from xlsx_rich_text.cell.run import Run
from xlsx_rich_text.sheets.record import Record

from comment_response.paragraph import Paragraph, Paragraphs


class Comment:
    """Prepare comment for writing to docx."""

    def __init__(self, record: Record, config: dict):
        self.record = record
        self.config = config
        self._tag_col: str = config["columns"]["other"]["tag"]
        self.comment_col: str = config["columns"]["commentresponse"]["comment"]

    @property
    def value(self):
        cell = self.record.col.get(self.comment_col).value
        if cell:
            return cell

    @property
    def runs(self) -> list[Run]:
        if self.value:
            return self.value.runs
        else:
            return []

    @property
    def tag(self) -> str:
        cell = self.record.col.get(self._tag_col)
        if cell:
            return str(cell.value)

    @property
    def paragraphs(self) -> Paragraphs:
        def gen(runs: list[Run]):
            for run in runs:
                for txt in re.split("(\n)", re.sub(r"[^\S\n]+", " ", run.text)):
                    if txt:
                        yield Run(txt, run.props)
            yield Run(f" ({self.tag})", None)

        _paragraphs = [
            Paragraph(run_group)
            for key, run_group in groupby(
                gen(self.runs), key=lambda run: run.text != "\n"
            )
            if key
        ]
        return Paragraphs(_paragraphs)


class Comments:
    """Prepare comments for writing to docx."""

    def __init__(self, records: list[Record], config: dict):
        self.records = records
        self.config = config
        self.comment_col: str = config["columns"]["commentresponse"]["comment"]

    def prepared(self) -> list[Comment]:
        return [
            Comment(record, self.config).paragraphs
            for record in self.records
            if record.col.get(self.comment_col).value
        ]
