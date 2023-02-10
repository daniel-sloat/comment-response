"""Prepare comments"""

import re
from itertools import groupby

from comment_response.paragraph import Paragraph, Paragraphs
from xlsx.cell.run import Run
from xlsx.sheets.record import Record


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
        _feed = (
            Run(txt, run.props)
            for run in self.runs
            for txt in re.split("(\n)", re.sub(r"\s+", " ", run.text))
            if txt
        )
        _paragraphs = [
            Paragraph(run_group)
            for key, run_group in groupby(_feed, key=lambda run: run.text != "\n")
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
