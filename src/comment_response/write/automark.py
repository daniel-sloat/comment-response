"""Create automark doc."""

from pathlib import Path

import docx
from docx.document import Document as _Document
from docx.table import _Cell
from xlsx_rich_text.sheets.record import Record

from comment_response.logger.logger import log_write


class AutoMark:
    """Creates automark table data and writes automark doc."""

    def __init__(self, records: list[Record], config: dict):
        self.records = records
        self.commenter = config["columns"]["commenter"]
        self.comment_tag = config["columns"]["comment_tag"]

    @property
    def entries(self) -> list[tuple[str, str]]:
        """Entries to be written to AutoMark document."""
        entry = {
            (str(record.col[self.comment_tag]), str(record.col[self.commenter]))
            for record in self.records
        }
        return sorted(entry)

    @log_write
    def write(self, filename=r"output\automark.docx") -> None:
        """Write AutoMark document."""
        path = Path(filename)
        path.parent.mkdir(exist_ok=True)

        doc: _Document = docx.Document()
        table = doc.add_table(rows=len(self.entries), cols=2)

        # https://theprogrammingexpert.com/write-table-fast-python-docx/
        table_cells: list[_Cell] = table._cells
        for row_number, row_values in enumerate(self.entries):
            for col_number, col_value in enumerate(row_values):
                position = col_number + row_number * 2
                table_cells[position].text = str(col_value)

        doc.save(filename)
