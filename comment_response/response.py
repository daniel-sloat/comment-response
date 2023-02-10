"""Prepare response"""

from comment_response.paragraph import Paragraph, Paragraphs
from xlsx.cell.richtext import RichText
from xlsx.sheets.record import Record


class Response:
    """Prepare response for writing to docx."""

    def __init__(self, records: list[Record], config: str):
        self.records = records
        self.response_col = config["columns"]["commentresponse"]["response"]

    @staticmethod
    def _get_paragraphs(records: list[Record], column: str) -> Paragraph:
        for record in records:
            cell = record.col.get(column)
            if cell:
                rich_text: RichText = cell.value
                if rich_text:
                    for paragraph in rich_text.paragraphs:
                        yield Paragraph(paragraph.runs)

    def prepared(self) -> Paragraphs:
        return Paragraphs(
            [
                paragraph
                for paragraph in self._get_paragraphs(self.records, self.response_col)
            ]
        )
