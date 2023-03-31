"""Prepare response"""

from xlsx_rich_text.cell.richtext import RichText
from xlsx_rich_text.sheets.record import Record

from comment_response.parts.paragraph import Paragraph


class Response:
    """Prepare response for writing to docx."""

    def __init__(self, records: list[Record], response_col: str, clean_config: dict):
        self.column = response_col
        self.records = [record for record in records if record.col.get(self.column)]
        self.clean_config = clean_config

    @property
    def paragraphs(self) -> list[Paragraph]:
        paras = []
        for record in self.records:
            cell = record[self.column]
            rich_text: RichText = cell.value
            if rich_text:
                for paragraph in rich_text.paragraphs:
                    paras.append(Paragraph(paragraph.runs, **self.clean_config))
        return paras
