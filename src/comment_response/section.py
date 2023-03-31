"""Comment section"""

from pathlib import Path

import docx
from docx.document import Document
from xlsx_rich_text.sheets.newdatasheet import NewDataSheet

from comment_response.group.recursive_group import group_records
from comment_response.group.sort_records import SortRecords
from comment_response.logger.logger import log, log_write
from comment_response.write.automark import AutoMark
from comment_response.write.docx import recursive_write
from comment_response.write.styles import create_style


@log()
class Section:
    """Write comment-response section to docx."""

    def __init__(self, sheet: NewDataSheet, **config):
        self.sheet = sheet
        self.sheetname: str = sheet.sheetname
        self.config: dict = config
        self.sort = SortRecords(config["sort"])

    @property
    def records(self):
        return list(self.sheet.records.values())

    def section_data(self):
        return group_records(self.records, self.sort.key(), self.sort.by_count)

    @log_write
    def write(self, filename: str = "output/section.docx", outline_level: int = 1):
        doc: Document = docx.Document()
        path = Path(filename)
        path.parent.mkdir(exist_ok=True)

        create_style(doc, "Comments")
        create_style(doc, "Response", left_indent=0.5, next_style="Response")
        recursive_write(doc, self.section_data(), self.config, outline_level)

        doc.save(path)

    @property
    def automark(self):
        return AutoMark(self.records)
