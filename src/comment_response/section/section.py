import docx
from docx.document import Document

from pathlib import Path
from comment_response.group.group_records import GroupRecords
from comment_response.section.write_docx import recursive_write, style_maker


class Section:
    """Write comment-response section to docx."""

    def __init__(self, sheet, **config):
        self._sheet = sheet
        self.config: dict = config
        self.outline_level: int = self.config["doc"]["custom"]["outline_level_start"]
        self.group_records = GroupRecords(self._sheet, **config)
        self.filename = Path(self.config["doc"]["savename"])

    def __repr__(self):
        return f"{self.__class__.__name__}(sheetname={self._sheet._name})"

    def write(self):
        doc: Document = docx.Document()
        style_maker(doc, "Comments")
        style_maker(doc, "Response", left_indent=0.5, next_style="Response")
        recursive_write(
            doc, self.group_records.group(), self.config, self.outline_level
        )
        self.filename.parent.mkdir(exist_ok=True)
        doc.save(self.filename)
