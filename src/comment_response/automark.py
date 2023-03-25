"""Create automark doc."""

import docx
from docx.document import Document as _Document


class AutoMark:
    """Creates automark table data and writes automark doc."""

    def __init__(self, grouped_records, savefile=r"output\AutoMark.docx"):
        self.grouped_records = grouped_records
        self.savefile = savefile

    @property
    def automark_entries(self):
        unique_tags = set()
        for group in self.grouped_records:
            unique_tags.add(group["tag"])
        unique_tags = sorted(unique_tags)
        entry_list = list(zip(unique_tags, unique_tags))
        return entry_list

    def write(self) -> str:
        # https://theprogrammingexpert.com/write-table-fast-python-docx/
        doc: _Document = docx.Document()
        table = doc.add_table(rows=len(self.automark_entries), cols=2)
        table_cells = table._cells
        for i in range(len(self.automark_entries)):
            for j in range(len(self.automark_entries[i])):
                table_cells[j + i * 2].text = str(self.automark_entries[i][j])
        doc.save(self.savefile)
