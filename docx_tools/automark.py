import docx

from logtools import logtools


def make_entry_list(groups):
    unique_tags = set()
    for group in groups:
        unique_tags.add(group["tag"])
    unique_tags = sorted(unique_tags)
    entry_list = list(zip(unique_tags, unique_tags))
    return entry_list


@logtools.log_automark
def automarkdoc(
    entry_list: list[list[str, str]],
    savename: str = "output\AutoMark.docx",
) -> str:
    # AutoMark document is document with two col table for automatically
    # marking index entries in another document.

    def _write_table(doc, entry_list):
        # Use table._cells to "pop" out the cells from the table, limiting
        # the amount of calls to the table in the Word document (improving
        # speed by multiple times). Updates Word document only after the
        # table is filled.
        # https://theprogrammingexpert.com/write-table-fast-python-docx/
        table = doc.add_table(rows=len(entry_list), cols=2)
        table_cells = table._cells
        for i in range(len(entry_list)):
            for j in range(len(entry_list[i])):
                table_cells[j + i * 2].text = str(entry_list[i][j])

    doc = docx.Document()
    _write_table(doc, entry_list)
    doc.save(savename)
    return savename
