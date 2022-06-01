import docx
from docx.enum.text import WD_UNDERLINE
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt


def _create_styles(doc: docx.Document) -> None:
    def _style_maker(
        doc: docx.Document,
        name: str,
        base_style: str = "Normal",
        left_indent: int | float = 0,
        space_before: int = 12,
        space_after: int = 12,
        next_style: str = "",
        keep_with_next: bool = False,
    ) -> None:
        styles = doc.styles
        style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles[base_style]
        style.paragraph_format.left_indent = Inches(left_indent)
        style.paragraph_format.space_before = Pt(space_before)
        style.paragraph_format.space_after = Pt(space_after)
        if next_style:
            style.next_paragraph_style = styles[next_style]
        if keep_with_next:
            style.paragraph_format.keep_with_next = keep_with_next
        return None

    _style_maker(doc, "Comments")
    _style_maker(doc, "Response", left_indent=0.5, next_style="Response")


def _word_formats(tag: str, run) -> None:
    for f in tag:
        match f:
            case "b":
                run.font.bold = True
            case "i":
                run.font.italic = True
            case "u":
                run.font.underline = True
            case "w":
                run.font.underline = WD_UNDERLINE.DOUBLE
            case "s":
                run.font.strike = True
            case "z":
                run.font.double_strike = True
            case "x":
                run.font.superscript = True
            case "v":
                run.font.subscript = True
    return None


def _write_comments_and_responses(doc, group_data):
    for comment in group_data["comment_data"]["comments"]:
        paragraph = doc.add_paragraph(style="Comments")
        intro = paragraph.add_run("Comment")
        intro.underline = True
        paragraph.add_run(": ")
        for para_no, para in enumerate(comment):
            if para_no == 0:
                p = paragraph.add_run(para[1])
                _word_formats(para[0], p)
            else:
                # paragraph = doc.add_paragraph()
                p = paragraph.add_run(para[1])
                _word_formats(para[0], p)
    for response in group_data["comment_data"]["response"]:
        paragraph = doc.add_paragraph(style="Response")
        intro = paragraph.add_run("Agency Response")
        intro.underline = True
        intro.bold = True
        paragraph.add_run(": ")
        for para_no, para in enumerate(response):
            if para_no == 0:
                p = paragraph.add_run(para[1])
                _word_formats(para[0], p)
            else:
                # paragraph = doc.add_paragraph()
                p = paragraph.add_run(para[1])
                _word_formats(para[0], p)


def _write_document(doc: docx.Document, top_level: list) -> None:
    def recursion(top_level, outline_level=1):
        for group in top_level:
            if group["heading"]:
                doc.add_heading(group["heading"], outline_level)
            if isinstance(group["data"], dict):
                _write_comments_and_responses(doc, group["data"])
            else:
                recursion(group["data"], outline_level + 1)

    recursion(top_level)
    return None


def commentsectiondoc(
    nested_comment_responses: list,
    savename: str = "output\CommentResponseSection.docx",
) -> None:
    print("Creating Comments and Response section document... ")
    doc = docx.Document()
    _create_styles(doc)
    _write_document(doc, nested_comment_responses)
    doc.save(savename)
    print("Comments and response section document created: " + savename)
    return None


def automarkdoc(entry_list: list, savename: str = "output\AutoMark.docx") -> None:
    # AutoMark document is document with two col table for automatically
    # marking index entries in another document.
    print("Creating AutoMark document... ")

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
    print("AutoMark document created: " + savename)
    return None
