"""Main script"""

from comment_response.write_docx import CommentSection
from config import toml_config
from xlsx.workbook import Workbook

# Final steps:
# - Implement remaining config (remove_all_double_spaces, header_row)
# - Implement tagging system for comments
# - Fix automark to use tagging system
# - Implement logging
# - Fix typing issues


def main():
    config = toml_config.load()
    book = Workbook(config["filename"])
    # print(book.datasheet(config["sheetname"], header_row=1))
    sheet = book.datasheets[config["sheetname"]]
    CommentSection(sheet, **config).write()


if __name__ == "__main__":
    main()
