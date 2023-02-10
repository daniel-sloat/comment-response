"""Main script"""

from comment_response.write_docx import CommentSection
from config import toml_config
from xlsx.workbook import Workbook

# Final steps:
# - Fix formatting
# - Implement remaining config
# - Implement tagging system for comments
# - Fix automark to use tagging system
# - Implement logging


def main():
    config = toml_config.load()
    book = Workbook(config["filename"])
    # print(book.datasheet(config["sheetname"], header_row=1))
    sheet = book.datasheets[config["sheetname"]]
    CommentSection(sheet, **config).write()


if __name__ == "__main__":
    main()
