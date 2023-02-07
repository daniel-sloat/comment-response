"""Main script"""

from comment_response.write_docx import CommentSection
from config import toml_config
from xlsx.workbook import Workbook

# Final steps:
# - Implement remaining config
# - Remove any leading or trailing spaces from end of paragraphs
# - Implement tagging system for comments
# - Fix automark to use tagging system
# - Implement logging


def main():
    config = toml_config.load()
    book = Workbook(config["filename"])
    sheet = book.datasheets[config["sheetname"]]
    comment_section = CommentSection(sheet, **config)
    comment_section.write()


if __name__ == "__main__":
    main()
