"""Main script"""

from comment_response.write_docx import CommentSection
from config import toml_config
from xlsx.workbook import Workbook


def main():
    config = toml_config.load()
    book = Workbook(config["filename"])
    sheet = book.datasheets[config["sheetname"]]
    comment_section = CommentSection(sheet, **config)
    comment_section.write()


if __name__ == "__main__":
    main()
