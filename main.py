"""Main script"""

import tomllib
from xlsx_rich_text import Workbook

from comment_response import Section


def main():
    with open("config.toml", "rb") as toml:
        config = tomllib.load(toml)

    book = Workbook(config["filename"])
    sheet = book.sheet(config["sheetname"], header_row=config["header_row"])

    section = Section(sheet, **config["section"])
    section.write(config["savename"], config["outline_level"])
    section.automark.write(config["automark"])


if __name__ == "__main__":
    main()
