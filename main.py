"""Main script"""

import tomllib
from xlsx_rich_text import Workbook

from comment_response import Section

# Final steps:
# - Implement remaining config (remove_all_double_spaces)
# - Implement logging
# - Fix typing issues


def main():
    with open("config.toml", "rb") as toml:
        config = tomllib.load(toml)
    book = Workbook(config["filename"])
    sheet = book.sheet(config["sheetname"], header_row=config["other"]["header_row"])
    section = Section(sheet, **config)
    section.write()
    section.automark.write()


if __name__ == "__main__":
    main()
