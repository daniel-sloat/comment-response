"""Main script"""

import tomllib
from comment_response import Section
from xlsx_rich_text import Workbook

# Final steps:
# - Implement remaining config (remove_all_double_spaces)
# - Implement tagging system for comments
# - Fix automark to use tagging system
# - Implement logging
# - Fix typing issues


def main():
    with open("config.toml", "rb") as f:
        config = tomllib.load(f)
    book = Workbook(config["filename"])
    sheet = book.sheet(config["sheetname"], header_row=config["other"]["header_row"])
    section = Section(sheet, **config)
    print(section.group_records)


if __name__ == "__main__":
    main()
