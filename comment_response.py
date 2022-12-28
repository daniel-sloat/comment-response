from pprint import pprint

from config import toml_config
from xlsx.workbook import Workbook


def main():
    config = toml_config.load()
    book = Workbook(config["filename"])
    sheet = book.datasheets[config["sheetname"]]
    for row in sheet:
        for cell in row:
            pprint(cell[1].style.props)
    # pprint(sheet[0]["CommentText"].style.props, sort_dicts=False)
    # Comments-response extracted
    # Write comments-section doc
    # Write automark doc


if __name__ == "__main__":
    main()
