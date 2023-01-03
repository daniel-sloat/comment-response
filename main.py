from pprint import pprint

from config import toml_config
from xlsx.workbook import Workbook
from comment_response.data_prep import PrepData


def main():
    config = toml_config.load()
    book = Workbook(config["filename"])
    sheet = book.datasheets[config["sheetname"]]
    data = PrepData(sheet, **config)
    grouped = data.grouped_records()
    pprint(list(grouped), sort_dicts=False, width=100)

    # Comments-response extracted
    # Write comments-section doc
    # Write automark doc


if __name__ == "__main__":
    main()
