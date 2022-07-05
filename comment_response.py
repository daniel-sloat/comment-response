from logtools import logtools
from config_loader import load_toml_config
from open_office_xml import Sheet
import comment_response.group as group
import comment_response.relevant_data as relevant_data
from docx_tools import write_docx, automark
from win32_tools import mark_index_entries


def main():
    logtools.initialize_logging()

    config_file = load_toml_config.load_toml_config()

    print(f"Reading sheet '{config_file['sheetname']}' from {config_file['filename']}...")
    logtools.logging.info(f"Reading sheet '{config_file['sheetname']}' from {config_file['filename']}...")
    sheet = Sheet.Sheet(
        filepath=config_file["filename"],
        sheetname=config_file["sheetname"],
        header_row=config_file["other"]["header_row"],
    )

    data = relevant_data.comment_data(sheet, config_file)
    comment_response_data = group.group_data(data, config_file["sort"])

    write_docx.commentsectiondoc(
        comment_response_data,
        outline_level_start=config_file["other"]["outline_level"],
    )
    entry_list = automark.make_entry_list(data)
    automark.automarkdoc(entry_list)

    if config_file["index"]["mark_index_entries"]:
        print("Marking index entries using Microsoft Word...")
        logtools.logging.info("Marking index entries using Microsoft Word...")
        mark_index_entries(add_index=config_file["index"]["append_comment_index"])

    logtools.quit_logging()


if __name__ == "__main__":
    main()
