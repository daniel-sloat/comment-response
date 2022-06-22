from logtools import logtools
from config_loader import load_toml_config
from open_office_xml import Sheet
import group
from docx_tools import write_docx, automark


def main():
    logtools.initialize_logging()

    config_file = load_toml_config.load_toml_config()

    print("Reading sheet...")
    logtools.logging.info("Reading sheet...")
    sheet = Sheet.Sheet(
        filepath=config_file["filename"],
        sheetname=config_file["sheetname"],
        header_row=config_file["other"]["header_row"],
    )

    data = group.relevant_data(
        sheet, config_file["columns"], config_file["other"]["clean"]
    )
    data = group.append_comment_tags(data)
    comment_response_data = group.group_data(data, config_file["sort"])

    write_docx.commentsectiondoc(
        comment_response_data,
        outline_level_start=config_file["other"]["outline_level"],
    )

    entry_list = automark.make_entry_list(data)
    automark.automarkdoc(entry_list)

    logtools.quit_logging()


if __name__ == "__main__":
    main()
