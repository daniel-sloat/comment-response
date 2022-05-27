from lxml import etree
import tomli

from pathlib import Path
import logging
from zipfile import ZipFile
import re
from itertools import groupby
from pprint import pprint

# Create dict of relevant xml content
# Use current formulas to create comment record
# See how to group data using itertools


def initialize_logging():
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    logging.info("Comment response script initalized.")
    return None


def load_toml_config(
    config_filename: str = "config.toml",
) -> dict:
    with open(config_filename, "rb") as f:
        return tomli.load(f)


def get_file(
    file_path: str,
) -> Path:
    logging.info(f"Reading file: {file_path}")
    p = Path(file_path)
    return p


def quit_logging() -> None:
    logging.info("Finished.")
    logging.shutdown()
    return None


NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "r1": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def get_xlsx_xml_tree(
    xlsx_path: str,
) -> dict[str : etree.Element]:
    """Gets dictionary of Office Open XML root element nodes in xlsx.

    Args:
        xlsx_path (str): Document location path.

    Returns:
        dict[str:etree.Element]: Returns dict with zipped filepath as keys
            and values of root etree element.
    """
    with ZipFile(xlsx_path, "r") as z:
        xlsx_xml_tree = {}
        regex = r".+(?:\.xml|\.rels)$"
        for xml_file in [name for name in z.namelist() if re.search(regex, name)]:
            xlsx_xml_tree[xml_file] = etree.fromstring(z.read(xml_file))
    return xlsx_xml_tree


def get_sheet_names(xlsx_xml_tree):
    sheet_ids = xlsx_xml_tree["xl/workbook.xml"].xpath(
        "w:sheets/w:sheet/@r:id", namespaces=NAMESPACES
    )
    sheet_names = xlsx_xml_tree["xl/workbook.xml"].xpath(
        "w:sheets/w:sheet/@name", namespaces=NAMESPACES
    )
    relationship_targets = xlsx_xml_tree["xl/_rels/workbook.xml.rels"].xpath(
        "r1:Relationship/@Target", namespaces=NAMESPACES
    )
    relationship_ids = xlsx_xml_tree["xl/_rels/workbook.xml.rels"].xpath(
        "r1:Relationship/@Id", namespaces=NAMESPACES
    )
    sheet_data = {k: v for k, v in zip(sheet_names, sheet_ids)}
    rels_data = {k: v for k, v in zip(relationship_ids, relationship_targets)}
    sheet_rels = {
        x: f"xl/{b}"
        for x, y in sheet_data.items()
        for a, b in rels_data.items()
        if y == a
    }
    return sheet_rels


def get_rows(xlsx_tree, sheet, row_no=1):
    rows = xlsx_tree[sheet].xpath(
        f"w:sheetData/w:row[@r>{row_no}]", namespaces=NAMESPACES
    )
    return rows


def excel_col_name_to_number(col_index: str) -> int:
    if not col_index.isalpha():
        raise TypeError
    pow = 1
    col_num = 0
    for letter in reversed(col_index.upper()):
        col_num += (ord(letter) - ord("A") + 1) * pow
        pow *= 26
    return col_num


def get_header(xlsx_tree, sheet, plain_strings, rich_strings, row_no):
    row = xlsx_tree[sheet].xpath(
        f"w:sheetData/w:row[@r={row_no}]", namespaces=NAMESPACES
    )[0]
    header_data = get_row_data(row, plain_strings, rich_strings)
    header_data = header_data[row_no]
    return header_data


def get_data_after_header(
    xlsx_tree,
    sheet,
    plain_strings,
    rich_strings,
    row_no,
    comment_col_num,
    response_col_num,
):
    rows = xlsx_tree[sheet].xpath(
        f"w:sheetData/w:row[@r>{row_no}]", namespaces=NAMESPACES
    )
    data_rows = {}
    for row in rows:
        row_data = get_row_data(
            row, plain_strings, rich_strings, comment_col_num, response_col_num
        )
        data_rows |= row_data
    return data_rows


def get_row_data(
    row,
    plain_strings,
    rich_strings,
    comment_col_num=0,
    response_col_num=0,
):
    row_no = int(row.xpath("string(@r)", namespaces=NAMESPACES))
    row_data = {}
    column_data = {}
    columns = row.xpath("w:c", namespaces=NAMESPACES)
    for column in columns:
        regex = re.compile(r"^(?P<letters>\w*?)(?P<numbers>\d*)$")
        position = column.xpath("string(@r)", namespaces=NAMESPACES)
        col_name = regex.search(position).group("letters")
        row_num = regex.search(position).group("numbers")
        col_num = excel_col_name_to_number(col_name)
        value_type = column.xpath("string(@t)", namespaces=NAMESPACES)
        value = column.xpath("string(w:v)", namespaces=NAMESPACES)
        # http://officeopenxml.com/SScontentOverview.php
        match value_type:
            case "e":
                cell_code = None
            case "b":
                cell_code = bool(value)
            case "inlineStr":
                # //TODO Inline String can be formatted. It is inside
                # of a <is> element, similar to sharedStrings.xml
                cell_code = str(value)
            case "s":
                if col_num == comment_col_num or col_num == response_col_num:
                    cell_code = rich_strings[int(value)]
                else:
                    cell_code = plain_strings[int(value)]
            case "str":
                # //TODO If value is present, get that. If not, display
                # formula in <f>
                cell_code = str(value)
            case "" | _:
                if not value:
                    cell_code = None
                elif value.isdigit():
                    cell_code = int(value)
                else:
                    try:
                        cell_code = float(value)
                    except ValueError:
                        cell_code = str(value)
        column_data[col_num] = cell_code
    assert row_no == int(row_num), "Mixup error."
    row_data[row_no] = column_data
    return row_data


def create_formatdict(run_node: etree.Element) -> dict:
    text_formats = run_node.xpath("./w:rPr/*", namespaces=NAMESPACES)
    formatdict = {}
    for format in text_formats:
        ftag = etree.QName(format).localname
        fattrib = format.attrib
        formatdict[ftag] = fattrib
    return formatdict


def run_format(run_node: etree.Element) -> str:
    formatdict = create_formatdict(run_node)
    formats = []
    for k, v in formatdict.items():
        match k:
            case "b":
                formats.append("b")
            case "i":
                formats.append("i")
            case "u":
                if "val" in v and "double" in v["val"]:
                    formats.append("w")
                else:
                    formats.append("u")
            case "vertAlign":
                if "val" in v:
                    if v["val"] == "superscript":
                        formats.append("x")
                    if v["val"] == "subscript":
                        formats.append("v")
            case "strike":
                # Double-strikethrough does not exist in Excel. Strike + red text
                # is used to denote double-strikethrough
                if "color" in formatdict and "rgb" in formatdict["color"]:
                    if formatdict["color"]["rgb"] == "FFFF0000":
                        formats.append("z")
                else:
                    formats.append("s")
    # Sort string of formats
    formats = "".join(sorted(formats))
    return formats


def get_rich_strings(xlsx_tree):
    sharedstrings_rich = []
    groups = xlsx_tree["xl/sharedStrings.xml"].xpath(
        "/w:sst/w:si", namespaces=NAMESPACES
    )
    for group in groups:
        runs = []
        for run in group:
            text = run.xpath("string(.)", namespaces=NAMESPACES)
            formats = run_format(run)
            runs.append([formats, text])
        sharedstrings_rich.append(runs)
    return sharedstrings_rich


def get_plain_ss(xlsx_tree) -> list:
    sharedstrings_plain = []
    groups = xlsx_tree["xl/sharedStrings.xml"].xpath(
        "/w:sst/w:si", namespaces=NAMESPACES
    )
    for group in groups:
        text = []
        for run in group:
            run_text = run.xpath("string(.//text())", namespaces=NAMESPACES)
            text.append(run_text)
        sharedstrings_plain.append("".join(text))
    return sharedstrings_plain


def find_columns(header, config_file):
    for k, v in header.items():
        if v == config_file["columns"]["comments"]:
            comment_col_num = k
        if v == config_file["columns"]["response"]:
            response_col_num = k
        if v == config_file["sort"]["level_1_column"]:
            level_1 = k
        if v == config_file["sort"]["level_2_column"]:
            level_2 = k
        if v == config_file["sort"]["level_3_column"]:
            level_3 = k
    return comment_col_num, response_col_num, level_1, level_2, level_3


def main():
    initialize_logging()
    config_file = load_toml_config()
    file = get_file(config_file["filename"])
    xlsx_tree = get_xlsx_xml_tree(file)
    sheet_rels = get_sheet_names(xlsx_tree)
    plain_strings = get_plain_ss(xlsx_tree)
    rich_strings = get_rich_strings(xlsx_tree)

    header = get_header(
        xlsx_tree,
        sheet_rels[config_file["sheetname"]],
        plain_strings,
        rich_strings,
        config_file["other"]["header_row"],
    )

    comment_col_num, response_col_num, level_1, level_2, level_3 = find_columns(
        header, config_file
    )

    data = get_data_after_header(
        xlsx_tree,
        sheet_rels[config_file["sheetname"]],
        plain_strings,
        rich_strings,
        config_file["other"]["header_row"],
        comment_col_num,
        response_col_num,
    )

    pprint(data[6][level_1], sort_dicts=False)
    
    quit_logging()


if __name__ == "__main__":
    main()
