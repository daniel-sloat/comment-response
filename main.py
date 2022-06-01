from lxml import etree
import tomli

from pathlib import Path
import logging
from zipfile import ZipFile
import re
from itertools import groupby
from pprint import pprint

from docx_tools import write_docx

# Create dict of relevant xml content
# Use current formulas to create comment record
# See how to group data using itertools


def initialize_logging():
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s",
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
    # header_data = header_data[row_no]
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
    data_rows = []
    for row in rows:
        row_data = get_row_data(
            row, plain_strings, rich_strings, comment_col_num, response_col_num
        )
        data_rows.append(row_data)
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
    # assert row_no == int(row_num), "Mixup error."
    # row_data[row_no] = column_data
    return column_data


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
                # //TODO This needs to not rely on color and rgb keys.
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


def find_column_num(header, config_file):
    for k, v in header.items():
        if v == config_file["columns"]["comment"]:
            comment_col_num = k
        if v == config_file["columns"]["response"]:
            response_col_num = k
    return comment_col_num, response_col_num


def find_columns(header, config_file):
    columns = {}
    for k, v in header.items():
        columns[k] = {}
        for k1, v1 in config_file["columns"].items():
            if v == v1:
                columns[k] = {k1: v}
        for heading_no, column in enumerate(config_file["sort"]["columns"], 1):
            if v == column:
                columns[k] = {f"heading{heading_no}": column}
    return columns


def create_new_rows(data, columns_relevant):
    new_rows = []
    for row in data:
        d = {}
        cols = []
        for col_no, col_data in columns_relevant.items():
            for col_name, value in col_data.items():
                if "heading" in col_name:
                    cols.append(row.get(col_no))
                else:
                    d[col_name] = row.get(col_no)
        d["sort"] = cols
        new_rows.append(d)
    return new_rows


def group_comments_and_responses(group):
    # Combine comments into list[list] (paragraphs still denoted by '\n')
    # Some response cells will be empty. In case there are multiple cells of
    # responses, combine into list as well (unfortunately the responses won't
    # be in any sort of order).
    grouped_comments = []
    grouped_responses = []
    comments_and_responses = {}
    comments_and_responses["comment_data"] = {}
    for g in group:
        # Selects only those rows with comments.
        if g["comment"]:
            grouped_comments.append(g["comment"])
            # Only select responses attached to a comment
            if g["response"]:
                grouped_responses.append(g["response"])
    comments_and_responses["comment_data"]["comments"] = grouped_comments
    comments_and_responses["comment_data"]["response"] = grouped_responses
    return comments_and_responses


def group_data(comment_response_data):

    def initial_sort_and_group(comment_response_data, key_sort):
        comment_response_data = sorted(comment_response_data, key=key_sort)
        initial_grouping = []
        for key, group in groupby(comment_response_data, key=key_sort):
            grouped_data = group_comments_and_responses(group)
            initial_grouping.append(
                {"sort": key[:-1], "heading": key[-1], "data": grouped_data},
            )
        return initial_grouping

    def following_groupings(grouped_data, key_sort):
        new_grouped_data = []
        for key, group in groupby(grouped_data, key=key_sort):
            combined = []
            for g in group:
                combined.append(g)
                g.pop("sort")
            if key[:-1]:
                new_grouped_data.append(
                    {"sort": key[:-1], "heading": key[-1], "data": combined}
                )
            else:
                new_grouped_data.append({"heading": key[-1], "data": combined})
        return new_grouped_data

    key_sort = lambda x: tuple(x["sort"])
    combo_list = initial_sort_and_group(comment_response_data, key_sort)
    while combo_list[0].get("sort"):
        combo_list = following_groupings(combo_list, key_sort)
    return combo_list


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

    comment_col_num, response_col_num = find_column_num(header, config_file)

    columns_relevant = find_columns(header, config_file)

    data = get_data_after_header(
        xlsx_tree,
        sheet_rels[config_file["sheetname"]],
        plain_strings,
        rich_strings,
        config_file["other"]["header_row"],
        comment_col_num,
        response_col_num,
    )

    new_rows = create_new_rows(data, columns_relevant)
    new_rows = group_data(new_rows)

    #pprint(new_rows[0]["data"][1]["data"], sort_dicts=False, width=150)

    write_docx.commentsectiondoc(new_rows)
    
    quit_logging()


if __name__ == "__main__":
    main()
