from lxml import etree

from zipfile import ZipFile
import re

from main import load_toml_config, get_file


NAMESPACES = {
        "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/package/2006/relationships"
        }


def get_xlsx_xml_tree(
    xlsx_path: str,
) -> dict[str : etree.Element]:
    """Gets dictionary of relevant Office Open XML root element nodes in xlsx.

    Args:
        xlsx_path (str): Document location path.

    Returns:
        dict[str:etree.Element]: Returns dict with zipped filepath as keys
            and values of root etree element.
    """
    with ZipFile(xlsx_path, "r") as z:
        xlsx_xml_tree = {}
        regex = r"^xl/(?:workbook|sharedStrings|worksheets/sheet\d*)\.xml$"
        for xml_file in [name for name in z.namelist() if re.search(regex, name)]:
            xlsx_xml_tree[xml_file] = etree.fromstring(z.read(xml_file))
    return xlsx_xml_tree


def get_sheet_names(xlsx_xml_tree):
    sheet_names = xlsx_xml_tree["xl/workbook.xml"].xpath("sheets/sheet/@name", namespaces=NAMESPACES)
    
    
    return sheet_names


def main():
    config_file = load_toml_config()
    file = get_file(config_file["FILENAME"])
    xlsx_tree = get_xlsx_xml_tree(file)
    sheet_names = get_sheet_names(xlsx_tree)
    print(sheet_names)


if __name__ == "__main__":
    main()