from lxml import etree

import re
from functools import cache

from .dataclasses import Cell, Run
from .helpers import excel_col_name_to_number


class SheetFunctions:
    @property
    def _sheet_relationships(self):
        sheet_ids = self._xml_tree["xl/workbook.xml"].xpath(
            "w:sheets/w:sheet/@r:id", namespaces=self.NAMESPACES
        )
        sheet_names = self._xml_tree["xl/workbook.xml"].xpath(
            "w:sheets/w:sheet/@name", namespaces=self.NAMESPACES
        )
        r_ids = self._xml_tree["xl/_rels/workbook.xml.rels"].xpath(
            "r1:Relationship/@Id", namespaces=self.NAMESPACES
        )
        filenames = self._xml_tree["xl/_rels/workbook.xml.rels"].xpath(
            "r1:Relationship/@Target", namespaces=self.NAMESPACES
        )

        sheet_data = {k: v for k, v in zip(sheet_names, sheet_ids)}
        rels_data = {k: v for k, v in zip(r_ids, filenames)}
        # Filenames do not have "xl/" directory before names
        sheet_rels = {
            x: f"xl/{b}"
            for x, y in sheet_data.items()
            for a, b in rels_data.items()
            if y == a
        }
        return sheet_rels

    @property
    def _sheet_roots(self):
        sheets = {
            x: b
            for x, y in self._sheet_relationships.items()
            for a, b in self._xml_tree.items()
            if y == a
        }
        return sheets

    def _get_cell_position(self, column):
        regex = re.compile(r"^(?P<letters>\w*?)(?P<numbers>\d*)$")
        position = column.xpath("string(@r)", namespaces=self.NAMESPACES)
        col_name, row_num = regex.search(position).groups()
        col_num = excel_col_name_to_number(col_name)
        return col_num, int(row_num)
    
    def _get_cell_position2(self, column):
        regex = re.compile(r"^(?P<letters>\w*?)(?P<numbers>\d*)$")
        if len(column) == 1:
            column = column[0]
        else:
            return 0,0
        position = column[0].xpath("string(@r)", namespaces=self.NAMESPACES)
        col_name, row_num = regex.search(position).groups()
        col_num = excel_col_name_to_number(col_name)
        return col_num, int(row_num)

    def _get_cell_value_and_dtype(self, column):
        dtype = column.xpath("string(@t)", namespaces=self.NAMESPACES)
        value = column.xpath("string(w:v)", namespaces=self.NAMESPACES)
        match dtype:
            case "e":
                typed_value = None
            case "b":
                typed_value = bool(value)
            case "inlineStr":
                # //TODO Inline String can be formatted. It is inside
                # of a <is> element, similar to sharedStrings.xml
                typed_value = str(value)
            case "s":
                typed_value = value
            case "str":
                # //TODO If value is present, get that. If not, display
                # formula in <f>
                typed_value = str(value)
            case "" | _:
                if not value:
                    typed_value = None
                elif value.isdigit():
                    typed_value = int(value)
                else:
                    try:
                        typed_value = float(value)
                    except ValueError:
                        typed_value = str(value)
        return typed_value, dtype

    def _get_cell_data(self, row):
        cols = []
        columns = (column for column in row.xpath("w:c", namespaces=self.NAMESPACES))
        for column in columns:
            col_num, row_num = self._get_cell_position(column)
            value, dtype = self._get_cell_value_and_dtype(column)
            cols.append(Cell(col=col_num, row=row_num, value=value, xl_dtype=dtype))
        return cols
    
    def _get_cell_data2(self, columns):
        cols = []
        for column in columns:
            col_num, row_num = self._get_cell_position(column)
            value, dtype = self._get_cell_value_and_dtype(column)
            cols.append(Cell(col=col_num, row=row_num, value=value, xl_dtype=dtype))
        return cols

    def _replace_shared_strings(
        self,
        row_data,
    ):
        for cell_data in row_data:
            if cell_data.xl_dtype == "s":
                xl_rich = self._rich_ss[int(cell_data.value)]
                plain = self._plain_ss[int(cell_data.value)]
                cell_data.value = plain
                cell_data.xl_rich = xl_rich
        return row_data
    
    @staticmethod
    def _create_formatdict(text_formats: list) -> dict:
        formatdict = {}
        for format in text_formats:
            ftag = etree.QName(format).localname
            fattrib = format.attrib
            formatdict[ftag] = fattrib
        return formatdict

    def _run_format(self, run_node: etree.Element) -> str:
        text_formats = run_node.xpath("./w:rPr/*", namespaces=self.NAMESPACES)
        formatdict = self._create_formatdict(text_formats)
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
                    #
                    # //TODO This needs to not rely on color and rgb keys. However, everything should
                    # work just fine, as the input formats going in have been cleaned and stripped of
                    # extraneous formatting properties.
                    if "color" in formatdict and "rgb" in formatdict["color"]:
                        if formatdict["color"]["rgb"] == "FFFF0000":
                            formats.append("z")
                    else:
                        formats.append("s")
        # Sort string of formats
        formats = "".join(sorted(formats))
        return formats

    @property
    @cache
    def _plain_ss(self) -> list:
        sharedstrings_plain = []
        groups = self._xml_tree["xl/sharedStrings.xml"].xpath(
            "/w:sst/w:si", namespaces=self.NAMESPACES
        )
        for group in groups:
            text = []
            for run in group:
                run_text = run.xpath("string(.//text())", namespaces=self.NAMESPACES)
                text.append(run_text)
            sharedstrings_plain.append("".join(text))
        return sharedstrings_plain

    @property
    @cache
    def _rich_ss(self):
        sharedstrings_rich = []
        groups = self._xml_tree["xl/sharedStrings.xml"].xpath(
            "/w:sst/w:si", namespaces=self.NAMESPACES
        )
        for group in groups:
            runs = []
            for run in group:
                text = run.xpath("string(.)", namespaces=self.NAMESPACES)
                formats = self._run_format(run)
                runs.append(Run(props=formats, text=text))
            sharedstrings_rich.append(runs)
        return sharedstrings_rich
