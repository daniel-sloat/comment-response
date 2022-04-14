#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
from lxml import etree
import pandas as pd
import numpy as np
import re
    
class SpreadSheetML():
    """Uses the Open Office Spreadsheet standard to create simple dataframes
    and retrieve rich formatted text strings from .xlsx spreadsheets.
    
    While the pandas package can easily transform spreadsheets to dataframes, 
    the pandas package can't read rich formatted text. The rich formatted text
    can be paired together with the coded dataframe.
    """
    # http://officeopenxml.com/anatomyofOOXML-xlsx.php
    # https://www.dyalog.com/uploads/conference/dyalog18/presentations/U10_Excel_Mining_pt2.pdf
    NAMESPACES = {
        "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/package/2006/relationships"
        }
    
    def __init__(self, filepath) -> None:
        self.unzipped = zipfile.ZipFile(filepath)
        self.sheetname = "sheet1"
        pass
        
    def sheet(self,sheet: str):
        self.sheetname = sheet
        return self
    
    def __sharedstrings_xml(self) -> etree.XML:
        xml = self.unzipped.read("xl/sharedStrings.xml")
        return etree.XML(xml)
    
    def __workbook_xml(self) -> etree.XML:
        xml = self.unzipped.read("xl/workbook.xml")
        return etree.XML(xml)
    
    def __sheet_xml(self) -> etree.XML:
        xml = self.unzipped.read("xl/worksheets/" + self.sheetname + ".xml")
        return etree.XML(xml)
    
    def __tag(self,element: etree.Element,type_char: str="") -> bool:
        return element.tag == "{%s}%s" % (self.NAMESPACES["w"],type_char)
    
    @staticmethod
    def __excel_col_name_to_number(col_index: str) -> int:
        if not col_index.isalpha(): raise TypeError
        pow = 1
        col_num = 0
        for letter in reversed(col_index.upper()):
            col_num += (ord(letter) - ord("A") + 1) * pow
            pow *= 26
        return col_num
    
    def __cell_code_value(self, column: etree.Element) -> str | int | float:
        # Default value for cell (default: NaN)
        code = np.nan
        # If string in sharedStrings or is inlineStr:
        if "t" in column.attrib:
            if column.attrib["t"] == "s" or \
                column.attrib["t"] == "inlineStr":
                for value in column:
                    if self.__tag(value, "v"):
                        code = value.text
        # If number, keep int or float:
        else:
            for value in column:
                if value.text.isdigit():
                    code = int(value.text)
                else:
                    code = float(value.text)
        return code
    
    def __sheet_grid_data(self) -> list:
        # Regex splits Excel column name (e.g., "A34") to letters and numbers
        regex = re.compile(r"^(?P<letters>\w*?)(?P<numbers>\d*)$")
        sheet_grid_data = []
        for data in [parent for parent in self.__sheet_xml() 
                     if self.__tag(parent,"sheetData")]:
            for row in [row for row in data if self.__tag(row, "row")]:
                row_num = int(row.attrib["r"])
                for column in [column for column in row if self.__tag(column, "c")]:
                    col_name = regex.search(column.attrib["r"]).group("letters")
                    col_num = self.__excel_col_name_to_number(col_name)
                    code = self.__cell_code_value(column)
                    sheet_grid_data.append([row_num, col_num, code])
        sheet_grid_data.sort()
        return sheet_grid_data
    
    @staticmethod
    def __create_formatdict(run_node: etree.Element) -> dict:
        text_formats = run_node.xpath("./w:rPr/*", namespaces=SpreadSheetML.NAMESPACES)
        formatdict = {}
        for format in text_formats:
            ftag = etree.QName(format).localname
            fattrib = format.attrib
            formatdict[ftag] = fattrib
        return formatdict
    
    @staticmethod
    def __run_format(run_node: etree.Element) -> str:
        formatdict = SpreadSheetML.__create_formatdict(run_node)
        formats = ""
        for k, v in formatdict.items():
            match k: 
                case "b": formats += "b"
                case "i": formats += "i"
                case "u":
                    if "val" in v and "double" in v["val"]:
                        formats += "w"
                    else:
                        formats += "u"
                case "vertAlign":
                    if "val" in v:
                        if v["val"] == "superscript":
                            formats += "x"
                        if v["val"] == "subscript":
                            formats += "v"
                case "strike":
                    # Double-strikethrough does not exist in Excel. Strike + red text 
                    # is used to denote double-strikethrough
                    if "color" in formatdict and "rgb" in formatdict["color"]:
                        if formatdict["color"]["rgb"] == "FFFF0000":
                            formats += "z"
                    else:
                        formats += "s"
        # Sort string of formats
        formats_sorted = sorted(formats)
        formats = "".join(formats_sorted)
        return formats
    
    def __get_plain_ss(self) -> list:
        sharedstrings_plain = []
        xml = self.__sharedstrings_xml()
        groups = xml.xpath("/w:sst/w:si", namespaces=self.NAMESPACES)
        for group in groups:
            text = ""
            for run in group:
                text += run.xpath("string(.//text())",namespaces=self.NAMESPACES)
            sharedstrings_plain.append(text)
        return sharedstrings_plain
    
    def __ss_dict(self) -> dict:
        ss_plain = {}
        xml = self.__sharedstrings_xml()
        groups = xml.xpath("/w:sst/w:si", namespaces=self.NAMESPACES)
        for group_no, group in enumerate(groups):
            text = ""
            for run in group:
                text += run.xpath("string(.//text())",namespaces=self.NAMESPACES)
            ss_plain[str(group_no)] = text
        return ss_plain
    
    def get_rich_strings(self) -> pd.DataFrame:
        sharedstrings_rich = []
        xml = self.__sharedstrings_xml()
        groups = xml.xpath("/w:sst/w:si", namespaces=self.NAMESPACES)
        for group_no, group in enumerate(groups):
            for run in group:
                text = run.xpath("string(.)", namespaces=self.NAMESPACES)
                formats = self.__run_format(run)
                sharedstrings_rich.append((group_no,formats,text))
        return sharedstrings_rich
    
    @staticmethod
    def __replace_with_strings(
        plain_ss: list,
        sheet_grid_data: list
    ) -> list:
        for data in sheet_grid_data:
            # Values as strings are replaced by actual string.
            if isinstance(data[2], str):
                data[2] = plain_ss[int(data[2])]
        return sheet_grid_data
    
    @staticmethod
    def __make_worksheet(grid_data: list) -> pd.DataFrame:
        columns=["Row","Column","Value"]
        df = pd.DataFrame(grid_data, columns=columns)
        df_worksheet = df.pivot(*columns).reset_index(drop=True)
        # Make first row the header
        header = df_worksheet.iloc[0]
        df_worksheet = df_worksheet[1:]
        df_worksheet.columns = header
        df_worksheet = df_worksheet.reset_index(drop=True)
        return df_worksheet
    
    def __get_headers(self) -> list:
        headers = []
        for data in self.__sheet_grid_data():
            if data[0] == 0:
                headers.append(data)
        return headers
    
    def to_dataframe(self) -> pd.DataFrame:
        #data_replaced = self.__replace_with_strings(self.__get_plain_ss(),self.__sheet_grid_data())
        columns=["Row","Column","Value"]
        df_data = pd.DataFrame(self.__sheet_grid_data(), columns=columns)
        df_worksheet = df_data.pivot(*columns).reset_index(drop=True)
        df_worksheet = df_worksheet.replace(self.__ss_dict())
        # Make first row the header
        header = df_worksheet.iloc[0]
        df_worksheet = df_worksheet[1:]
        df_worksheet.columns = header
        df_worksheet = df_worksheet.reset_index(drop=True)
        return df_worksheet
    
    def to_dataframe_codes(self) -> pd.DataFrame:
        # Create dataframe, and pivot
        grid_data_columns = ["Row","Column","Value"]
        df_sheet = pd.DataFrame(self.__sheet_grid_data(), columns=grid_data_columns)
        df_sheet = df_sheet.pivot(*grid_data_columns).reset_index(drop=True)
        # Make first row the header
        df_sheet = df_sheet[1:]
        df_sheet.columns = self.to_dataframe().columns
        df_sheet = df_sheet.reset_index(drop=True)
        return df_sheet