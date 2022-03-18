#!/usr/bin/env python3.10
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
    
    def __sheet_xml(self) -> etree.XML:
        xml = self.unzipped.read("xl/worksheets/" + self.sheetname + ".xml")
        return etree.XML(xml)
    
    def __tag(self,element: etree.Element,type_char: str="") -> bool:
        return element.tag == "{%s}%s" % (self.NAMESPACES["w"],type_char)
    
    @staticmethod
    def __excel_col_name_to_number(col_index: str) -> int:
        pow = 1
        col_num = 0
        for letter in reversed(col_index.upper()):
            col_num += (ord(letter) - ord("A") + 1) * pow
            pow *= 26
        return col_num
    
    def __cell_code_value(self, column: etree.Element) -> str | int | float:
        # Default value for cell (default: NaN)
        code = np.nan
        # If string in sharedStrings:
        type_text = "t"
        shared_string = "s"
        inline_string = "inlineStr"
        if type_text in column.attrib:
            if column.attrib[type_text] == shared_string or \
                column.attrib[type_text] == inline_string:
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
    def __run_format(formatdict: dict) -> str:
        formats = ""
        # Double-strikethrough does not exist in Excel. Strike + red text 
        # is used to denote double-strikethrough
        if "strike" in formatdict.keys():
            if "color" in formatdict.keys():
                if "rgb" in formatdict["color"].keys():
                    if formatdict["color"]["rgb"] == "FFFF0000":
                        formats += "z"
            else:
                formats += "s"
        # All other formatting
        for k, v in formatdict.items():
            if k == "b":
                formats += "b"
            if k == "i":
                formats += "i"
            if k == "u":
                if "val" in v.keys():
                    if v["val"] == "double":
                        formats += "w"
                else:
                    formats += "u"
            if k == "vertAlign":
                if "val" in v.keys():
                    if v["val"] == "superscript":
                        formats += "x"
                    if v["val"] == "subscript":
                        formats += "v"
        # Alphabetically sort string of formats
        formats_sorted = sorted(formats)
        formats = "".join(formats_sorted)
        return formats
    
    def __get_plain_ss(self) -> list:
        sharedstrings_plain = []
        for group_no, parent in enumerate(self.__sharedstrings_xml()):
            if self.__tag(parent, "si"):
                for run in parent:
                    text = run.xpath("string(.)",namespaces=self.NAMESPACES)
                    sharedstrings_plain.append([group_no,text])
        return sharedstrings_plain
    
    def get_rich_strings(self) -> pd.DataFrame:
        sharedstrings_rich = []
        for group_no, parent in enumerate(self.__sharedstrings_xml()):
            if self.__tag(parent, "si"):
                for run in parent:
                    if self.__tag(run, "t"):
                        text = run.xpath("string(.)",namespaces=self.NAMESPACES)
                        newline = False
                        formats = ""
                    elif not self.__tag(run, "t"):
                        formats = ""
                        for formatted_text_run in run:
                            if self.__tag(formatted_text_run, "t"):
                                text = formatted_text_run.text
                            if self.__tag(formatted_text_run, "rPr"):
                                formatdict = {}
                                for format in formatted_text_run:
                                    ftag = etree.QName(format).localname
                                    fattrib = format.attrib
                                    formatdict[ftag] = fattrib
                                formats = self.__run_format(formatdict)
                        # For determining paragraph breaks later
                        newline = False        
                        if text == None:
                            text = ""
                        if "\n" in text:
                            newline = True
                            text = text.removesuffix("\n")
                    sharedstrings_rich.append([group_no,newline,formats,text])
        return sharedstrings_rich
    
    def to_dataframe(self) -> pd.DataFrame:
        df_sharedstrings = pd.DataFrame(self.__get_plain_ss(), columns=["Group","String"])
        sharedstrings = df_sharedstrings.groupby(["Group"])["String"].apply("".join)
        
        data_replaced = []
        for data in self.__sheet_grid_data():
            # Values as strings are replaced by actual string.
            if type(data[2]) == str:
                data_replaced.append([data[0],data[1],sharedstrings[int(data[2])]])
            else:
                data_replaced.append(data)
            
        columns=["Row","Column","Value"]
        df_data = pd.DataFrame(data_replaced, columns=columns)
        df_worksheet = df_data.pivot(*columns).reset_index(drop=True)
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