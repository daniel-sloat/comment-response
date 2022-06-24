from functools import cache
from itertools import groupby
import re

from lxml import etree
from lxml.etree import XPath

from .dataclasses import SheetData
from .FileTree import FileTree
from .SheetFunctions import SheetFunctions


class Sheet(FileTree, SheetFunctions):
    def __init__(self, filepath, sheetname, header_row=1):
        super(Sheet, self).__init__(filepath)
        self.sheetname = sheetname
        self.header_row = header_row

    @property
    @cache
    def header(self):
        header_r = self._sheet_roots[self.sheetname].xpath(
            f"w:sheetData/w:row[@r={self.header_row}]", namespaces=self.NAMESPACES
        )[0]
        row_data = self._get_cell_data(header_r)
        row_data = self._replace_shared_strings(row_data)
        return {cell.col: cell.value for cell in row_data}

    @property
    @cache
    def sheetdata(self) -> SheetData:
        rows = self._sheet_roots[self.sheetname].xpath(
            f"w:sheetData/w:row[@r>{self.header_row}]", namespaces=self.NAMESPACES
        )
        sheet_data = SheetData()
        for row in rows:
            # print("Getting cell data...")
            row_data = self._get_cell_data(row)
            # print("Replacing shared strings...")
            row_data = self._replace_shared_strings(row_data)
            sheet_data.data.extend(row_data)
        return sheet_data

    @property
    def sheetdata_regex(self) -> SheetData:
        print("Getting specific sheetdata")
        col_letters = "W", "X", "K", "AC", "AD", "AE"
        row_num = 1
        columns = []
        etree.XPath()
        c = self._sheet_roots[self.sheetname].XPath(
            f"w:sheetData/w:row[@r>'{row_num}']/w:c", namespaces=self.NAMESPACES
        )
        for col in col_letters:
            print("RUN")
            regex = f"node()[re:match(@r,'([A-Z]+)') = '{col}']"
            c(self._sheet_roots[self.sheetname])
            columns.extend(
                self._sheet_roots[self.sheetname].xpath(
                    regex, namespaces=self.NAMESPACES
                )
            )
        print("Getting cell data")
        data = self._get_cell_data2(columns)
        print("Replacing shared strings")
        data = self._replace_shared_strings(data)
        return data

    @property
    def sheetdata_regex(self) -> SheetData:
        columns = []
        row_num = self.header_row - 1
        cols = ["W", "X", "K", "AC", "AD", "AE"]

        regex = f"w:sheetData/w:row[@r>'{self.header_row}']/w:c"

        self._sheet_roots[self.sheetname].findall("c")
        columns.extend(
            self._sheet_roots[self.sheetname].xpath(regex, namespaces=self.NAMESPACES)
        )
        k = []

        k = [
            data_node
            for column in columns
            if (data_node := re.match("([A-Z]+)", column.attrib["r"]).group()) in cols
        ]

        # for column in columns:
        #    for col in cols:
        #        data_node = re.match("([A-Z]+)", column.attrib["r"]).group()
        #        k.append(data_node)
        print(k)
        data = self._get_cell_data2(columns)
        data = self._replace_shared_strings(data)
        return data

    @property
    def sheetdata_regex_with_column_names(self):
        sheet_data = self.sheetdata_regex
        for cell_data in sheet_data:
            for k, v in self.header.items():
                if cell_data.col == k:
                    cell_data.col_name = v
        # print(sheet_data)
        return sheet_data

    @property
    @cache
    def sheetdata_with_column_names(self):
        sheet_data = self.sheetdata.data
        for cell_data in sheet_data:
            for k, v in self.header.items():
                if cell_data.col == k:
                    cell_data.col_name = v
        return sheet_data

    def get_column(self, column_num=1, rich=True):
        column = {
            cell.row: (cell.rich if rich else cell.value)
            for cell in self.sheetdata.data
            if cell.col == column_num
        }
        return column

    def get_row(self, row_num=1, rich=False):
        row = {
            cell.col: (cell.rich if rich else cell.value)
            for cell in self.sheetdata.data
            if cell.row == row_num
        }
        return row

    def group_by_row(self):
        keyfunc = lambda x: x.row
        sheet_col_names = sorted(self.sheetdata_with_column_names, key=keyfunc)
        return (list(group) for _, group in groupby(sheet_col_names, key=keyfunc))

    def group_by_row2(self):
        keyfunc = lambda x: x.row
        sheet_col_names = sorted(self.sheetdata_regex_with_column_names, key=keyfunc)
        return (list(group) for _, group in groupby(sheet_col_names, key=keyfunc))
