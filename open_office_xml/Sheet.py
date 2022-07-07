from functools import cache
from itertools import groupby
import logging

from lxml import etree

from .dataclasses import SheetData
from .FileTree import FileTree
from .SheetFunctions import SheetFunctions


class Sheet(FileTree, SheetFunctions):
    def __init__(self, filepath, sheetname, header_row=1):
        super(Sheet, self).__init__(filepath)
        self.sheetname = sheetname
        self.header_row = header_row

        logging.info(f"Reading sheet '{self.sheetname}' from {self.filepath}...")

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
            row_data = self._get_cell_data(row)
            row_data = self._replace_shared_strings(row_data)
            sheet_data.data.extend(row_data)
        # Add column name data
        for cell_data in sheet_data.data:
            for k, v in self.header.items():
                if cell_data.col == k:
                    cell_data.col_name = v
        return sheet_data

    def get_column(self, column_num=1, rich=True):
        return {
            cell.row: (cell.rich if rich else cell.value)
            for cell in self.sheetdata.data
            if cell.col == column_num
        }

    def get_row(self, row_num=1, rich=False):
        return {
            cell.col: (cell.rich if rich else cell.value)
            for cell in self.sheetdata.data
            if cell.row == row_num
        }

    def group_by_row(self):
        sheet_col_names = sorted(self.sheetdata.data, key=(keyfunc := lambda x: x.row))
        return (list(group) for _, group in groupby(sheet_col_names, key=keyfunc))
