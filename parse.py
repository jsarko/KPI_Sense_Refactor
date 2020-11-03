from datetime import datetime
from calendar import monthrange
from pprint import pprint
import os
import sys
import json

import xlrd

import utilities


class Workbook:
    """ 
    Opens and returns a valid workbook object. 

    Keyword arguments:
    file_path <str> -- The absolute or relative file path of
        workbook to be opened.
    """

    def __init__(self, file_path):
        try:
            self.wb = xlrd.open_workbook(
                file_path)
        except FileNotFoundError:
            print("File does not exist.")
            sys.exit(1)

    def get_sheets(self):
        raise NotImplementedError


class Reader(Workbook):
    def get_sheet_by_name(self, sheet_name):
        try:
            return self.wb.sheet_by_name(sheet_name)
        except xlrd.biffh.XLRDError as xl_err:
            print(xl_err)
            sys.exit(1)

    def get_sheet_by_index(self):
        raise NotImplementedError

    def get_sheet_chunks(self, sheet_name, page_size, page_num):
        raise NotImplementedError


class Parser(Reader):
    def __init__(self, file_path, sheet_name):
        Workbook.__init__(self, file_path)
        self.sheet = self.get_sheet_by_name(sheet_name)
        self.last_col = self.sheet.ncols - 1
        self.last_row = self.sheet.nrows - 1

    def get_category_schema(self):
        """
        The function uses the field index and known columns 
        to parse a given sheet object and return a dictionary 
        schema of categories.
        """
        category_schema = []
        fields = self.sheet.col_values(2)
        for index, value in enumerate(fields):
            temp = next(
                (item for item in category_schema if item['name'] == value), None)
            col_d_current_cell_value = self.sheet.cell_value(index, 3)
            try:
                excel_date = datetime(
                    *xlrd.xldate_as_tuple(col_d_current_cell_value, self.wb.datemode))
                excel_end_date = datetime(
                    *xlrd.xldate_as_tuple(self.sheet.cell_value(index, self.last_col), self.wb.datemode))
                if utilities.is_category(excel_date):
                    if temp is None:
                        temp = {
                            "name": value,
                            "fields": utilities.get_fields(index, self.last_row, self.sheet, self.wb),
                            "subsets": ["All"],
                            "start_date": utilities.get_date(excel_date, is_first_of_month=True),
                            "end_date": utilities.get_date(excel_end_date),
                            "data": []
                        }
                        category_schema.append(temp)
                    else:
                        col_b_current_value = utilities.has_sub(
                            self.sheet.cell_value(index, 1))
                        if col_b_current_value:
                            temp["subsets"].append(col_b_current_value)
            except (TypeError, ValueError):
                continue
        return category_schema

    def parse_category_values(self, category_schema):
        pass
