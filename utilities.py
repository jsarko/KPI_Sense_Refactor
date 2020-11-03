from datetime import datetime
from calendar import monthrange

import xlrd


def get_date(excel_date, is_first_of_month=False):
    """This function checks that the serialized dates read from the excel file are
    an actual date and not a fake date parsed from raw values. It does this by verifying the
    format is YYYY-mm-ddTHH:MM:SS.s+z and that the day is either the first or last of the month,
    whichever the function specified.
    """
    try:
        converted_date = excel_date.strftime("%Y-%m-%dT%H:%M:%S.000+00:00")
        datetime.strptime(
            converted_date, f"%Y-%m-{monthrange(excel_date.year, excel_date.month)[1]}T00:00:00.000+00:00")
        if is_first_of_month:
            return excel_date.strftime("%Y-%m-01T%H:%M:%S.000+00:00")
        else:
            return converted_date
    except (ValueError, TypeError, AttributeError):
        return False


def has_sub(value_to_validate):
    if value_to_validate != "":
        return value_to_validate
    else:
        return False


def get_fields(index, last_row, sheet, wb):
    # iterate over fields until a field evaluates to category
    fields = []
    index = index + 1
    while(index != last_row):
        cell_value = sheet.cell_value(index, 2)
        cell_d_value = sheet.cell_value(index, 3)
        try:
            excel_date = datetime(
                *xlrd.xldate_as_tuple(cell_d_value, wb.datemode))
            if is_category(excel_date) is True:
                break
        except (ValueError, TypeError, xlrd.xldate.XLDateNegative):
            pass
        if cell_value != "":
            fields.append(cell_value.strip())
        index = index + 1
    return fields


def is_category(cell_to_validate):
    """
    This function evalates if the given cell value is a
    date or not. If so, the value in the current row of
    column C is a category.
    """
    return True if get_date(cell_to_validate) else False
