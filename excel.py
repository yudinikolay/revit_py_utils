import clr

clr.AddReference('Microsoft.Office.Interop.Excel')
from Microsoft.Office.Interop import Excel


import System
from System.Runtime.InteropServices import Marshal
from System import Array

import sys

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os.path
# Regular expressions
import re

import string


def to_str(val):
    if val is not None:
        return "=\"" + str(val) + "\""
    return None


def check_import_table_input(file_name_str, table_name_str):
    if not (isinstance(file_name_str, str) and isinstance(table_name_str, str)):
        return 1
    if not (file_name_str.endswith('.xlsx')):
        return 2
    if not os.path.isfile(file_name_str):
        return 3
    return 0


def check_import_range_input(file_name_str, sheet_name_str, range_name_str):
    if not (isinstance(file_name_str, str) and isinstance(sheet_name_str, str) and isinstance(range_name_str, str)):
        return 1
    if not (file_name_str.endswith('.xlsx')):
        return 2
    if not os.path.isfile(file_name_str):
        return 3
    return 0


def setup_excel_app(excel):
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    return excel


def error_message(err_code_int):
    if err_code_int == 0:
        return 'No errors found'
    if err_code_int == 1:
        return 'Input data types is incorrect'
    if err_code_int == 2:
        return 'File extension is incorrect'
    if err_code_int == 3:
        return 'File not found'
    if err_code_int == 4:
        return ' not found'
    return 'UNDEFINED ERROR!'


def convert_chars(chars_str):
    num = 0
    for char_str in chars_str:
        if char_str in string.ascii_letters:
            num = num * 26 + (ord(char_str.upper()) - ord('A')) + 1
    return num


def cell_index(cell_address_str):
    match = re.match(r"([a-z]+)([0-9]+)", cell_address_str, re.I)
    if match:
        address_items = match.groups()
        row = convert_chars(address_items[0])
        column = int(address_items[1])
        return [row, column]
    return None


def range_address(address_str):
    address_split = address_str.split(":")
    if len(address_split) == 2:
        origin_address = cell_index(address_split[0])
        extent_address = cell_index(address_split[1])
        if not (origin_address is None or extent_address is None):
            origin_row = int(origin_address[0])
            origin_col = int(origin_address[1])
            extent_row = int(extent_address[0])
            extent_col = int(extent_address[1])
            return [origin_col, origin_row, extent_col, extent_row]
    return None


def table_by_name(work_book, name_str):
    for sheet in work_book.Worksheets:
        count_tables = sheet.ListObjects.Count
        if count_tables > 0:
            for i in range(count_tables):
                if sheet.ListObjects[i + 1].Name == name_str:
                    res = sheet.ListObjects[name_str]
                    Marshal.ReleaseComObject(sheet)
                    return res
            Marshal.ReleaseComObject(sheet)
    return None


def table_data(work_table):
    table_headers_val = list(work_table.HeaderRowRange.Value2)
    table_data_val = work_table.DataBodyRange.Value2
    result = [table_headers_val]
    result.extend(range_data(table_data_val))
    return result


def range_by_address(work_sheet, range_address_str):
    range_address_list = range_address(range_address_str)
    if range_address_list is not None:
        cell_origin = work_sheet.Cells(range_address_list[0], range_address_list[1])
        cell_extent = work_sheet.Cells(range_address_list[2], range_address_list[3])
        return work_sheet.Range[cell_origin, cell_extent].Value2
    return None


def range_by_name(work_sheet, range_name_str):
    try:
        return work_sheet.Range(range_name_str).Value2
    except EnvironmentError:
        return None


def range_by_string(work_sheet, range_str):
    if ':' in range_str:
        return range_by_address(work_sheet, range_str)
    else:
        # try get range by name
        return range_by_name(work_sheet, range_str)


def range_data(work_range):
    result = list()
    for i in range(work_range.GetLowerBound(0) - 1, work_range.GetUpperBound(0), 1):
        result_row = list()
        for j in range(work_range.GetLowerBound(1) - 1, work_range.GetUpperBound(1), 1):
            result_row.append(work_range[i, j])
        result.append(result_row)
    return result


def exit_excel(excel, obj_list):
    # clean up before exiting excel, if any COM object remains
    # unreleased then excel crashes on open following time
    def clean_up(_list):
        if isinstance(_list, list):
            for item in _list:
                Marshal.ReleaseComObject(item)
        else:
            Marshal.ReleaseComObject(_list)
        return None

    excel.ActiveWorkbook.Close(True)
    excel.ScreenUpdating = True
    clean_up(obj_list)
    return None


def import_table(file_name_str, table_name_str):
    check_results = check_import_table_input(file_name_str, table_name_str)
    if check_results == 0:
        excel = setup_excel_app(Excel.ApplicationClass())
        excel.Workbooks.open(file_name_str)
        work_book = excel.ActiveWorkbook
        work_table = table_by_name(work_book, table_name_str)
        if work_table is not None:
            result = table_data(work_table)
        else:
            result = table_name_str + error_message(4)
        exit_excel(excel, [work_book, excel])
        return result
    else:
        return error_message(check_results)


def sheet_by_name(work_book, sheet_name_str):
    try:
        return work_book.Sheets(sheet_name_str)
    except EnvironmentError:
        return None


def import_range(file_name_str, sheet_name_str, range_name_str):
    check_results = check_import_range_input(file_name_str, sheet_name_str, range_name_str)
    if check_results == 0:
        excel = setup_excel_app(Excel.ApplicationClass())
        excel.Workbooks.open(file_name_str)
        work_book = excel.ActiveWorkbook
        work_sheet = sheet_by_name(work_book, sheet_name_str)
        if work_sheet is not None:
            work_range = range_by_string(work_sheet, range_name_str)
            if work_range is not None:
                result = range_data(work_range)
            else:
                result = range_name_str + error_message(4)
            Marshal.ReleaseComObject(work_sheet)
        else:
            result = sheet_name_str + error_message(4)
        exit_excel(excel, [work_book, excel])
        return result
    else:
        return error_message(check_results)


def create_array(data_list):
    len_x = len(data_list[0])
    len_y = len(data_list)
    arr = Array.CreateInstance(object, len_y, len_x)
    for y in range(len(data_list)):
        for x in range(len(data_list[0])):
            arr[y, x] = data_list[y][x]
    return arr


def convert_number(num):
    letters = ''
    while num:
        mod = num % 26
        num = num // 26
        letters += chr(mod + 64)
    return ''.join(reversed(letters))


def get_len_x(arr_list):
    len_max = 0
    for row in arr_list:
        len_curr = len(row)
        if len_curr > len_max:
            len_max = len_curr
    return len_max


def get_range_str(col_num, row_num, arr_list):
    col_start = convert_number(col_num + 1)
    row_start = str(row_num + 1)
    col_end = convert_number(col_num + get_len_x(arr_list))
    row_end = str(row_num + len(arr_list))
    return col_start + row_start + ':' + col_end + row_end


def get_xl_table(workbook, name_str):
    for sheet in workbook.Worksheets:
        count_tables = sheet.ListObjects.Count
        if count_tables > 0:
            for i in range(count_tables):
                if sheet.ListObjects[i + 1].Name == name_str:
                    res = sheet.ListObjects[name_str]
                    Marshal.ReleaseComObject(sheet)
                    return res
            Marshal.ReleaseComObject(sheet)
    return None


def get_worksheet(workbook, sheet_name_str):
    worksheets = [sn for sn in workbook.Worksheets if sn.Name == sheet_name_str]
    if len(worksheets) > 0:
        worksheet = worksheets[0]
        worksheet.Cells.ClearContents()
    else:
        worksheet = workbook.Worksheets.Add()
        worksheet.Name = sheet_name_str
    return worksheet


# Creating new workbook
def new_workbook(excel, filename_str):
    workbook = excel.Workbooks.Add()
    workbook.SaveAs(filename_str)
    return workbook


# Getting workbook
def get_workbook(excel, filename_str):
    workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename_str]
    if len(workbooks) > 0:
        return workbooks[0]
    elif os.path.isfile(filename_str):
        return excel.Workbooks.Open(filename_str)
    else:
        return new_workbook(excel, filename_str)


def create_table(work_book, name_str, data_list, sheet_visible=False, col_int=0, row_int=0):
    sheet_name = name_str + ""
    table_name = name_str + "_table"
    # opening/creating worksheet
    work_sheet = get_worksheet(work_book, sheet_name)
    # get excel range address by data size
    start_cell = work_sheet.Cells(1, 1)
    end_cell = work_sheet.Cells(len(data_list), len(data_list[0]))
    # get excel range by start/end cells
    xl_range = work_sheet.Range(start_cell, end_cell)
    # convert data from list to array
    range_value = create_array(data_list)
    # write data to range value
    xl_range.Value2 = range_value
    # creating named range
    work_sheet.ListObjects.Add(1, xl_range, System.Type.Missing, 1, System.Type.Missing).Name = table_name
    # get named range
    table = work_sheet.ListObjects(table_name)
    # apply table style to range
    table.TableStyle = "TableStyleMedium15"
    # hide data list if needed
    if not sheet_visible:
        work_sheet.Visible = False
    # release COM-objects
    Marshal.ReleaseComObject(xl_range)
    Marshal.ReleaseComObject(table)
    Marshal.ReleaseComObject(work_sheet)
    return True


def export_tables(file_name_str, data_names, data_arrays, sheet_visible=False):
    # Get Excel Application
    excel = Excel.ApplicationClass()
    # Setup Excel
    excel = setup_excel_app(excel)
    # Opening/Creating workbook
    work_book = get_workbook(excel, file_name_str)
    # Creating table in workbook
    for data_name, data_array in zip(data_names, data_arrays):
        create_table(work_book, data_name, data_array, sheet_visible)

    # Save&Close
    work_book.Save()
    exit_excel(excel, work_book)
    return file_name_str


def export_table(file_name_str, data_name_str, array_list, sheet_visible=True):
    # Get Excel Application
    excel = Excel.ApplicationClass()
    # Setup Excel
    excel = setup_excel_app(excel)
    # Opening/Creating workbook
    work_book = get_workbook(excel, file_name_str)
    # Creating table in workbook
    create_table(work_book, data_name_str, array_list, sheet_visible)
    # Save&Close
    work_book.Save()
    # Marshal.ReleaseComObject(table)
    exit_excel(excel, work_book)
    return [file_name_str, data_name_str, array_list]

