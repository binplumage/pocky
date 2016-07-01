import xlrd
import xlwt

def read_excel(file_name, sheet_number):
    data = xlrd.open_workbook(file_name)
    table = data.sheets()[sheet_number]
    return table

def get_init_excel(sheet_name):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)
    return ws

def write_data(ws, row, col, value, style):
    ws.write(row, col, value, style)

def get_row_number(table):
    return table.nrows

def get_col_number(table):
    return table.ncols
