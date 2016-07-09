import xlrd
import xlwt

def read_excel(file_name, sheet_number):
    data = xlrd.open_workbook(file_name)
    table = data.sheets()[sheet_number]
    return table

def get_new_sheet(wb, sheet_name):
    ws = wb.add_sheet(sheet_name)
    return ws

def get_init_excel():
    wb = xlwt.Workbook()
    return wb

def write_data(ws, row, col, value, style):
    ws.write(row, col, value, style)

def get_row_number(table):
    return table.nrows

def get_col_number(table):
    return table.ncols
