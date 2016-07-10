import xlrd
import xlwt

style = xlwt.easyxf('border: top thin, bottom thin, left thin, right thin; font: bold on; align: vert centre, horz center;')

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

def write_all_row_data(ws, table, row, ori_row):
    for j in range(9):
        write_data(ws, row, j, table.cell(ori_row, j).value.rstrip())

def write_data(ws, row, col, value):
    global style
    ws.write(row, col, value, style)

def get_row_number(table):
    return table.nrows

def get_col_number(table):
    return table.ncols
