# -*- coding: utf-8 -*-
import xlrd
import xlwt
import setup_env

default_style = xlwt.easyxf('border: top thin, bottom thin, left thin, right thin; align: vert centre, horz center;')

def read_excel(file_name, sheet_number):
    try:
        data = xlrd.open_workbook(file_name)
        table = data.sheets()[sheet_number]
        return table
    except:
        setup_env.display_message(u"Cannot open file : " + file_name)

def get_new_sheet(wb, sheet_name):
    ws = wb.add_sheet(sheet_name)
    return ws

def get_init_excel():
    wb = xlwt.Workbook()
    return wb

def write_all_row_data(ws, table, row, ori_row):
    for j in range(9):
        write_data(ws, row, j, table.cell(ori_row, j).value.rstrip())

def write_data(ws, row, col, value, style=default_style):
    ws.write(row, col, value, style)

def get_row_number(table):
    return table.nrows

def get_col_number(table):
    return table.ncols

def get_table_value(table, i, j):
    return table.cell(i, j).value

def get_grade(table, row, register_year):
    semester_table = {0:u"一", 1:u"二", 2:u"三", 3:u"四", 4:u"五", 5:u"六", "01":u"上", "02":u"下", "0h":u"暑假"}
    raw_grade = get_table_value(table, row, 0)
    semester = semester_table[raw_grade[-2:]]
    take_year = raw_grade[:-2]
    year = semester_table[int(take_year)-int(register_year)]
    return year+semester

def create_title(ws):
    ws.write_merge(0, 1, 0, 0, u"年級",default_style)
    ws.write_merge(0, 1, 1, 1, u"課程名稱",default_style)
    ws.write_merge(0, 1, 2, 2, u"授課教授",default_style)
    ws.write_merge(0, 1, 3, 3, u"必/選修",default_style)
    ws.write_merge(0, 0, 4, 7, u"學分數",default_style)
    ws.write_merge(1, 1, 5, 6, u"工程專業課程\n(含設計實作\n請打V)",default_style)
    write_data(ws, 1, 4, u"數學及\n基礎科學",default_style)
    write_data(ws, 1, 7, u"通識課程",default_style)
    ws.row(1).height_mismatch = True
    ws.row(1).height = 256*8
