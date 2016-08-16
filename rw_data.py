# -*- coding: utf-8 -*-

import xlrd
import xlwt
import setup_env

default_style = xlwt.easyxf('border: top thin, bottom thin, left thin, right thin; align: vert centre, horz center;')
OUTUT_DATA_TITLE_ORDER = [u"修課學年期",u"課程名稱",u"授課教授",u"必選修",u"學分",u"成績",u"學號",u"姓名"]
GET_DATA_TITLE_COL = {}

def get_title_col(table):
    global GET_DATA_TITLE_COL
    for i, j in enumerate(range(9)):
        GET_DATA_TITLE_COL[get_cell_value(table,0,j)] = i

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

def copy_all_row_data(ws, table, row, ori_row):
    for j, i in enumerate(OUTUT_DATA_TITLE_ORDER):
        write_data(ws, row, j, get_cell_value(table, ori_row, GET_DATA_TITLE_COL[i]))

def write_merge_data(ws, row, col, value, style=default_style):
    ws.write_merge(row[0], row[1], col[0], col[1], value, style)

def write_data(ws, row, col, value, style=default_style):
    ws.write(row, col, value, style)

def get_row_number(table):
    return table.nrows

def get_col_number(table):
    return table.ncols

def get_cell_value(table, i, j):
    return table.cell(i, j).value.strip()

def get_grade(table, row, register_year):
    semester_table = {0:u"一", 1:u"二", 2:u"三", 3:u"四", 4:u"五", 5:u"六", "01":u"上", "02":u"下", "0h":u"暑假"}
    raw_grade = get_cell_value(table, row, 0)
    semester = semester_table[raw_grade[-2:]]
    take_year = raw_grade[:-2]
    year = semester_table[int(take_year)-int(register_year)]
    return year+semester

def create_title(ws):
    write_merge_data(ws, [0, 1], [0, 0], u"年級")
    write_merge_data(ws, [0, 1], [1, 1], u"課程名稱")
    write_merge_data(ws, [0, 1], [2, 2], u"授課教授")
    write_merge_data(ws, [0, 1], [3, 3], u"必/選修")
    write_merge_data(ws, [0, 0], [4, 7], u"學分數")
    write_merge_data(ws, [1, 1], [5, 6], u"工程專業課程\n(含設計實作\n請打V)")
    write_data(ws, 1, 4, u"數學及\n基礎科學",default_style)
    write_data(ws, 1, 7, u"通識課程",default_style)
    ws.row(1).height_mismatch = True
    ws.row(1).height = 256*8

def write_credit(ws, data, i):
    for col, credit in enumerate(data, start=4):
        if credit==None:
            credit = " "
        write_data(ws, i, col, credit)

def write_processed_data(ws, row, data):
    for col, value in enumerate(data[0:-1]):
        write_data(ws, row, col, value)
    write_credit(ws, data[-1], row)

def create_result_title(ws, row):
    write_data(ws, row+3, 4, u"34學分\n(25%)")
    write_data(ws, row+3,7," ")
    write_merge_data(ws, [row, row], [0, 3], u"修課總學分數(A)")
    write_merge_data(ws, [row + 1, row + 1], [0, 3], u"最低畢業學分數(B)")
    write_merge_data(ws, [row + 1, row + 1], [4, 7], setup_env.GRADUATION_CREDIT_THRESHOLD )
    write_merge_data(ws, [row + 2, row + 2], [0, 3], u"修課佔畢業學分數百分比(A/B)")
    write_merge_data(ws, [row + 3, row + 3], [0, 3], u"IEET認證規範4課程學分數之要求")
    write_merge_data(ws, [row + 4, row + 4], [0, 3], u"是否符合")
    write_merge_data(ws, [row + 5, row + 5], [0, 3], u"是否選修實務專題")
    write_merge_data(ws, [row + 3, row + 3], [5, 6], u"57學分\n(37.5%)")

def is_credit_enough(ws, row, col, percent):
    formula_if = 'IF({0} > {1};"Yes";"No")'.format(str(col[0]) + str(row[0]), percent)
    transf_to_chinese = (("Yes", u"是") , ("No", u"否"))
    for i in transf_to_chinese:
        formula_if = (lambda ori_str, replace_word: ori_str.replace(replace_word[0], replace_word[1]))(formula_if, i)
    write_data(ws, row[1], col[1], xlwt.Formula(formula_if))

def write_result_data(ws, row, is_project):
    style_percent = xlwt.easyxf(num_format_str='0.00%')
    col = {4:("E", 0.25), 5:("F", 0.375), 6:("G",None), 7:("H", None)}
    is_take_project = {True : u"是", False : u"否"}

    for col_num in [4, 5, 7]:
        formula_SUM = "SUM("+str(col[col_num][0])+"3:"+str(col[col_num][0])+str(row)+")"
        write_data(ws, row, col_num, xlwt.Formula(formula_SUM))

    for col_num in col:
        formula_AVE = str(col[col_num][0])+str(row + 1)+"/"+str(setup_env.GRADUATION_CREDIT_THRESHOLD)
        if col[col_num][1]:
            is_credit_enough(ws, [row + 3 , row + 4 ], [col[col_num][0], col_num], col[col_num][1])
            write_data(ws, row + 2, col_num, xlwt.Formula(formula_AVE), style_percent)
        else:
            write_data(ws, row + 2, col_num, " ")
            write_data(ws, row + 4, col_num, " ")

    write_merge_data(ws, [row + 5, row + 5], [4, 7], is_take_project[is_project])

def write_result(ws, row, is_project):
    create_result_title(ws, row)
    write_result_data(ws, row, is_project)
