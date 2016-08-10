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

def copy_all_row_data(ws, table, row, ori_row):
    for i, j in enumerate([ i for i in range(9) if i !=1 ]):
        write_data(ws, row, i, get_cell_value(table, ori_row, j))

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

def write_credit(ws, data, i):
    for col, credit in enumerate(data, start=4):
        if credit==None:
            credit = " "
        write_data(ws, i, col, credit)

def write_processed_data(ws, line_number, data):
    for col, value in enumerate(data[0:-1]):
        write_data(ws, line_number, col, value)
    write_credit(ws, data[-1], line_number)

def create_result_title(ws, row):
    ws.write_merge(row,row,0,3,u"修課總學分數(A)",default_style)
    ws.write_merge(row+1,row+1,0,3,u"最低畢業學分數(B)",default_style)
    ws.write_merge(row+1,row+1,4,7,u"136",default_style)
    ws.write_merge(row+2,row+2,0,3,u"修課佔畢業學分數百分比(A/B)",default_style)
    ws.write_merge(row+3,row+3,0,3,u"IEET認證規範4課程學分數之要求",default_style)
    ws.write(row+3,4,u"34學分\n(25%)",default_style)
    ws.write_merge(row+3,row+3,5,6,u"57學分\n(37.5%)",default_style)
    ws.write_merge(row+4,row+4,0,3,u"是否符合",default_style)
    ws.write_merge(row+5,row+5,0,3,u"是否選修實務專題",default_style)

def write_result(ws, line_number, is_project):
    style_percent = xlwt.easyxf(num_format_str='0.00%')
    col = {4:"E", 5:"F", 7:"H"}
    for num in [4, 5, 7]:
        formula_SUM = "SUM("+str(col[num])+"3:"+str(col[num])+str(line_number)+")"
        ws.write(line_number, num, xlwt.Formula(formula_SUM))
        formula_AVE = str(col[num])+str(line_number+1)+"/136"
        ws.write(line_number + 2, num, xlwt.Formula(formula_AVE), style_percent)
        if num==4:
            ws.write(line_number + 4, num, xlwt.Formula("IF(%(l)s > 0.25;\"YES\";\"NO\")"% {"l":str(col[num]+str(line_number+3))}))
        if num==5:
            ws.write(line_number + 4, num, xlwt.Formula("IF(%(l)s > 0.375;\"YES\";\"NO\")"% {"l":str(col[num]+str(line_number+3))}))
    if is_project:
        ws.write(line_number + 5, 4, u"是")
    else:
        ws.write(line_number + 5, 4, u"否")
