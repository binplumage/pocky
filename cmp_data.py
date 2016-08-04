#-*- coding: utf-8 -*-
import setup_env
import rw_data
import re

def is_pass(table, i):
    scores = table.cell(i, 6).value.rstrip()
    if scores in ["A+","A ","A-","B+","B ","B-","C+","C ","C-", "P"]:
        return True
    else:
        return False

def is_bd(table, i):
    sid = table.cell(i, 7).value.rstrip()
    if sid[0] == 'B':
        return True
    else:
        return False

def filter_data(table):
    wb = rw_data.get_init_excel()
    ws = rw_data.get_new_sheet(wb, "sheet 1")
    count = 0
    for i in range(1, table.nrows):
        if is_bd(table, i) and is_pass(table, i):
            rw_data.copy_all_row_data(ws, table, count, i)
            count = count + 1
    wb.save(setup_env.FILTER_DATA)

def get_field_table():
    table = rw_data.read_excel(setup_env.FIELD_TABLE_FILE, 0)
    field_table = {}

    for i in range(1, table.nrows):
        field_table[table.cell(i,0).value] = [table.cell(i,1).value, table.cell(i,2).value]
    return field_table

def get_credit_in_field(title, ori_credit):

    field_table = get_field_table()
    if title in field_table:
        credit = field_table[title]
        if re.search(u"實習",title):
            credit.extend(["V",None])
        else:
            credit.extend([None, None])
    else:
        credit = [None, None, None, ori_credit]

    return credit

def change_grade_format(register_year, ori_grade):

    change_content = {0:u"一", 1:u"二", 2:u"三", 3:u"四", 4:u"五", 5:u"六", "01":u"上", "02":u"下", "0h":u"暑假"}
    grade = change_content[int(ori_grade[0:-2]) - register_year]
    semester = change_content[ori_grade[-2:]]
    return grade + semester

def get_data(table, row, register_year):
    grade = change_grade_format(table, row, register_year)
    title = rw_data.get_cell_value(table, row, 2)
    teacher = rw_data.get_cell_value(table, row, 3)
    requ_or_ele = rw_data.get_cell_value(table, row, 4)
    credit = rw_data.get_cell_value(table, row, 5)
    return [grade, title, teacher, requ_or_ele, credit]


def get_register_year(sid):
    return int(sid[1:-5])

def get_sid_and_line_number(table):
    sid_and_ln = {}

    for i in range(table.nrows):
        sid = table.cell(i, 7).value
        if not sid in sid_and_ln:
            sid_and_ln[sid] = [i]
        else:
            sid_and_ln[sid].append(i)

    return sid_and_ln
