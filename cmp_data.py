#-*- coding: utf-8 -*-
import setup_env
import rw_data
import copy
import re

GRADUATION_CREDIT_THRESHOLD = 128

def is_pass(table, i):
    pass_sco = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "P"]
    scores = rw_data.get_cell_value(table, i, rw_data.GET_DATA_TITLE_COL[u"成績"])

    return True if scores in pass_sco else False

def is_bd(table, i):
    sid = rw_data.get_cell_value(table, i, rw_data.GET_DATA_TITLE_COL[u"學號"])

    return True if sid[0] == 'B' else False

def filter_data(table):
    wb = rw_data.get_init_excel()
    ws = rw_data.get_new_sheet(wb, "sheet 1")
    count = 0
    for i in range(1, rw_data.get_row_number(table)):
        if is_bd(table, i) and is_pass(table, i):
            rw_data.copy_all_row_data(ws, table, count, i)
            count = count + 1
    wb.save(setup_env.FILTER_DATA)
    setup_env.display_message(u"Create tmp file new.xls ...")

def get_field_table():
    global GRADUATION_CREDIT_THRESHOLD
    table = rw_data.read_excel(setup_env.CONFIG_FILE, 0)
    field_table = {}

    if rw_data.get_cell_value(table, 0, 0) == u"最低總學分門檻" and table.cell(0,1).value:
        GRADUATION_CREDIT_THRESHOLD = int(table.cell(0,1).value)
    if rw_data.get_cell_value(table, 1, 1) == u"數學及基礎科學" and rw_data.get_cell_value(table, 1, 2) == u"工程專業課程":
        for i in range(2, table.nrows):
            field_table[table.cell(i,0).value] = [int(table.cell(i,1).value), int(table.cell(i,2).value)]
    else:
        raise Exception(u"Config file is wrong.")
    return field_table

def get_credit_in_field(title, ori_credit, field_table):

    if title in field_table:
        credit = copy.deepcopy(field_table[title])
        if re.search(u"實習", title) or re.search(u"專題", title):
            credit.extend(["V",None])
        else:
            credit.extend([None, None])
        credit = [i if i!=0 else None for i in credit]
    else:
        credit = [None, None, None, ori_credit]

    return credit

def change_grade_format(register_year, ori_grade):
    change_content = {0:u"一", 1:u"二", 2:u"三", 3:u"四", 4:u"五", 5:u"六", "01":u"上", "02":u"下", "0h":u"暑假"}
    grade = change_content[int(ori_grade[0:-2]) - register_year]
    semester = change_content[ori_grade[-2:]]
    return grade + semester

def is_take_project(title):
    return True if title == u"實務專題" else False

def get_data(table, row, register_year, field_table):
    data = []
    for col in range(5):
        data.append(rw_data.get_cell_value(table, row, col))

    ori_grade = data[0]
    ori_credit = int(data[4])
    data[0] = change_grade_format(register_year, ori_grade)
    data[4] = get_credit_in_field(data[1], ori_credit, field_table)

    return data

def get_register_year(sid):
    return int(sid[1:-5])

def get_sid_and_line_number(table):
    sid_and_ln = {}

    for i in range(table.nrows):
        sid = rw_data.get_cell_value(table, i, 6)

        if not sid in sid_and_ln:
            sid_and_ln[sid] = [i]
        else:
            sid_and_ln[sid].append(i)

    return sid_and_ln

def cmp_data():
    field_table = get_field_table()
    table = rw_data.read_excel(setup_env.FILTER_DATA, 0)
    sid_and_ln = get_sid_and_line_number(table)

    for sid in list(sid_and_ln.keys()):
        wb = rw_data.get_init_excel()
        ws = rw_data.get_new_sheet(wb, "sheet 1")
        rw_data.create_title(ws)
        register_year = get_register_year(sid)
        is_project = False

        for i, ori_i in enumerate(sid_and_ln[sid], start = 2):
            data = get_data(table, ori_i, register_year, field_table)
            rw_data.write_processed_data(ws, i, data)
            if not is_project:
                is_project = is_take_project(data[1])

        rw_data.write_result(ws, i+1, is_project)
        wb.save(setup_env.RESULT_FOLDER + "\\" + str(sid) + ".xls")
        setup_env.display_message(u"Create "+ str(sid) +".xls ...")
