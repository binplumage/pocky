#-*- coding: utf-8 -*-
import setup_env
import rw_data
import re

def get_split_point(table, num_rows):
    STUDENT_ID_COL = 7
    student_id = table.cell(1,STUDENT_ID_COL).value
    # First row is title.
    split_point = [1]
    for i in range(2, num_rows):
        next_id = table.cell(i,STUDENT_ID_COL).value
        if not next_id is student_id:
            student_id = next_id
            split_point.append(i)
    return split_point

def split_data(file_name):
    table = rw_data.read_excel(file_name, 0)
    row_num = rw_data.get_row_number(table)
    col_num = rw_data.get_col_number(table)
    split_point = get_split_point(table, row_num)
    split_point.append(row_num)

    for s in range(len(split_point)-1):
        wb = rw_data.get_init_excel()
        ws = rw_data.get_new_sheet(wb, "sheet 1")
        for i, ori_i in enumerate(range(split_point[s], split_point[s+1])):
            sid = table.cell(ori_i, 7).value.rstrip()
            if sid[0] == 'B':
                rw_data.write_all_row_data(ws, table, i, ori_i)

        wb.save(setup_env.TMP_FOLDER + "\\" + str(sid) + ".xls")
        setup_env.display_message(u"Create "+ str(sid) +".xls tmp file...")
    setup_env.display_message(u"Split finish...")
