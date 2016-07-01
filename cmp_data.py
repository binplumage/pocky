#-*- coding: utf-8 -*-
import setup_env
import rw_data
import re

def count_student_number(table, num_rows):
    STUDENT_ID_COL = 7
    student_id = table.cell(1,STUDENT_ID_COL).value
    # First row is title.
    split_point = [2]
    for i in range(1, num_rows):
        next_id = table.cell(i,STUDENT_ID_COL).value
        if not next_id is student_id:
            student_id = next_id
            slit_point.append(i)
    return split_point

def split_data(file_name):
    table = rw_data.read_excel(file_name, 0)
    row_num = rw_data.get_row_number(table)
    col_num = rw_data.get_col_number(table)
    split_point = count_student_number(table, row_num)
    ws = rw_data.get_init_excel()
