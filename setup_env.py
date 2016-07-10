#-*- coding: utf-8 -*-
import os
import shutil
import datetime

SCRIPT_DIR = os.getcwd()
TMP_FOLDER = ""
RESULT_FOLDER = ""
FIELD_TABLE_FILE = SCRIPT_DIR + "\\field_table.xlsx"

def create_folder(folder):
    try:
        os.makedirs(folder)
        display_message("Create folder: " + folder)
    except:
        display_message("Can not create folder : " + folder)

def delete_folder(folder):
    try:
        shutil.rmtree(folder)
        display_message("Delete folder: " + folder)
    except:
        display_message("Can not delte folder : " + folder)

def make_sure_folder_exists(folder):
    # If folder exists, delete folder and create it.
    if os.path.exists(folder):
        delete_folder(folder)
    create_folder(folder)

def get_time():
    return datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

def set_environment():
    global TMP_FOLDER, SCRIPT_DIR, RESULT_FOLDER

    now_time = get_time()
    TMP_FOLDER = SCRIPT_DIR + "\\tmp"
    RESULT_FOLDER = SCRIPT_DIR + "\\" + now_time
    make_sure_folder_exists(TMP_FOLDER)
    make_sure_folder_exists(RESULT_FOLDER)
    display_message("Setup Environment successful.")

def display_message(mes):
    time = get_time()
    print time + " " + mes

def clean_envirnoment():
    pass
