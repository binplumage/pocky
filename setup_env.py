#-*- coding: utf-8 -*-
import os
import shutil
import datetime
import rw_data

SCRIPT_DIR = os.getcwd()
TMP_FOLDER = ""
RESULT_FOLDER = ""
CONFIG_FILE = ""
FILTER_DATA = ""
GRADUATION_CREDIT_THRESHOLD = 128

def create_folder(folder):
    try:
        os.makedirs(folder)
        display_message("Create folder: " + folder)
    except:
        raise Exception("Can not create folder : " + folder)

def delete_folder(folder):
    try:
        shutil.rmtree(folder)
        display_message("Delete folder: " + folder)
    except:
        raise Exception("Can not delte folder : " + folder)

def is_folder_exists(folder):
    return os.path.exists(folder)

def create_new_folder(folder):
    #If folder exists, delete folder and create it.
    if is_folder_exists(folder):
        delete_folder(folder)
    create_folder(folder)

def get_time():
    return datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

def set_environment(confg_file):
    global TMP_FOLDER, SCRIPT_DIR, RESULT_FOLDER, FILTER_DATA, CONFIG_FILE

    now_time = get_time()
    TMP_FOLDER = SCRIPT_DIR + "\\tmp"
    RESULT_FOLDER = SCRIPT_DIR + "\\" + now_time
    create_new_folder(TMP_FOLDER)
    create_folder(RESULT_FOLDER)
    CONFIG_FILE = confg_file
    FILTER_DATA = TMP_FOLDER + "\\new.xls"
    display_message("Setup Environment successful.")

def display_message(mes):
    time = get_time()
    print time + " " + mes

def clean_envirnoment():
    if is_folder_exists(TMP_FOLDER):
        display_message("Clean Environment ...")
        delete_folder(TMP_FOLDER)
        display_message("Clean Environment successful.")
