#-*- coding: utf-8 -*-
import os
import datetime

tmp_folder = ""

def make_sure_folder_exists(folder):
    if not os.path.exists(folder):
        os.makedirs(folder)
        display_message("Create folder: " + folder)

def get_time():
    return datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

def set_environment():
    global tmp_folder
    SCRIPT_DIR = os.getcwd()
    NOW_TIME = get_time()
    tmp_folder = SCRIPT_DIR + "\\tmp"
    result_folder = SCRIPT_DIR + "\\" + NOW_TIME
    make_sure_folder_exists(tmp_folder)
    make_sure_folder_exists(result_folder)
    display_message("Setup Environment successful.")

def display_message(mes):
    time = get_time()
    print time + " " + mes

def clean_envirnoment():
    pass
