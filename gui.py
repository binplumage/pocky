#-*- coding: utf-8 -*-
import Tkinter as tk
import ttk
import tkMessageBox
import tkFileDialog
from ScrolledText import ScrolledText
import sys
import setup_env
import cmp_data
import rw_data
import tkFont

class MainApplication():
    def __init__(self, master):

        self.filename = ""
        self.confg = ""
        self.master = master
        #setting color.
        bkc = "floral white"
        fgc = "RosyBrown4"
        bgc = "misty rose"

        self.FontForButton = tkFont.Font(family="Verdana", size=12)
        self.FontForLabel = tkFont.Font(family="Verdana", weight="bold")
        self.frame = tk.Frame(self.master, bg=bkc)
        self.frame.pack()
        self.chose_button = tk.Button(self.frame, text = u"選擇檔案", command = self.open_file, font=self.FontForButton, width = 20, bg=bgc, fg=fgc)
        self.chose_button.pack(padx = 5, pady = 10, side = tk.TOP)
        self.chose_confg_button = tk.Button(self.frame, text = u"選擇設定檔", command = self.get_confg, font=self.FontForButton, width = 20, bg=bgc, fg=fgc)
        self.chose_confg_button.pack(padx = 5, pady = 10, side = tk.TOP)
        self.run_button = tk.Button(self.frame, text = u"執行", command = self.run, font=self.FontForButton, width = 20, bg=bgc, fg=fgc)
        self.run_button.pack(padx = 5, pady = 10, side = tk.TOP)
        self.text = ScrolledText(self.frame)
        self.text.pack()
        self.mdby = tk.Label(self.frame, text="\nPowered By MITLab", font=self.FontForLabel, fg="SkyBlue1", bg=bkc)
        self.mdby.pack(side='bottom')

    def open_file(self):
        self.filename = tkFileDialog.askopenfilename()
        if self.filename :
            setup_env.display_message(u"選擇檔案: " + self.filename)

    def get_confg(self):
        self.confg = tkFileDialog.askopenfilename()
        if self.confg :
            setup_env.display_message(u"選擇設定檔案: " + self.confg)

    def write(self, massage):
        self.text.insert(tk.END, massage)
        self.text.see(tk.END)
        self.text.update_idletasks()#display message real time

    def run(self):
        try:
            if not self.filename or not self.confg:
                raise Exception('請選擇檔案!')
            setup_env.display_message(u"開始執行...")
            setup_env.set_environment(self.confg)
            table = rw_data.read_excel(self.filename, 0)
            rw_data.get_title_col(table)
            cmp_data.filter_data(table)
            cmp_data.cmp_data()
        except Exception as e:
            setup_env.display_message(e.message)
        finally:
            setup_env.clean_envirnoment()

def create_windows():
    root = tk.Tk()
    root.title(u"畢業學分核定系統 V1.0.0")
    root.minsize(width = 600, height = 500)
    root.attributes("-toolwindow", 1)
    sys.stdout = MainApplication(root)
    root.mainloop()

create_windows()
