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


class MainApplication():
    def __init__(self, master):

        self.filename = ""
        self.confg = ""
        self.master = master
        self.frame = tk.Frame(self.master)
        self.frame.pack()
        self.chose_button = tk.Button(self.frame, text = u"選擇檔案", width = 10, command = self.open_file)
        self.chose_button.pack(padx = 5, pady = 10, side = tk.TOP)
        self.chose_confg_button = tk.Button(self.frame, text = u"選擇設定檔", width = 10, command = self.get_confg)
        self.chose_confg_button.pack(padx = 5, pady = 10, side = tk.TOP)
        self.run_button = tk.Button(self.frame, text = u"執行", width = 10, command = self.run)
        self.run_button.pack(padx = 5, pady = 10, side = tk.TOP)

        self.text = ScrolledText(self.master)
        #self.text = tk.Text(self.master)
        self.text.pack()

        #self.scrollb = tk.Scrollbar(self.frame, command=self.text.yview)
        #self.scrollb.grid(row=0, column=1, sticky='nsew')
        #self.text['yscrollcommand'] = self.scrollb.set


    def open_file(self):
        self.filename = tkFileDialog.askopenfilename()
        if self.filename :
            setup_env.display_message(u"選擇檔案: " + self.filename)

    def get_confg(self):
        self.confg = tkFileDialog.askopenfilename()
        if self.confg :
            setup_env.display_message(u"選擇confg檔案: " + self.confg)

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
    root.title(u"工程認證 B1.0.0.01")
    root.minsize(width = 600, height = 500)
    sys.stdout = MainApplication(root)
    root.mainloop()

create_windows()
