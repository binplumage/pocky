#-*- coding: utf-8 -*-
import Tkinter as tk
import ttk
import tkMessageBox
import tkFileDialog
import sys
import setup_env
import cmp_data


class MainApplication():
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.run_button = tk.Button(self.frame, text = u"執行", width = 10, command = self.run)
        self.chose_button = tk.Button(self.frame, text = u"選擇檔案", width = 10, command = self.open_file)
        self.chose_button.pack(padx = 5, pady = 10, side = tk.TOP)
        self.run_button.pack(padx = 5, pady = 10, side = tk.BOTTOM)
        self.filename = ""
        self.text = tk.Text(self.master)
        self.frame.pack()
        self.text.pack()

    def open_file(self):
        self.filename = tkFileDialog.askopenfilename()
        if self.filename :
            setup_env.display_message(u"選擇檔案: " + self.filename)

    def write(self, massage):
        self.text.insert(tk.END, massage)
        self.text.see(tk.END)

    def run(self):
        setup_env.display_message(u"開始執行...")
        setup_env.set_environment()
        cmp_data.split_data(self.filename)

def create_windows():
    root = tk.Tk()
    root.title(u"工程認證 B1.0.0.01")
    root.minsize(width = 600, height = 500)
    sys.stdout = MainApplication(root)
    root.mainloop()

create_windows()
