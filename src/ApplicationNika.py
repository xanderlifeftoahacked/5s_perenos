import tkinter as tk
import openpyxl as ox
import xlwings as xw
import xlrd
import pandas as pd
import os
from ParserForNika import DF_parser
from tkinter import ttk, filedialog

class Application_Nika(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def restart(self):
        self.destroy()
        self.master.destroy()
        root = tk.Tk()
        root.title("Ника")
        app = Application_Nika(master=root)
        app.mainloop()

    def create_widgets(self):
        style = ttk.Style(self.master)
        style.configure('TLabel', font=('Arial', 13))
        style.configure('TButton', font=('Arial', 13))
    
        file_frame1 = ttk.Frame(self)
        self.file0_label = ttk.Label(file_frame1, text="Файл, из программы 'Ника', содержащий информацию о нагрузке учителей: ")
        self.file0_label.pack(side="left", padx=10, pady=5, anchor="w")
        self.file_input_path = tk.StringVar()
        self.file0_entry = ttk.Entry(file_frame1, textvariable=self.file_input_path, state='readonly', width=50)
        self.file0_entry.pack(side="left", padx=10, pady=5, anchor="e")
        self.file0_button = ttk.Button(file_frame1, text="Выбрать", command=self.select_file1)
        self.file0_button.pack(side="right", padx=10, pady=5)
        
        file_frame1.pack(side="top", padx=10, pady=5, anchor='e')
        file_frame2 = ttk.Frame(self)
        self.file1_label = ttk.Label(file_frame2, text="Растановка.xlsm:")
        self.file1_label.pack(side="left", padx=10, pady=5, anchor="w")
        self.file_output_path = tk.StringVar()
        self.file1_entry = ttk.Entry(file_frame2, textvariable=self.file_output_path, state='readonly', width=50)
        self.file1_entry.pack(side="left", padx=10, pady=5, anchor="e")
        self.file1_button = ttk.Button(file_frame2, text="Выбрать", command=self.select_file2)
        self.file1_button.pack(side="right", padx=10, pady=5)

        file_frame2.pack(side="top", padx=10, pady=5, anchor='e')
        

    def select_file1(self):
        filename = filedialog.askopenfilename()
        if filename:
            self.file_input_path.set(filename)

    def select_file2(self):
        filename = filedialog.askopenfilename()
        if filename:
            self.file_output_path.set(filename)
            self.df_out = pd.read_excel(filename, sheet_name='Расстановка')
            self.df_out = pd.DataFrame(self.df_out)

        self.confirm_button = ttk.Button(self, text="Подтвердить. Внимание, проверьте правильность указанных файлов!", command=self.work_with_path)
        self.confirm_button.pack(side="top", padx=9, pady=20)

    def work_with_path(self):
        self.confirm_button.config(state=tk.DISABLED)
        self.file0_button.config(state=tk.DISABLED)
        self.file1_button.config(state=tk.DISABLED)
        self.confirm_button = ttk.Button(self, text="Выбрать файлы заново (нажимать при выборе неправильных файлов)", command=self.restart)
        self.confirm_button.pack(side="top", padx=9, pady=20)

        filename = self.file_input_path.get()
        sheets = []

        if not filename:
            return
        
        sheet_name_var = tk.StringVar(value="Лист0")
        sheet_name_combobox = ttk.Combobox(self, textvariable=sheet_name_var, state='readonly')


        if os.path.splitext(filename)[0] != '.xlsx':
            workbook_xls = xlrd.open_workbook(filename)
            sheets = workbook_xls.sheet_names()

        else:
            workbook_xlsx = ox.load_workbook(filename=filename, read_only=True, data_only=True) #* data_only - ?
            sheets = workbook_xlsx.sheetnames

        sheet_name_combobox['values'] = sheets
        sheet_name_combobox.pack(side="top", padx=9, pady=5)
    
        attention_label = ttk.Label(self, text="Выберите лист, содержащий столбец с ФИО, столбец с предметом и столбцы с названиями классов (по умолчанию - тарификация)")
        attention_label.pack(side="top", padx=9, pady=5)
 
        confirm_sheet_button = ttk.Button(self, text="Выбрать лист", command=lambda: self.confirm_sheet(sheet_name_combobox, filename, confirm_sheet_button))
        confirm_sheet_button.pack(side="top", padx=9, pady=5)

    def confirm_sheet(self, sheet_name_combobox, filename, button):
        button.config(state=tk.DISABLED)
        sheet_name = sheet_name_combobox.get()

        self.df_in = pd.read_excel(filename, sheet_name=sheet_name)
        self.df_in = pd.DataFrame(self.df_in)

        self.sheet_label = ttk.Label(self, text=f"Выбранный лист: {sheet_name}")
        self.sheet_label.pack(side="top", padx=9, pady=5)
        self.filename_label = ttk.Label(self, text=f"Выбранный файл: {filename}")
        self.filename_label.pack(side="top", padx=9, pady=5)
        self.filename_label.pack(side="top", padx=9, pady=5)

        apply_button = ttk.Button(self, text="Начать перенос", command=lambda: self.parse_dataframe(self.df_in, self.df_out)) #command=lambda: self.apply_changes(apply_button, comboboxes)
        apply_button.pack(side="top", padx=9, pady=10)

    def apply_changes(self, button, comboboxes):
        button.config(state=tk.DISABLED)
        for combobox in comboboxes:
            combobox.config(state = tk.DISABLED)


    
    def parse_dataframe(self, df_input: pd.DataFrame, df_output: pd.DataFrame):
        self.master.destroy()
        _in = self.file_input_path.get()
        _out = self.file_output_path.get()
        
        DF_parser.parse(df_input, df_output, _in, _out)

