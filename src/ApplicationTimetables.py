import tkinter as tk
import pandas as pd
import os
import openpyxl as ox
import xlrd
from tkinter import ttk, filedialog
from ParserForTimetables import DF_parser


class Application_Timetables(tk.Frame):
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
        app = Application_Timetables(master=root)
        app.mainloop()

    def create_widgets(self):
        style = ttk.Style(self.master)
        style.configure('TLabel', font=('Arial', 14))
        style.configure('TButton', font=('Arial', 14))
    
        file_frame1 = ttk.Frame(self)
        self.file1_label = ttk.Label(file_frame1, text="Файл, из программы 'Timetables', содержащий информацию о нагрузке учителей: ")
        self.file1_label.pack(side="left", padx=10, pady=5, anchor="w")
        self.file_input_path = tk.StringVar()
        self.file1_entry = ttk.Entry(file_frame1, textvariable=self.file_input_path, state='readonly', width=50)
        self.file1_entry.pack(side="left", padx=10, pady=5, anchor="e")
        self.file1_button = ttk.Button(file_frame1, text="Выбрать", command=self.select_file1)
        self.file1_button.pack(side="right", padx=10, pady=5)
        
        file_frame1.pack(side="top", padx=10, pady=5, anchor='e')
        file_frame2 = ttk.Frame(self)
        self.file2_label = ttk.Label(file_frame2, text="Растановка.xlsm:")
        self.file2_label.pack(side="left", padx=10, pady=5, anchor="w")
        self.file_output_path = tk.StringVar()
        self.file2_entry = ttk.Entry(file_frame2, textvariable=self.file_output_path, state='readonly', width=50)
        self.file2_entry.pack(side="left", padx=10, pady=5, anchor="e")
        self.file2_button = ttk.Button(file_frame2, text="Выбрать", command=self.select_file2)
        self.file2_button.pack(side="right", padx=10, pady=5)

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
        self.confirm_button.pack(side="top", padx=10, pady=20)

    def work_with_path(self):
        self.confirm_button.config(state=tk.DISABLED)
        self.file1_button.config(state=tk.DISABLED)
        self.file2_button.config(state=tk.DISABLED)
        self.confirm_button = ttk.Button(self, text="Выбрать файлы заново (нажимать при выборе неправильных файлов)", command=self.restart)
        self.confirm_button.pack(side="top", padx=10, pady=20)

        filename = self.file_input_path.get()
        sheets = []

        if not filename:
            return
        
        sheet_name_var = tk.StringVar(value="Лист1")
        sheet_name_combobox = ttk.Combobox(self, textvariable=sheet_name_var, state='readonly')

        if os.path.splitext(filename)[1] != '.xlsx':
            workbook_xls = xlrd.open_workbook(filename)
            sheets = workbook_xls.sheet_names()
        else:
            workbook_xlsx = ox.load_workbook(filename=filename, read_only=True, data_only=True) #* data_only - ?
            sheets = workbook_xlsx.sheetnames

        sheet_name_combobox['values'] = sheets
        sheet_name_combobox.pack(side="top", padx=10, pady=5)

        attention_label = ttk.Label(self, text="Выберите лист, содержащий столбец с ФИО, столбец с предметом и столбцы с названиями классов\nВнимание, столбцы без названия учтены не будут! ")
        attention_label.pack(side="top", padx=10, pady=5)
 
        confirm_sheet_button = ttk.Button(self, text="Выбрать лист", command=lambda: self.confirm_sheet(sheet_name_combobox, filename, confirm_sheet_button))
        confirm_sheet_button.pack(side="top", padx=10, pady=5)

    def confirm_sheet(self, sheet_name_combobox, filename, button):
        button.config(state=tk.DISABLED)
        sheet_name = sheet_name_combobox.get()

        self.df_in = pd.read_excel(filename, sheet_name=sheet_name)

        self.df_in = pd.DataFrame(self.df_in)
        self.df_in = self.df_in.loc[:, ~self.df_in.columns.str.contains('^Unnamed')]

        self.table = tk.Frame(self)
        self.table.pack(side="top", padx=10, pady=10)

        comboboxes = []

        for i, col in enumerate(self.df_in.columns):
            label = ttk.Label(self.table, text=col, font=('Arial', 12, 'bold'))
            label.grid(row=0, column=i)
            entry_var = tk.StringVar(value=col)
            combobox_var = tk.StringVar(value="Не выбрано")
            combobox = ttk.Combobox(self.table, textvariable=combobox_var, state='readonly', values=["ФИО учителя", "Класс", "Группа", "Количество детей" ,"Предмет", "Количество уроков", "Другое"])
            combobox.grid(row=2, column=i)
            combobox.bind('<<ComboboxSelected>>', lambda event, col=col, combobox=combobox_var, entry=entry_var: self.change_column_name(self.df_in, col, combobox.get(), combobox, entry))
            comboboxes.append(combobox)

        self.sheet_label = ttk.Label(self, text=f"Выбранный лист: {sheet_name}")
        self.sheet_label.pack(side="top", padx=10, pady=5)
        self.filename_label = ttk.Label(self, text=f"Выбранный файл: {filename}")
        self.filename_label.pack(side="top", padx=10, pady=5)
        self.filename_label = ttk.Label(self, text=f"Установите соответствие между названиями столбцов в вашем файле и предложенными названиями\nСтолбцы с пометкой 'Другое' в переносе использованы не будут\nПри неправильном выборе листа или файла перезапустите программу")
        self.filename_label.pack(side="top", padx=10, pady=5)

        apply_button = ttk.Button(self, text="Начать перенос", command=lambda: self.parse_dataframe(self.df_in, self.df_out)) #command=lambda: self.apply_changes(apply_button, comboboxes)
        apply_button.pack(side="top", padx=10, pady=10)

    def change_column_name(self, df, old_name, new_name, combobox, entry):
        df.rename(columns={old_name: new_name}, inplace=True)
        combobox.set(new_name)
        entry.set(new_name)
        print(df.columns)

    def apply_changes(self, button, comboboxes):
        button.config(state=tk.DISABLED)
        for combobox in comboboxes:
            combobox.config(state = tk.DISABLED)

    def parse_dataframe(self, df_input: pd.DataFrame, df_output: pd.DataFrame):
        self.master.destroy()
        _in = self.file_input_path.get()
        _out = self.file_output_path.get()
        
        DF_parser.parse(df_input, df_output, _in, _out)
