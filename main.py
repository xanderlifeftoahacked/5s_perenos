from par_ser import*
import urllib.request
import requests
import time

# currentVersion = "1.0.0"
# URL = urllib.request.urlopen('https://example.com/yourapp/version.html')

# data = URL.read()
# if (data == currentVersion):
#     print("App is up to date!")
# else:
#     print("App is not up to date! App is on version " + currentVersion + " but could be on version " + data + "!")
#     print("Downloading new version now!")
#     newVersion = requests.get("https://github.com/yourapp/app-"+data+".exe")
#     open("Перенос в тарификацию.exe", "wb").write(newVersion.content)
#     print("New version downloaded, restarting in 5 seconds!")
#     time.sleep(5)
#     quit()
# import time



class Start_Window(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style(self.master)
        style.configure('TLabel', font=('Arial', 14))
        style.configure('TButton', font=('Arial', 14))

        self.file1_label = ttk.Label(self, text="Выберите вашу программу")
        self.file1_label.pack(side="top", padx=10, pady=5)

        self.file1_button = ttk.Button(self, text="Ника", command=self.selected_nika)
        self.file1_button.pack(side="top", padx=10, pady=5)

        self.file2_button = ttk.Button(self, text="TimeTables", command=self.selected_timetables)
        self.file2_button.pack(side="top", padx=10, pady=5)

    def selected_nika(self):
        self.master.destroy()
        root = tk.Tk()
        root.title("Ника")
        app = Application_2(master=root)
        app.mainloop()

    def selected_timetables(self):
        #self.master.destroy()
        #root = tk.Tk()
        root.title("Timetables")
        app = Application_in_proccess(parent=self)
        app.mainloop()

class Application_in_proccess(tk.Toplevel):
        def __init__(self, parent):
            super().__init__(parent)
            self.title("В процессе разработки")
            self.geometry("400x120")
            self.resizable(False, False)
            ttk.Label(self, text="Данная функция в процессе разработки.").pack(pady=20)
            ttk.Button(self, text="Закрыть", command=self.destroy).pack(pady=10)


class Application_2(tk.Frame):
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
        app = Application_2(master=root)
        app.mainloop()

    def create_widgets(self):
        # создаем объекты стиля
        style = ttk.Style(self.master)
        style.configure('TLabel', font=('Arial', 14))
        style.configure('TButton', font=('Arial', 14))



        # добавляем метки для вывода названий файлов
    
        file_frame1 = ttk.Frame(self)
        self.file1_label = ttk.Label(file_frame1, text="Файл, из программы 'Ника', содержащий информацию о нагрузке учителей: ")
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

        # добавляем кнопку для подтверждения выбора
        

    def select_file1(self):
        # открываем диалоговое окно выбора первого файла
        filename = filedialog.askopenfilename()
        if filename:
            # выводим название выбранного файла в соответствующее поле
            self.file_input_path.set(filename)
        # self.file1_button.config(state=tk.DISABLED)

    def select_file2(self):
        # открываем диалоговое окно выбора второго файла
        filename = filedialog.askopenfilename()
        if filename:
            # выводим название выбранного файла в соответствующее поле
            self.file_output_path.set(filename)
            self.df_out = pd.read_excel(filename, sheet_name='Расстановка')
            self.df_out = pd.DataFrame(self.df_out)
        # self.file2_button.config(state=tk.DISABLED)
        self.confirm_button = ttk.Button(self, text="Подтвердить. Внимание, проверьте правильность указанных файлов!", command=self.work_with_path)
        self.confirm_button.pack(side="top", padx=10, pady=20)

    def work_with_path(self):
        self.confirm_button.config(state=tk.DISABLED)
        self.file1_button.config(state=tk.DISABLED)
        self.file2_button.config(state=tk.DISABLED)
        self.confirm_button = ttk.Button(self, text="Выбрать файлы заново (нажимать при выборе неправильных файлов)", command=self.restart)
        self.confirm_button.pack(side="top", padx=10, pady=20)
        # получаем путь к выбранной Excel таблице
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

        # добавляем кнопку для подтверждения выбора листа
        attention_label = ttk.Label(self, text="Выберите лист, содержащий столбец с ФИО, столбец с предметом и столбцы с названиями классов (по умолчанию - тарификация)")
        attention_label.pack(side="top", padx=10, pady=5)
 
        confirm_sheet_button = ttk.Button(self, text="Выбрать лист", command=lambda: self.confirm_sheet(sheet_name_combobox, filename, confirm_sheet_button))
        confirm_sheet_button.pack(side="top", padx=10, pady=5)

    def confirm_sheet(self, sheet_name_combobox, filename, button):
        button.config(state=tk.DISABLED)
        sheet_name = sheet_name_combobox.get()

        # загружаем данные из выбранного листа
        self.df_in = pd.read_excel(filename, sheet_name=sheet_name)

        # создаем DataFrame из загруженных данных
        self.df_in = pd.DataFrame(self.df_in)
        # выводим таблицу в виджет
        # self.table = tk.Frame(self)
        # self.table.pack(side="top", padx=10, pady=10)

        # comboboxes = []

        # for i, col in enumerate(df.columns):
        #     # создаем заголовок столбца таблицы и размещаем его на форме
        #     label = ttk.Label(self.table, text=col, font=('Arial', 12, 'bold'))
        #     label.grid(row=0, column=i)
        #     # создаем ttk.Entry для каждого столбца таблицы и размещаем его на форме
        #     entry_var = tk.StringVar(value=col)
        #     # entry = ttk.Entry(self.table, textvariable=entry_var)
        #     # entry.grid(row=1, column=i)
        #     # создаем ttk.Combobox для каждого столбца таблицы и размещаем его на форме
        #     combobox_var = tk.StringVar(value=col)
        #     combobox = ttk.Combobox(self.table, textvariable=combobox_var, state='readonly', values=["Имя учителя", "Название класса", "Название предмета", "И ТД"])
        #     combobox.grid(row=2, column=i)
        #     # добавляем обработчик события для изменения названия столбца при выборе нового значения в ttk.Combobox
        #     combobox.bind('<<ComboboxSelected>>', lambda event, col=col, combobox=combobox_var, entry=entry_var: self.change_column_name(df, col, combobox.get(), combobox, entry))
        #     comboboxes.append(combobox)
        # # функция для изменения названия столбца таблицы



        # выводим название выбранной Excel таблицы и листа
        self.sheet_label = ttk.Label(self, text=f"Выбранный лист: {sheet_name}")
        self.sheet_label.pack(side="top", padx=10, pady=5)
        self.filename_label = ttk.Label(self, text=f"Выбранный файл: {filename}")
        self.filename_label.pack(side="top", padx=10, pady=5)
        # self.filename_label = ttk.Label(self, text=f"Установите соответствие между названиями столбцов в вашем файле и предложенными названиями\n При неправильном выборе листа или файла перезапустите программу")
        self.filename_label.pack(side="top", padx=10, pady=5)

        # создаем кнопку для применения изменений
        apply_button = ttk.Button(self, text="Начать перенос", command=lambda: self.parse_dataframe(self.df_in, self.df_out)) #command=lambda: self.apply_changes(apply_button, comboboxes)
        apply_button.pack(side="top", padx=10, pady=10)

    # def change_column_name(self, df, old_name, new_name, combobox, entry):
    #     # заменяем название столбца в DataFrame
    #     df.rename(columns={old_name: new_name}, inplace=True)
    #     # обновляем значения ttk.Entry и ttk.Combobox
    #     combobox.set(new_name)
    #     entry.set(new_name)
    #     print(df.columns)

    def apply_changes(self, button, comboboxes):
        button.config(state=tk.DISABLED)
        for combobox in comboboxes:
            combobox.config(state = tk.DISABLED)


    
    def parse_dataframe(self, df_input: pd.DataFrame, df_output: pd.DataFrame):
        self.master.destroy()
        _in = self.file_input_path.get()
        _out = self.file_output_path.get()
        
        DF_parser.parse(df_input, df_output, _in, _out)



if __name__ == '__main__':
    root = tk.Tk()
    app = Start_Window(master=root)
    app.mainloop()
    

# import urllib.request
# import os
# import subprocess
# import tempfile

# # Set the URL of the file containing the latest version number
# version_url = 'https://example.com/version.txt'

# # Set the path to the configuration file containing the current version number
# config_file_path = 'config.ini'

# # Retrieve the latest version number from the remote source
# with urllib.request.urlopen(version_url) as response:
#     latest_version = response.read().decode('utf-8').strip()

# # Read the current version number from the configuration file
# with open(config_file_path, 'r') as f:
#     current_version = f.read().strip()

# # Compare the current version number with the latest version number
# if latest_version > current_version:
#     # Set the URL of the updated version of your application
#     update_url = 'https://example.com/update.zip'
    
#     # Download the updated version of your application from the remote source
#     update_file_path = os.path.join(tempfile.gettempdir(), 'update.zip')
#     urllib.request.urlretrieve(update_url, update_file_path)
    
#     # Unzip the update file to a temporary directory
#     update_dir_path = os.path.join(tempfile.gettempdir(), 'update')
#     subprocess.run(['unzip', update_file_path, '-d', update_dir_path])
    
#     # Run the update script to install the updated version of your application
#     update_script_path = os.path.join(update_dir_path, 'update.py')
#     subprocess.run(['python', update_script_path])
    
#     # Update the version number in the configuration file
#     with open(config_file_path, 'w') as f:
#         f.write(latest_version)
