import xlrd
import os
import tkinter as tk
import pandas as pd
import openpyxl as ox
import xlwings as xw
import re
from tkinter import ttk, filedialog





class VocabularyApp:
    def __init__(self, master, vocabulary):
        self.master = master
        self.vocabulary = vocabulary
        self.labels = []
        self.combos = []
        self.not_ready = True
        self.create_widgets()
    
    def is_ready(self):
        return self.not_ready

    def get_vocubalary(self):
        return self.vocabulary

    def create_widgets(self):
        # Создаем Canvas контейнер
        canvas = tk.Canvas(self.master, borderwidth=0, highlightthickness=0, width=800, height=800)
        canvas.pack(side="left", fill="both", expand=True)

        # Создаем Frame внутри Canvas контейнера
        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")
        vsb = tk.Scrollbar(self.master, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        # Создаем метки с ключами словаря и comboboxes для выбора значений
        for i, (key, value) in enumerate(self.vocabulary.items()):
            label = tk.Label(frame, text=key)
            label.grid(row=i, column=0, padx=5, pady=5)
            combo = ttk.Combobox(frame, width=75, values=[
                                    "Начальные классы",
                                    "Русский язык и литература",
                                    "Иностранный язык",
                                    "Математика и информатика",
                                    "Общественные науки",
                                    "Естественные науки",
                                    "Технология",
                                    "Искусство",
                                    "Физическая культура, ОБЖ",
                                    "Курсы по выбору"
                                ])
            combo.set("Выберите предметную область, к которой относится предмет, написаный слева")
            combo.grid(row=i, column=1, padx=5, pady=5)
            self.labels.append(label)
            self.combos.append(combo)

        # Создаем кнопку подтверждения выбора
        self.button = tk.Button(frame, text="Подтвердить", command=self.confirm_selection)
        self.button.grid(row=len(self.vocabulary), column=1, padx=5, pady=5)

        self.button_break = tk.Button(frame, text="Выйти, не выбирая предметных областей", command=self.just_exit)
        self.button_break.grid(row=len(self.vocabulary) + 1, column=1, padx=0, pady=5)


        # Добавляем возможность прокручивания Canvas контейнера с помощью колесика мыши
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Устанавливаем размеры Canvas контейнера
        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
    
    def just_exit(self):
        self.master.destroy()
        self.not_ready = False

# Функция для подтверждения выбора и присвоения значений словаря
    def confirm_selection(self):
        all_selected = True
        for i, (key, value) in enumerate(self.vocabulary.items()):
            area = self.combos[i].get()
            if area != "Выберите предметную область, к которой относится предмет, написаный слева":
                self.vocabulary[key] = area
            else:
                all_selected = False
        
        if all_selected:
            self.master.destroy()
            self.not_ready = False
        else:
            self.button['text'] = "Выберите предметную область для каждого предмета. И нажмите сюда снова"



class DF_parser:
    
    # subj_area = {'Русский язык': 'Русский язык и литература', 'Литература' : 'Русский язык и литература'}

    @classmethod
    def clear_df(cls, df_input: pd.DataFrame):
        cls.df_input = df_input.dropna(how='all')
        cls.df_input = df_input.dropna(axis=1, how='all')
        #print(df_input.columns)

    @classmethod
    def get_grades(cls, df_input: pd.DataFrame):

        cls.grades_set = set()
        cls.grades_list = []

        for column_name in df_input.columns:
        # используем регулярное выражение для поиска цифр и русских букв в названии столбца
            if re.search(r'\d.*[а-яА-Я]', column_name):
                # если столбец удовлетворяет условию, добавляем его имя в множество
                cls.grades_set.add(column_name)
                cls.grades_list.append(column_name)

    @classmethod
    def parse(cls, df_input: pd.DataFrame, df_output: pd.DataFrame, file_in, file_out):


        # открываем файл
        wb = xw.Book(file_out)

        # выбираем лист
        sheet_output = wb.sheets("Расстановка")

        sheet_output.range('A1').number_format = '@'
        cls.clear_df(df_input)
        cls.get_grades(df_input)

        columns = list(cls.df_input.columns)

        last_col_index = max([i for i, col in enumerate(columns) if any(c.isalpha() for c in col) and any(c.isdigit() for c in col)])


        cls.df_input = cls.df_input.iloc[:, :last_col_index + 1]


        teachers = [row for index, row in cls.df_input.iterrows() if not pd.isna(row[1])]
        teachers_low_grade = []
        teachers_high_grade = []
        subjects = dict()


        grades_under_five = 0
        for column_name in cls.df_input.columns:
            match = re.search(r'\d+', column_name)  # Извлечение числа из строки
            if match is not None and int(match.group()) <= 4:
                grades_under_five+=1

        # for teacher in teachers:
        #     for i in range(2, grades_under_five + 2):
        #         if str(teacher[i]) != 'nan':
        #             teachers_low_grade.append(teacher)
        #             break

        for teacher in teachers:
            # if str(teacher[1]) != 'Начальная школа':
            #     teachers_high_grade.append(teacher)
            if teacher[1] not in subjects:
                subjects[teacher[1]] = 'nan'

        not_ready = True
        while not_ready:
            root = tk.Tk()
            app = VocabularyApp(root, subjects)
            root.mainloop()
            not_ready = app.is_ready()
        subjects = app.get_vocubalary()


        teachers = sorted(teachers, key=(lambda x, subjects = subjects: subjects[x[1]]))



        # print(cls.df_input.columns)        
        # print(cls.grades_set)
        # print(cls.grades_list)


        for grade in range(len(cls.grades_list)):
            sheet_output.cells(1, 6 + grade).value = cls.grades_list[grade]

        # sheet_output.cells(4, 1).value = "Начальная школа"

        current_row = 4
        current_col = 2

        teacher_prev_name = 'null'
        subject_area_prev_name = 'null'

        for teacher in teachers:
            if subject_area_prev_name != subjects[teacher[1]]:
                sheet_output.cells(current_row, 1).value = subjects[teacher[1]]
                current_row += 1
    
            if teacher[0] != teacher_prev_name:
                sheet_output.cells(current_row, current_col).value = teacher[0]
            sheet_output.cells(current_row, current_col + 1).value = teacher[1]

            for item in teacher[2:]:
                current_col+=1
                sheet_output.cells(current_row, current_col + 3).number_format = '@'
                sheet_output.cells(current_row, current_col + 3).value = item

            teacher_prev_name = teacher[0]
            subject_area_prev_name = subjects[teacher[1]]
            current_col = 2
            current_row+=1

        # teacher_prev_name = 'null'
        # for teacher in teachers_low_grade:
        #     if teacher[0] != teacher_prev_name:
        #         sheet_output.cells(current_row, current_col).value = teacher[0]
        #     sheet_output.cells(current_row, current_col + 1).value = teacher[1]

        #     for item in teacher[2:]:
        #         current_col+=1
        #         sheet_output.cells(current_row, current_col + 3).number_format = '@'
        #         sheet_output.cells(current_row, current_col + 3).value = item
        #     teacher_prev_name = teacher[0]
        #     current_col = 2
        #     current_row+=1

        # sheet_output.cells(current_row, 1).value = "Средняя и старшая школа" 
        # current_row+=1

        # teacher_prev_name = 'null'
        # for teacher in teachers_high_grade:
        #     if teacher[0] != teacher_prev_name:
        #         sheet_output.cells(current_row, current_col).value = teacher[0]
        #     sheet_output.cells(current_row, current_col + 1).value = teacher[1]
        #     for item in teacher[2:]:
        #         current_col+=1
        #         sheet_output.cells(current_row, current_col + 3).number_format = '@'
        #         sheet_output.cells(current_row, current_col + 3).value = item

        #     teacher_prev_name = teacher[0]
        #     current_col = 2
        #     current_row+=1

        wb.save(file_out)

# DF_parser.parse(pd.read_excel("C:/Users/User/Downloads/upupu/5saplication/Школа79.xls"), 
#                 pd.read_excel("C:/Users/User/Downloads/upupu/5saplication/Расстановка.xlsm"), 
#                 "C:/Users/User/Downloads/upupu/5saplication/Школа79.xls",
#                 "C:/Users/User/Downloads/upupu/5saplication/Расстановка.xlsm")
