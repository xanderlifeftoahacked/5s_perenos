import json
import tkinter as tk
from tkinter import ttk

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
        canvas = tk.Canvas(self.master, borderwidth=0,
                           highlightthickness=0, width=800, height=800)
        canvas.pack(side="left", fill="both", expand=True)
        try:
            with open("subject_areas.json", "r") as infile:
                saved_vocabulary = json.load(infile)
                for key, value in self.vocabulary.items():
                    if key in saved_vocabulary:
                        self.vocabulary[key] = saved_vocabulary[key]
                    if key.lower() in saved_vocabulary:
                        self.vocabulary[key] = saved_vocabulary[key.lower()]

        except FileNotFoundError:
            pass

        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")
        vsb = tk.Scrollbar(self.master, orient="vertical",
                           command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

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
            if value != 'nan':
                combo.set(value)
            else:
                combo.set("Выберите предметную область, к которой относится предмет, написаный слева")
            combo.grid(row=i, column=1, padx=5, pady=5)
            self.labels.append(label)
            self.combos.append(combo)

        self.button = tk.Button(frame, text="Подтвердить",
                                command=self.confirm_selection)
        self.button.grid(row=len(self.vocabulary), column=1, padx=5, pady=5)
        self.button_break = tk.Button(
            frame, text="Выйти, не выбирая предметных областей", command=self.just_exit)
        self.button_break.grid(row=len(self.vocabulary) +
                               1, column=1, padx=0, pady=5)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    def just_exit(self):
        self.master.destroy()
        self.not_ready = False

    def confirm_selection(self):
        all_selected = True
        for i, (key, value) in enumerate(self.vocabulary.items()):
            area = self.combos[i].get()
            if area != "Выберите предметную область, к которой относится предмет, написаный слева":
                self.vocabulary[key] = area
            else:
                all_selected = False

        if all_selected:
            with open("subject_areas.json", "w") as outfile:
                json.dump(self.vocabulary, outfile)
            self.master.destroy()
            self.not_ready = False
        else:
            self.button['text'] = "Выберите предметную область для каждого предмета. И нажмите сюда снова"
