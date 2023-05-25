from ParserForNika import*
from ApplicationNika import*
from ApplicationTimetables import*
import threading
import urllib.request

current_version = 10
new_app_url = "https://github.com/xanderlifeftoahacked/5s_perenos/releases/download/new/Perenos.exe"
version_url = "https://raw.githubusercontent.com/xanderlifeftoahacked/5s_perenos/main/version.txt"

class UpdateWindow:
    def __init__(self, parent, real_version, current_version):
        self.parent = parent
        self.real_version = real_version
        self.current_version = current_version

        if self.real_version > self.current_version:
            self.top = tk.Toplevel(parent)
            self.top.title("Доступна новая версия")

            self.label = tk.Label(self.top, text="Идет обновление, пожалуйста, подождите")
            self.label.pack(padx=20, pady=10)

            self.progress = ttk.Progressbar(self.top, mode="indeterminate")
            self.progress.pack(padx=20, pady=10)
            self.progress.start()

            self.thread = threading.Thread(target=self.update_get)
            self.thread.start()

    def update_get(self):
        new_app_filename = "Perenos_v1.1.exe"

        try:
            with urllib.request.urlopen(new_app_url) as url:
                with open(new_app_filename, "wb") as f:
                    f.write(url.read())
            os.remove(__file__)
        except Exception as e:
            print("An error occurred: ", e)
        self.label.config(text="Обновление завершено, закройте все окна приложения")
        self.progress.destroy()

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
        app = Application_Nika(master=root)
        app.mainloop()

    def selected_timetables(self):
        self.master.destroy()
        root = tk.Tk()
        root.title("Timetables")
        app = Application_Timetables(master=root)
        app.mainloop()

class Application_in_proccess(tk.Toplevel):
        def __init__(self, parent):
            super().__init__(parent)
            self.title("В процессе разработки")
            self.geometry("400x120")
            self.resizable(False, False)
            ttk.Label(self, text="Данная функция в процессе разработки.").pack(pady=20)
            ttk.Button(self, text="Закрыть", command=self.destroy).pack(pady=10)


if __name__ == '__main__':
    root = tk.Tk()
    app = Start_Window(master=root)

    try:
        with urllib.request.urlopen(version_url) as f:
            real_version = int(f.read().decode().strip())
        if real_version > current_version:
            app1 = UpdateWindow(root, real_version, current_version)
    except Exception as e:
        print("An error occurred: ", e)

    app.mainloop()

