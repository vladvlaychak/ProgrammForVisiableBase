from pathlib import Path
import pandas as pd
import tkinter as tk
import pyperclip  # Импортируем библиотеку pyperclip

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"C:/Podchinenost/build/assets/frame0")

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

file_path = "C:/Podchinenost/base.xlsx"

class ProcessTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Подчиненность")
        self.selected_process = None
        self.create_widgets()
        self.load_processes_from_file()

    def on_process_select(self, event):
        selected_index = self.process_listbox.curselection()
        if selected_index:
            selected_index = selected_index[0]  # Получаем первый выбранный индекс
            selected_process_name = self.process_listbox.get(selected_index)
            # Найти выбранный процесс по его имени
            for process_info in self.process_info:
                if process_info["Наименование воинской части"] == selected_process_name:
                    self.selected_process = process_info
                    self.update_info_labels()
                    break

    def load_processes_from_file(self):
        try:
            data = pd.read_excel(file_path, dtype=str)
            # Заменяем отсутствующие значения на пустую строку
            data.fillna("Отсутсвует", inplace=True)
            self.process_info = data.to_dict(orient="records")
            for process_info in self.process_info:
                self.process_listbox.insert(tk.END, process_info["Наименование воинской части"])
        except FileNotFoundError:
            pass

    def create_widgets(self):
        # Список процессов
        self.process_listbox = tk.Listbox(self)
        self.process_listbox.bind("<<ListboxSelect>>", self.on_process_select)
        self.process_listbox.pack(side=tk.LEFT, fill=tk.BOTH, padx=0, pady=30)
        # Метки информации
        self.info_labels = []
        self.labels_info = [
            "Наименование в. ч.: ",
            "Вид (род) войск МО РФ: ",
            "Подчиненность (для МО РФ): ",
            "Принадлежность к данному округу (ЮВО): ",
            "Образец указания в справке к судебному решению: ",
            "Дислокация воинской части: ",
            "Подчиненность: "
        ]

        for label_info in self.labels_info:
            label = tk.Label(self, text=label_info, wraplength=700)  # Устанавливаем максимальную длину строки
            label.pack(padx=5, pady=5, anchor='w')
            self.info_labels.append(label)
        # Фрейм для строки поиска и выбора параметра поиска
        search_frame = tk.Frame(self)
        search_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

        # Строка поиска
        self.search_entry = tk.Entry(search_frame, width=30)
        self.search_entry.pack(pady=8)

        # Кнопка поиска
        search_button = tk.Button(search_frame, text="Поиск", command=self.perform_search)
        search_button.pack()

        # Кнопка "Копировать"
        copy_button = tk.Button(search_frame, text="Копировать", command=self.copy_info)
        copy_button.pack()
        self.resizable(False, False)

    def update_info_labels(self):
        if self.selected_process:
            for i, label in enumerate(self.info_labels):
                field_name = list(self.process_info[0].keys())[i]
                label.config(text=f"{self.labels_info[i]}{self.selected_process.get(field_name, '')}")

    def perform_search(self):
        search_text = self.search_entry.get().lower()
        self.process_listbox.delete(0, tk.END)  # Clear the listbox

        for process_info in self.process_info:
            for field_name, field_value in process_info.items():

                if search_text in field_value.lower():
                    print(search_text)
                    self.process_listbox.insert(tk.END, process_info["Наименование воинской части"])
                    break  # Move to the next process after adding the current one

    def copy_info(self):
        if self.selected_process:
            # Собираем информацию в одну строку
            info_text = "\n".join([f"{label.cget('text')}{self.selected_process.get(field_name, '')}" for label, field_name in zip(self.info_labels, self.process_info[0].keys())])
            # Копируем информацию в буфер обмена
            pyperclip.copy(info_text)

if __name__ == "__main__":
    app = ProcessTrackerApp()
    app.mainloop()
