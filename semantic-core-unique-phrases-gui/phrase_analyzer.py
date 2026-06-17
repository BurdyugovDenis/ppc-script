import pymorphy3
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class PhraseProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор ключевых фраз")
        self.root.geometry("600x400")
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Современный стиль для Windows 11

        # Переменные
        self.existing_file = tk.StringVar()
        self.new_file = tk.StringVar()
        self.output_file = tk.StringVar(value="output.txt")

        self.create_widgets()

    def create_widgets(self):
        # Фрейм для файлов
        file_frame = ttk.LabelFrame(self.root, text="Файлы", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)

        # Поля выбора файлов
        ttk.Label(file_frame, text="Существующие фразы:").grid(row=0, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.existing_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Обзор...", command=lambda: self.browse_file(self.existing_file)).grid(row=0,
                                                                                                           column=2)

        ttk.Label(file_frame, text="Новые фразы:").grid(row=1, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.new_file, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Обзор...", command=lambda: self.browse_file(self.new_file)).grid(row=1, column=2)

        ttk.Label(file_frame, text="Результат:").grid(row=2, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_file, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(file_frame, text="Обзор...", command=lambda: self.save_file(self.output_file)).grid(row=2, column=2)

        # Кнопка обработки
        process_btn = ttk.Button(self.root, text="Найти уникальные фразы", command=self.process_files)
        process_btn.pack(pady=10)

        # Прогрессбар
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=5)

        # Лог
        self.log = tk.Text(self.root, height=10, state="disabled")
        self.log.pack(fill="both", expand=True, padx=10, pady=5)

    def browse_file(self, file_var):
        filename = filedialog.askopenfilename(filetypes=[("Текстовые файлы", "*.txt")])
        if filename:
            file_var.set(filename)

    def save_file(self, file_var):
        filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Текстовые файлы", "*.txt")])
        if filename:
            file_var.set(filename)

    def log_message(self, message):
        self.log.config(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.config(state="disabled")
        self.root.update()

    def process_files(self):
        try:
            # Проверка файлов
            if not Path(self.existing_file.get()).exists():
                raise FileNotFoundError("Файл с существующими фразами не выбран!")
            if not Path(self.new_file.get()).exists():
                raise FileNotFoundError("Файл с новыми фразами не выбран!")

            self.progress["value"] = 0
            self.log_message("Начата обработка файлов...")

            # Чтение файлов
            existing_phrases = self.read_phrases(self.existing_file.get())
            new_phrases = self.read_phrases(self.new_file.get())
            self.progress["value"] = 20
            self.log_message(f"Прочитано {len(existing_phrases)} существующих и {len(new_phrases)} новых фраз")

            # Обработка
            morph = pymorphy3.MorphAnalyzer()
            stop_words = self.get_stop_words()

            self.log_message("Обработка существующих фраз...")
            existing_canonical = set()
            for i, phrase in enumerate(existing_phrases):
                canonical = self.process_phrase(phrase, stop_words, morph)
                existing_canonical.add(canonical)
                if i % 100 == 0:
                    self.progress["value"] = 20 + (i / len(existing_phrases)) * 30

            self.log_message("Фильтрация новых фраз...")
            unique_phrases = []
            for i, phrase in enumerate(new_phrases):
                canonical = self.process_phrase(phrase, stop_words, morph)
                if canonical not in existing_canonical:
                    unique_phrases.append(phrase)
                if i % 100 == 0:
                    self.progress["value"] = 50 + (i / len(new_phrases)) * 40

            # Сохранение
            Path(self.output_file.get()).write_text('\n'.join(unique_phrases), encoding='utf-8')
            self.progress["value"] = 100
            self.log_message(f"Готово! Найдено {len(unique_phrases)} уникальных фраз")
            messagebox.showinfo("Успех", f"Обработка завершена!\nСохранено {len(unique_phrases)} уникальных фраз")

        except Exception as e:
            self.log_message(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", str(e))

    def read_phrases(self, filename):
        with open(filename, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]

    def process_phrase(self, phrase, stop_words, morph):
        phrase_clean = re.sub(r'[^\w\s]', '', phrase.lower())
        words = phrase_clean.split()
        processed_words = []

        for word in words:
            if word in stop_words:
                continue
            lemma = morph.parse(word)[0].normal_form
            processed_words.append(lemma)

        processed_words.sort()
        return ' '.join(processed_words)

    def get_stop_words(self):
        return {
            "в", "на", "с", "по", "и", "а", "но", "к", "у", "о", "от", "до", "за",
            "из", "или", "то", "же", "бы", "вот", "как", "для", "но", "да", "я",
            "ты", "он", "она", "они", "мы", "вы", "это", "тот", "этот", "такой",
            "также", "тоже", "при", "над", "под", "перед", "после", "через", "без",
            "со", "ли", "ни", "не", "будто", "что", "чтобы", "хотя", "если", "потому",
            "какой", "который", "где", "когда", "чем", "как", "чему", "кого", "чего",
            "зачем", "почему", "откуда", "куда", "сколько", "чей", "надо", "можно",
            "нужно", "нет", "да", "ну", "ага", "ок", "ладно"
        }


if __name__ == '__main__':
    root = tk.Tk()
    app = PhraseProcessorApp(root)
    root.mainloop()