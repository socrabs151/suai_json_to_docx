# gui/main_window.py

"""Главное окно приложения для генерации DOCX файлов из JSON и шаблонов."""

import tkinter as tk
from tkinter import ttk, simpledialog, filedialog, messagebox
from typing import Dict, Callable, Any

from modules.json_reader import (
    load_json,
    papers_json_to_dataframe,
    report_json_to_dataframe,
)
from modules import (
    publish_docx_generator,
    program_docx_generator,
    report_docx_generator,
)
from modules.template_manager import choose_template


class MainWindow(tk.Tk):
    """Главное окно приложения для генерации документов."""

    def __init__(self) -> None:
        """Инициализирует главное окно приложения."""
        super().__init__()
        self.title("Генератор DOCX по JSON и шаблону")
        self.geometry("800x600")
        self.resizable(True, True)

        self.dataframes: Dict[str, Any] = {}
        self.placeholders: Dict[str, list] = {}
        self.placeholder_values: Dict[str, Dict[str, str]] = {}

        self.json_to_df_functions: Dict[str, Callable] = {
            "Список представляемых к публикации докладов": papers_json_to_dataframe,
            "Программа": report_json_to_dataframe,
            "Отчет о проведении": report_json_to_dataframe,
        }

        self.create_widgets()

    def create_widgets(self) -> None:
        """Создает все виджеты интерфейса."""
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tabs: Dict[str, ttk.Frame] = {}

        for label in [
            "Программа",
            "Отчет о проведении",
            "Список представляемых к публикации докладов",
        ]:
            self.create_tab(label)

        self.create_actions()
        self.create_status_bar()

    def create_tab(self, name: str) -> None:
        """
        Создает вкладку с элементами управления.

        Args:
            name: Название вкладки.
        """
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=name)
        self.tabs[name] = tab

        # --- Файлы
        frame_files = ttk.Frame(tab)
        frame_files.pack(fill="x", pady=5)

        setattr(self, f"{name}_json_path", tk.StringVar())
        setattr(self, f"{name}_template_path", tk.StringVar())

        ttk.Label(frame_files, text="JSON-файл:").grid(row=0, column=0, padx=5, sticky="w")
        ttk.Entry(frame_files, textvariable=getattr(self, f"{name}_json_path"), width=50).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(
            frame_files,
            text="\U0001F4C1 Загрузить",
            command=lambda n=name: load_json(self, n, self.json_to_df_functions),
        ).grid(row=0, column=2)

        ttk.Label(frame_files, text="Шаблон DOCX:").grid(row=1, column=0, padx=5, sticky="w")
        ttk.Entry(frame_files, textvariable=getattr(self, f"{name}_template_path"), width=50).grid(
            row=1, column=1, padx=5
        )
        ttk.Button(
            frame_files,
            text="\U0001F4C1 Загрузить",
            command=lambda n=name: choose_template(self, n),
        ).grid(row=1, column=2)

        # --- Плейсхолдеры
        frame_placeholders = ttk.LabelFrame(tab, text="Плейсхолдеры")
        frame_placeholders.pack(fill="x", pady=10)

        placeholder_tree = ttk.Treeview(frame_placeholders, height=5)
        placeholder_tree.pack(fill="both", expand=True, padx=5, pady=5)
        placeholder_tree["columns"] = ("Placeholder", "Value")
        placeholder_tree["show"] = "headings"
        placeholder_tree.heading("Placeholder", text="Плейсхолдер")
        placeholder_tree.heading("Value", text="Значение")
        placeholder_tree.column("Placeholder", width=200)
        placeholder_tree.column("Value", width=200)
        setattr(self, f"{name}_placeholder_tree", placeholder_tree)

        # --- Таблица предпросмотра
        frame_table = ttk.Frame(tab)
        frame_table.pack(fill="both", expand=True, pady=5)

        tree = ttk.Treeview(frame_table, height=5)
        tree.pack(fill="both", expand=True, padx=5, pady=5)
        setattr(self, f"{name}_tree", tree)

        frame_options = ttk.LabelFrame(tab, text="Параметры")
        frame_options.pack(fill="x", padx=5, pady=5)

        setattr(self, f"{name}_event_name", tk.StringVar())
        setattr(self, f"{name}_output_path", tk.StringVar())

    def show_dataframe_in_tree(self, name: str, df: Any) -> None:
        """
        Отображает DataFrame в Treeview.

        Args:
            name: Название вкладки.
            df: DataFrame для отображения.
        """
        tree = getattr(self, f"{name}_tree")
        tree.delete(*tree.get_children())

        tree["columns"] = list(df.columns)
        tree["show"] = "headings"

        for col in df.columns:
            tree.heading(col, text=col)
            max_width = max([len(str(val)) for val in df[col]] + [len(str(col))])
            tree.column(col, width=max(80, min(max_width * 7 + 20, 500)))

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def show_placeholders_in_tree(self, name: str) -> None:
        """
        Отображает плейсхолдеры в Treeview.

        Args:
            name: Название вкладки.
        """
        tree = getattr(self, f"{name}_placeholder_tree")
        tree.delete(*tree.get_children())

        for ph in sorted(self.placeholders.get(name, [])):
            value = self.placeholder_values[name].get(ph, "")
            tree.insert("", "end", values=(ph, value))

        def on_double_click(event: tk.Event) -> None:
            """Обрабатывает двойной клик по плейсхолдеру."""
            item = tree.selection()
            if not item:
                return
            placeholder = tree.item(item)["values"][0]
            current_value = self.placeholder_values[name].get(placeholder, "")
            new_value = simpledialog.askstring(
                "Ввод значения",
                f"Введите значение для {placeholder}:",
                initialvalue=current_value,
            )
            if new_value is not None:
                self.placeholder_values[name][placeholder] = new_value
                self.show_placeholders_in_tree(name)

        tree.bind("<Double-1>", on_double_click)

    def select_output_folder(self, output_var: tk.StringVar) -> None:
        """
        Открывает диалог выбора папки.

        Args:
            output_var: Переменная для хранения пути.
        """
        path = filedialog.askdirectory()
        if path:
            output_var.set(path)

    def create_actions(self) -> None:
        """Создает кнопки действий."""
        frame = ttk.Frame(self)
        frame.pack(pady=10)

        ttk.Button(
            frame,
            text="\U0001F4BE Сформировать DOCX",
            command=self.generate_docx,
        ).grid(row=0, column=0, padx=10)

    def create_status_bar(self) -> None:
        """Создает строку состояния."""
        self.status = tk.StringVar(value="Готово.")
        ttk.Label(self, textvariable=self.status, foreground="gray").pack(pady=5)

    def generate_docx(self) -> None:
        """Генерирует DOCX файл в зависимости от активной вкладки."""
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        if current_tab == "Список представляемых к публикации докладов":
            publish_docx_generator.generate_docx(self, current_tab)
        elif current_tab == "Программа":
            program_docx_generator.generate_docx(self, current_tab)
        elif current_tab == "Отчет о проведении":
            report_docx_generator.generate_docx(self, current_tab)
        else:
            messagebox.showinfo(
                "Информация",
                "DOCX можно сгенерировать только во вкладке "
                "'Список представляемых к публикации докладов'.",
            )