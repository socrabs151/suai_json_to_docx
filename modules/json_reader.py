# modules/json_reader.py

"""Модуль для чтения и обработки JSON файлов."""

from tkinter import filedialog, messagebox
from typing import Any, Dict, List, Callable
import json
import pandas as pd


def load_json(
    self: Any,
    name: str,
    json_to_df_functions: Dict[str, Callable[[Dict[str, Any]], pd.DataFrame]]
) -> None:
    """
    Загружает JSON файл и преобразует его в DataFrame.

    Args:
        self: Экземпляр главного окна.
        name: Название вкладки, для которой загружается файл.
        json_to_df_functions: Словарь функций для преобразования JSON в DataFrame.
    """
    path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if not path:
        return
    getattr(self, f"{name}_json_path").set(path)

    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)

        json_to_df_func = json_to_df_functions.get(name)
        if not json_to_df_func:
            messagebox.showerror(
                "Ошибка", f"Нет функции обработки JSON для вкладки {name}"
            )
            return

        df = json_to_df_func(data)
        self.dataframes[name] = df
        self.show_dataframe_in_tree(name, df)
        self.status.set(f"Загружен файл: {path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")


def papers_json_to_dataframe(data: Dict[str, Any]) -> pd.DataFrame:
    """
    Преобразует JSON с данными о докладах в DataFrame.

    Args:
        data: Словарь с данными из JSON файла.

    Returns:
        DataFrame с отфильтрованными и обработанными данными о докладах.
    """
    papers = data.get("papers", [])
    rows = []

    for paper in papers:
        state = paper.get("state", {}).get("name", "").lower()
        if state != "accepted":
            continue

        contribution = paper.get("contribution", {})
        last_revision = next(
            (
                rev
                for rev in paper.get("revisions", [])
                if rev.get("is_last_revision", False)
            ),
            paper.get("revisions", [{}])[0],
        )
        submitter = last_revision.get("submitter", {})
        full_name = submitter.get("full_name", "")

        formatted_name = convert_full_name(full_name)

        row = {
            "Title": contribution.get("title", ""),
            "State": paper.get("state", {}).get("title", ""),
            "Submitter": formatted_name,
        }
        rows.append(row)

    return pd.DataFrame(rows)


def convert_full_name(full_name: str) -> str:
    """
    Форматирует полное имя в сокращенный вид (Фамилия И.О.).

    Args:
        full_name: Полное имя в формате "Имя Отчество Фамилия" или "Имя Фамилия".

    Returns:
        Отформатированное имя в виде "Фамилия И.О.".
    """
    parts = full_name.strip().split()
    if len(parts) == 3:
        first_name, patronymic, last_name = parts
        return f"{last_name} {first_name[0]}.{patronymic[0]}."
    elif len(parts) == 2:
        first_name, last_name = parts
        return f"{last_name} {first_name[0]}."
    else:
        return full_name


def report_json_to_dataframe(data: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Преобразует JSON с данными о докладах в DataFrame для программы/отчета.

    Args:
        data: Список словарей с данными о докладах.

    Returns:
        DataFrame с обработанными данными о докладах.
    """
    rows = []
    for abstract in data:
        group_number = ""
        for field in abstract.get("custom_fields", []):
            if field["name"] == "Номер группы основного автора (докладчика)":
                group_number = field.get("value", "")
                break

        speaker_name = ""
        if abstract.get("persons"):
            speaker_name = abstract["persons"][0].get("full_name", "")

        title = abstract.get("title", "")
        start_dt = abstract.get("start_dt", "")
        room_name = abstract.get("room_name", "")

        rows.append(
            {
                "Номер группы": group_number,
                "ФИО докладчика": speaker_name,
                "Название доклада": title,
                "Дата и время начала": start_dt,
                "Ауд.": room_name,
            }
        )

    return pd.DataFrame(rows)