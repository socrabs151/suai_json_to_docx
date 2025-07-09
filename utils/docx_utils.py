# utils/docx_utils.py

"""Утилиты для работы с DOCX документами."""

import re
from typing import Dict, Any
from docx.text.paragraph import Paragraph

# Месяцы на русском языке для форматирования дат
months_ru = {
    1: "января",
    2: "февраля",
    3: "марта",
    4: "апреля",
    5: "мая",
    6: "июня",
    7: "июля",
    8: "августа",
    9: "сентября",
    10: "октября",
    11: "ноября",
    12: "декабря",
}


def replace_placeholders_in_para(
        para: Paragraph,
        placeholders: Dict[str, str]
) -> None:
    """
    Заменяет плейсхолдеры в тексте параграфа на соответствующие значения.

    Плейсхолдеры должны быть в формате {ключ} или {ключ::значение_по_умолчанию}.
    Сохраняет исходное форматирование текста.

    Args:
        para: Объект параграфа из docx
        placeholders: Словарь замен {ключ: значение}

    Example:
        >>> para.text = "Привет, {имя::гость}!"
        >>> replace_placeholders_in_para(para, {"имя": "Алексей"})
        >>> para.text
        'Привет, Алексей!'
    """
    placeholder_pattern = re.compile(r"\{([^{}]+?)\}")
    new_text_parts = []
    full_text = para.text
    matches = list(placeholder_pattern.finditer(full_text))
    last_index = 0

    for match in matches:
        start, end = match.span()
        raw_key = match.group(1)

        # Разделяем имя и значение по умолчанию
        if "::" in raw_key:
            key, default = raw_key.split("::", 1)
        else:
            key, default = raw_key, ""

        key = key.strip()
        default = default.strip()

        # Получаем значение для подстановки
        value = placeholders.get(key, default)

        # Добавляем текст до плейсхолдера
        if start > last_index:
            new_text_parts.append(full_text[last_index:start])

        # Добавляем заменённое значение
        new_text_parts.append(value)
        last_index = end

    # Добавляем оставшийся текст после последнего плейсхолдера
    if last_index < len(full_text):
        new_text_parts.append(full_text[last_index:])

    # Собираем итоговый текст
    new_text = "".join(new_text_parts)

    # Очищаем существующие runs и добавляем новый текст
    # с сохранением форматирования первого run'а
    if para.runs:
        # Сохраняем форматирование первого run'а
        first_run = para.runs[0]
        for run in para.runs[1:]:
            run.text = ""
        first_run.text = new_text
    else:
        # Если runs нет вообще - создаем новый
        para.add_run(new_text)