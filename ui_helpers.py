"""Мелкие помощники интерфейса tkinter/ttk.

Содержит:

* :func:`attach_placeholder` — подсказка-образец в пустом :class:`ttk.Entry`;
* :func:`get_value` — чтение реального значения (placeholder не попадает);
* :func:`refresh_placeholder` — повторно показать подсказку, если поле пустое;
* :func:`normalize_datetime_input` — приведение «сырых» цифр к
  каноническому виду ``YYYY-MM-DD HH:MM`` (или ``DD.MM.YYYY``);
* :func:`on_datetime_key` — пост-обработка ``<KeyRelease>`` для полей
  даты-времени (ввод не блокируется, только переформатируется готовый ввод).

Новых зависимостей нет — только стандартная библиотека и tkinter.
"""

from __future__ import annotations

import re
import tkinter as tk
from datetime import date
from tkinter import ttk
from typing import Optional

#: Имя стиля ttk для «серого» текста подсказки.
PLACEHOLDER_STYLE = "Placeholder.TEntry"
#: Цвет текста подсказки.
PLACEHOLDER_COLOR = "#888888"

#: Количество цифр для полного «сырого» ввода даты-времени.
DATETIME_DIGITS = 12
#: Количество цифр для ввода только даты (DDMMYYYY).
DATE_DIGITS = 8

_DIGITS_AND_SEPARATORS = re.compile(r"^[\d\s\-\.:]+$")
_NON_DIGITS = re.compile(r"\D")


# ---------------------------------------------------------------------- #
# Нормализация ввода даты-времени
# ---------------------------------------------------------------------- #
def normalize_datetime_input(text: str) -> str:
    """Привести «сырые» цифры даты-времени к каноническому виду.

    * 12 цифр (``YYYYMMDDHHMM``) -> ``YYYY-MM-DD HH:MM``;
    * 8 цифр (``DDMMYYYY``) -> ``DD.MM.YYYY`` (время пользователь дополнит);
    * иначе текст возвращается без изменений — пусть ``parse_dt_text``
      попробует разобрать его как есть.
    """
    digits = _NON_DIGITS.sub("", text or "")
    if len(digits) == DATETIME_DIGITS:
        return (
            f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]} "
            f"{digits[8:10]}:{digits[10:12]}"
        )
    if len(digits) == DATE_DIGITS:
        return f"{digits[0:2]}.{digits[2:4]}.{digits[4:8]}"
    return text


def _is_ddmmyyyy(digits: str) -> bool:
    """Проверить, что 8 цифр — реальная дата в формате ``DDMMYYYY``."""
    try:
        date(int(digits[4:8]), int(digits[2:4]), int(digits[0:2]))
    except ValueError:
        return False
    return True


def on_datetime_key(event: "tk.Event | None", entry: ttk.Entry, var: tk.StringVar) -> None:
    """Пост-обработка ``<KeyRelease>`` для поля даты-времени.

    Ввод не блокируется: если в поле только цифры и разделители и «сырых»
    цифр достаточно, содержимое переформатируется (курсор — в конец).
    Ввод с разделителями вручную (``2026-09-17 13:28``) не меняется.
    """
    if getattr(entry, "_placeholder_active", False):
        return
    text = var.get()
    if not text:
        return
    if not _DIGITS_AND_SEPARATORS.match(text):
        return

    digits = _NON_DIGITS.sub("", text)
    enough_digits = len(digits) >= DATETIME_DIGITS
    if not enough_digits:
        # 8 цифр — только если это уже похоже на готовую дату DDMMYYYY.
        # Так ввод ``20260917`` (начало ``YYYYMMDDHHMM``) не мешает
        # пользователю продолжить набор до полных 12 цифр.
        enough_digits = len(digits) == DATE_DIGITS and _is_ddmmyyyy(digits)

    if not enough_digits:
        return

    normalized = normalize_datetime_input(text)
    if normalized != text:
        var.set(normalized)
        entry.icursor("end")


# ---------------------------------------------------------------------- #
# Placeholder для ttk.Entry
# ---------------------------------------------------------------------- #
def _entry_var(entry: ttk.Entry) -> Optional[tk.StringVar]:
    return getattr(entry, "_placeholder_var", None)


def _entry_text(entry: ttk.Entry) -> str:
    var = _entry_var(entry)
    return var.get() if var is not None else entry.get()


def _set_entry_text(entry: ttk.Entry, text: str) -> None:
    var = _entry_var(entry)
    if var is not None:
        var.set(text)
        return
    entry.delete(0, "end")
    if text:
        entry.insert(0, text)


def _set_gray(entry: ttk.Entry, gray: bool) -> None:
    normal_style = getattr(entry, "_placeholder_normal_style", "")
    entry.configure(style=PLACEHOLDER_STYLE if gray else normal_style)


def _hide_placeholder(entry: ttk.Entry) -> None:
    """Убрать подсказку при получении фокуса."""
    if not getattr(entry, "_placeholder_active", False):
        return
    _set_entry_text(entry, "")
    _set_gray(entry, False)
    entry._placeholder_active = False


def refresh_placeholder(entry: ttk.Entry) -> None:
    """Показать подсказку, если поле пустое (и убрать, если уже есть текст)."""
    placeholder = getattr(entry, "_placeholder_text", "")
    if not placeholder:
        return
    if getattr(entry, "_placeholder_active", False):
        return
    if _entry_text(entry).strip():
        _set_gray(entry, False)
        return
    _set_entry_text(entry, placeholder)
    _set_gray(entry, True)
    entry._placeholder_active = True


def get_value(entry: ttk.Entry) -> str:
    """Реальное значение поля: ``""``, если сейчас показана подсказка."""
    if getattr(entry, "_placeholder_active", False):
        return ""
    return _entry_text(entry)


def attach_placeholder(
    entry: ttk.Entry, placeholder: str, var: Optional[tk.StringVar] = None
) -> ttk.Entry:
    """Повесить подсказку-образец на :class:`ttk.Entry`.

    ``var`` — связанная строка (``textvariable``), чтобы :func:`get_value`
    и :func:`refresh_placeholder` могли работать с реальным содержимым.
    """
    entry._placeholder_text = placeholder
    entry._placeholder_active = False
    entry._placeholder_var = var
    entry._placeholder_normal_style = entry.cget("style")

    try:
        ttk.Style(entry).configure(PLACEHOLDER_STYLE, foreground=PLACEHOLDER_COLOR)
    except tk.TclError:
        # Тема не поддерживает foreground для TEntry — подсказка просто
        # останется обычным цветом, функциональность не страдает.
        pass

    entry.bind("<FocusIn>", lambda _event: _hide_placeholder(entry), add="+")
    entry.bind("<FocusOut>", lambda _event: refresh_placeholder(entry), add="+")
    refresh_placeholder(entry)
    return entry
