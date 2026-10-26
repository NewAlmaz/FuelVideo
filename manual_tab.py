"""Вкладка «Ручное копирование» — копирование по параметрам, заданным вручную.

Пользователь задаёт интервал (начало/конец), список ТРК и АЗС, добавляет строки
в таблицу, затем нажимает «Найти все» (только проверка) или «Копировать все».
Таблицу строк можно импортировать из xlsx и экспортировать обратно
(см. :mod:`manual_xlsx`).

Логика поиска и копирования не дублируется: используются общие функции
:func:`copier.find_files_for_cameras` и :func:`copier.copy_found_files`
(те же, что и на вкладке «Задания»).

``minutes_before``/``minutes_after`` здесь НЕ применяются: интервал берётся
ровно таким, каким его задал пользователь. Сдвиг часов видео-сервера
(``mtime_fix_sec`` + ``clock_drift_minutes``) применяется всегда.
"""

from __future__ import annotations

import csv
import logging
import os
import queue
import threading
import tkinter as tk
from dataclasses import replace
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Sequence, Tuple

import drives
from config import AppConfig, AzsConfig
from copier import (
    STATUS_ALREADY_EXISTS,
    STATUS_COPIED,
    STATUS_ERROR,
    STATUS_SIZE_MISMATCH,
    CopyRecord,
    copy_found_files,
    destination_for,
    find_files_for_cameras,
    status_label,
)
from manual_xlsx import (
    DATETIME_FORMAT,
    ManualRow,
    detect_xlsx_format,
    parse_dt_text,
    read_manual_xlsx,
    write_manual_xlsx,
)
from video_finder import FoundFile
from xlsx_reader import XlsxError

LOGGER = logging.getLogger(__name__)

DATETIME_HINT = (
    "Неверный формат. Ожидается YYYY-MM-DD HH:MM, например 2026-09-17 13:28"
)

#: Служебный статус «найдено, но не скопировано» (только для этой вкладки).
STATUS_WILL_COPY = "will_copy"
WILL_COPY_LABEL = "будет скопировано"

#: Колонки CSV-отчёта ручного копирования.
MANUAL_REPORT_COLUMNS = ["row_index", "kind", "camera", "src", "dst", "status", "error"]

ROWS_COLUMNS = ("num", "azs", "start", "end", "trk", "status")
ROWS_HEADINGS = {
    "num": "№",
    "azs": "АЗС",
    "start": "Начало",
    "end": "Конец",
    "trk": "ТРК",
    "status": "Статус",
}
ROWS_WIDTHS = {
    "num": 40,
    "azs": 55,
    "start": 140,
    "end": 140,
    "trk": 110,
    "status": 220,
}

RESULT_COLUMNS = ("row", "kind", "camera", "mtime", "src", "dst", "status")
RESULT_HEADINGS = {
    "row": "Строка",
    "kind": "Тип",
    "camera": "Камера",
    "mtime": "Время файла",
    "src": "Источник",
    "dst": "Назначение",
    "status": "Статус",
}
RESULT_WIDTHS = {
    "row": 55,
    "kind": 55,
    "camera": 60,
    "mtime": 135,
    "src": 330,
    "dst": 330,
    "status": 200,
}


def format_minutes(value: float) -> str:
    """Число минут без лишних нулей: ``5.0`` -> «5», ``7.50`` -> «7.5»."""
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def clock_drift_text(value: float) -> str:
    """Read-only описание сдвига часов: ``5 мин (часы отстают)``.

    Положительный ``clock_drift_minutes`` означает, что часы видео-сервера
    отстают (к ``mtime`` прибавляется сдвиг), отрицательный — спешат.

    :param value: значение ``clock_drift_minutes`` из settings.ini
    :return: текст для строки «Сдвиг часов: …»
    """
    drift = float(value or 0.0)
    if abs(drift) < 1e-9:
        return "0 мин (без сдвига)"
    if drift > 0:
        return f"{format_minutes(drift)} мин (часы отстают)"
    return f"{format_minutes(-drift)} мин (часы спешат)"


class ManualTab(ttk.Frame):
    """Вкладка «Ручное копирование»: редактор строк и их выполнение."""

    def __init__(self, master: tk.Misc, app: object) -> None:
        super().__init__(master, padding=6)
        self.app = app
        self.manual_rows: List[ManualRow] = []
        self._worker: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._event_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self._busy = False

        self.start_var = tk.StringVar()
        self.end_var = tk.StringVar()
        self.trk_var = tk.StringVar()
        # Read-only строка «Сдвиг часов: …»; сам редактор живёт на вкладке «Настройки».
        self.clock_value_var = tk.StringVar(value="Сдвиг часов: —")
        self.status_var = tk.StringVar(value="Готов к работе")

        self._build_ui()
        self.refresh_clock_label()
        self.refresh_expiry_state()
        self.after(150, self._poll_events)

    # ------------------------------------------------------------------ #
    # Доступ к текущим настройкам приложения
    # ------------------------------------------------------------------ #
    @property
    def config(self) -> AppConfig:
        """Текущие настройки приложения."""
        return self.app.config_data  # type: ignore[attr-defined]

    @property
    def azs(self) -> AzsConfig:
        """Текущая АЗС."""
        return self.app.azs  # type: ignore[attr-defined]

    # ------------------------------------------------------------------ #
    # Построение интерфейса
    # ------------------------------------------------------------------ #
    def _build_ui(self) -> None:
        """Создать форму, таблицы, панели запуска и строку сдвига часов."""
        self._build_form()
        self._build_rows_tree()
        self._build_run_panel()
        self._build_progress()
        self._build_results_tree()

    def _build_form(self) -> None:
        """Форма добавления строки."""
        form = ttk.LabelFrame(self, text="Добавить строку", padding=6)
        form.pack(fill="x")

        ttk.Label(form, text="Начало").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.start_var, width=18).grid(
            row=0, column=1, sticky="w", padx=(4, 12)
        )
        ttk.Label(form, text="Конец").grid(row=0, column=2, sticky="w")
        ttk.Entry(form, textvariable=self.end_var, width=18).grid(
            row=0, column=3, sticky="w", padx=(4, 12)
        )
        ttk.Label(form, text="ТРК").grid(row=0, column=4, sticky="w")
        ttk.Entry(form, textvariable=self.trk_var, width=14).grid(
            row=0, column=5, sticky="w", padx=(4, 12)
        )
        self.add_button = ttk.Button(form, text="Добавить строку", command=self.on_add_row)
        self.add_button.grid(row=0, column=6, sticky="w", padx=(12, 0))
        form.columnconfigure(6, weight=1)

        ttk.Label(
            form,
            text=(
                f"Формат: {DATETIME_FORMAT}, например 2026-09-17 13:28. "
                "ТРК — через запятую (1,2); пусто = все ТРК из azs.ini."
            ),
        ).grid(row=1, column=0, columnspan=7, sticky="w", pady=(6, 0))

    def _build_rows_tree(self) -> None:
        """Таблица добавленных строк и кнопки её редактирования."""
        frame = ttk.LabelFrame(self, text="Строки", padding=6)
        frame.pack(fill="x", pady=(6, 0))

        tree_holder = ttk.Frame(frame)
        tree_holder.pack(fill="x")
        self.rows_tree = ttk.Treeview(
            tree_holder, columns=ROWS_COLUMNS, show="headings", height=5
        )
        for column in ROWS_COLUMNS:
            self.rows_tree.heading(column, text=ROWS_HEADINGS[column])
            anchor = "center" if column in ("num", "azs", "trk") else "w"
            self.rows_tree.column(column, width=ROWS_WIDTHS[column], anchor=anchor)
        self.rows_tree.tag_configure("error", foreground="#b00020")
        self.rows_tree.tag_configure("ok", foreground="#0a7d24")
        self.rows_tree.tag_configure("warn", foreground="#8a6d00")

        rows_vsb = ttk.Scrollbar(tree_holder, orient="vertical", command=self.rows_tree.yview)
        self.rows_tree.configure(yscrollcommand=rows_vsb.set)
        self.rows_tree.pack(side="left", fill="both", expand=True)
        rows_vsb.pack(side="right", fill="y")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(6, 0))
        self.delete_button = ttk.Button(
            buttons, text="Удалить выбранную", command=self.on_delete_selected
        )
        self.delete_button.pack(side="left")
        self.clear_button = ttk.Button(buttons, text="Очистить всё", command=self.on_clear_all)
        self.clear_button.pack(side="left", padx=6)
        self.export_button = ttk.Button(
            buttons, text="Экспорт в xlsx", command=self.on_export_xlsx
        )
        self.export_button.pack(side="left", padx=(12, 6))
        self.import_button = ttk.Button(
            buttons, text="Импорт из xlsx", command=self.on_import_xlsx
        )
        self.import_button.pack(side="left")

    def _build_run_panel(self) -> None:
        """Кнопки выполнения."""
        frame = ttk.LabelFrame(self, text="Выполнение", padding=6)
        frame.pack(fill="x", pady=(6, 0))

        self.find_button = ttk.Button(frame, text="Найти все", command=self.on_find_all)
        self.find_button.pack(side="left")
        self.copy_button = ttk.Button(frame, text="Копировать все", command=self.on_copy_all)
        self.copy_button.pack(side="left", padx=6)
        self.stop_button = ttk.Button(
            frame, text="Остановить", command=self.on_stop, state="disabled"
        )
        self.stop_button.pack(side="left", padx=(0, 6))
        self.open_button = ttk.Button(
            frame, text="Открыть результат", command=self.on_open_output
        )
        self.open_button.pack(side="left")

        # Сдвиг часов задаётся на вкладке «Настройки»; здесь только показ и переход.
        ttk.Button(frame, text="Изменить", command=self.on_edit_clock).pack(side="right")
        ttk.Label(frame, textvariable=self.clock_value_var).pack(side="right", padx=(12, 6))

    def _build_progress(self) -> None:
        """Прогресс-бар и строка статуса."""
        frame = ttk.Frame(self)
        frame.pack(fill="x", pady=(6, 0))
        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=100)
        self.progress.pack(side="left", fill="x", expand=True)
        ttk.Label(frame, textvariable=self.status_var, padding=(8, 0)).pack(side="left")

    def _build_results_tree(self) -> None:
        """Таблица результатов поиска/копирования."""
        frame = ttk.LabelFrame(self, text="Результаты", padding=6)
        frame.pack(fill="both", expand=True, pady=(6, 0))

        self.results_tree = ttk.Treeview(
            frame, columns=RESULT_COLUMNS, show="headings", height=8
        )
        for column in RESULT_COLUMNS:
            self.results_tree.heading(column, text=RESULT_HEADINGS[column])
            anchor = "center" if column in ("row", "kind", "camera") else "w"
            self.results_tree.column(column, width=RESULT_WIDTHS[column], anchor=anchor)
        self.results_tree.tag_configure("error", foreground="#b00020")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.results_tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    # ------------------------------------------------------------------ #
    # Сдвиг часов видео-сервера (редактор — на вкладке «Настройки»)
    # ------------------------------------------------------------------ #
    def refresh_clock_label(self) -> None:
        """Обновить read-only строку «Сдвиг часов: …» из settings.ini."""
        value = float(getattr(self.config.finder, "clock_drift_minutes", 0.0) or 0.0)
        try:
            self.clock_value_var.set(f"Сдвиг часов: {clock_drift_text(value)}")
        except tk.TclError:  # окно уже закрывается
            pass

    def on_edit_clock(self) -> None:
        """Перейти к редактированию сдвига часов на вкладке «Настройки»."""
        handler = getattr(self.app, "on_edit_clock", None)
        if callable(handler):
            handler()  # type: ignore[misc]

    # ------------------------------------------------------------------ #
    # Добавление и удаление строк
    # ------------------------------------------------------------------ #
    @staticmethod
    def _parse_datetime(text: str) -> Optional[datetime]:
        """Разобрать дату-время формата ``YYYY-MM-DD HH:MM``.

        Используется общий helper из :mod:`manual_xlsx`, чтобы форма и чтение
        xlsx трактовали даты одинаково.
        """
        return parse_dt_text(text)

    def _parse_trks(self, text: str) -> Optional[List[int]]:
        """Разобрать список ТРК; пустое поле — все ТРК текущей АЗС."""
        text = text.strip()
        if not text:
            return sorted(self.azs.trk_map.keys())

        trks: List[int] = []
        for chunk in text.replace(";", ",").split(","):
            chunk = chunk.strip()
            if not chunk:
                continue
            try:
                trk = int(chunk)
            except ValueError:
                messagebox.showerror(
                    "Ручное копирование",
                    "ТРК должны быть числами через запятую, например 1,2.",
                    parent=self,
                )
                return None
            if trk not in trks:
                trks.append(trk)

        if not trks:
            messagebox.showerror(
                "Ручное копирование", "Не удалось разобрать список ТРК.", parent=self
            )
            return None
        return trks

    def on_add_row(self) -> None:
        """Проверить поля формы и добавить строку в таблицу."""
        start_text = self.start_var.get().strip()
        end_text = self.end_var.get().strip()
        if not start_text or not end_text:
            messagebox.showerror(
                "Ручное копирование",
                "Заполните поля «Начало» и «Конец».",
                parent=self,
            )
            return

        start = self._parse_datetime(start_text)
        if start is None:
            messagebox.showerror("Неверный формат", DATETIME_HINT, parent=self)
            return
        end = self._parse_datetime(end_text)
        if end is None:
            messagebox.showerror("Неверный формат", DATETIME_HINT, parent=self)
            return

        if end <= start:
            messagebox.showerror(
                "Ручное копирование",
                "Время окончания должно быть больше времени начала.",
                parent=self,
            )
            return

        if start > datetime.now() + timedelta(hours=1):
            if not messagebox.askyesno(
                "Интервал в будущем",
                f"Начало интервала ({start.strftime(DATETIME_FORMAT)}) позже текущего "
                "времени более чем на час. Добавить строку?",
                parent=self,
            ):
                return

        if end - start > timedelta(hours=24):
            if not messagebox.askyesno(
                "Интервал больше 24 часов",
                "Интервал больше 24 часов — будет найдено очень много файлов. Продолжить?",
                parent=self,
            ):
                return

        trks = self._parse_trks(self.trk_var.get())
        if trks is None:
            return

        for trk in trks:
            if not self.azs.cameras_for_trk(trk):
                messagebox.showerror(
                    "Ручное копирование",
                    f"Для ТРК {trk} нет настройки в azs.ini.",
                    parent=self,
                )
                return

        self.manual_rows.append(
            ManualRow(azscode=self.azs.number, start=start, end=end, trks=trks)
        )
        self._refresh_rows_tree()
        self.start_var.set("")
        self.end_var.set("")
        self.status_var.set(f"Строк в списке: {len(self.manual_rows)}")
        LOGGER.info(
            "Ручное копирование: добавлена строка АЗС %s %s..%s, ТРК %s",
            self.azs.number,
            start.strftime(DATETIME_FORMAT),
            end.strftime(DATETIME_FORMAT),
            ",".join(str(trk) for trk in trks),
        )

    def on_delete_selected(self) -> None:
        """Удалить выбранные строки из таблицы и из списка."""
        selection = self.rows_tree.selection()
        if not selection:
            messagebox.showinfo(
                "Ручное копирование", "Выберите строку в таблице.", parent=self
            )
            return

        children = list(self.rows_tree.get_children())
        indices = sorted(
            (children.index(item) for item in selection if item in children), reverse=True
        )
        for index in indices:
            if 0 <= index < len(self.manual_rows):
                del self.manual_rows[index]

        self._clear_results()
        self._refresh_rows_tree()
        self.status_var.set(f"Строк в списке: {len(self.manual_rows)}")

    def on_clear_all(self) -> None:
        """Удалить все строки и результаты."""
        self.manual_rows.clear()
        self._clear_results()
        self._refresh_rows_tree()
        self.status_var.set("Список строк пуст")

    def _refresh_rows_tree(self) -> None:
        """Перерисовать таблицу строк (порядок совпадает с ``manual_rows``)."""
        self.rows_tree.delete(*self.rows_tree.get_children())
        for index, row in enumerate(self.manual_rows, start=1):
            self.rows_tree.insert(
                "", "end", values=self._row_values(index, row), tags=self._row_tags(row)
            )

    def _row_values(self, index: int, row: ManualRow) -> Tuple[object, ...]:
        """Значения колонок таблицы строк."""
        return (
            index,
            "" if row.azscode is None else str(row.azscode),
            row.start_text,
            row.end_text,
            row.trks_text,
            row.status,
        )

    @staticmethod
    def _row_tags(row: ManualRow) -> Tuple[str, ...]:
        """Цвет строки по её статусу."""
        if row.status.startswith("ошибка"):
            return ("error",)
        if row.status.startswith("найдено") or row.status.startswith("скопировано"):
            return ("ok",)
        if row.status.startswith("пропущено"):
            return ("warn",)
        return ()

    def _update_row_status(self, position: int, text: str) -> None:
        """Обновить статус строки таблицы по её номеру (1-based)."""
        if not 1 <= position <= len(self.manual_rows):
            return
        row = self.manual_rows[position - 1]
        row.status = text
        children = self.rows_tree.get_children()
        if position - 1 < len(children):
            self.rows_tree.item(
                children[position - 1],
                values=self._row_values(position, row),
                tags=self._row_tags(row),
            )

    def _clear_results(self) -> None:
        """Очистить таблицу результатов."""
        self.results_tree.delete(*self.results_tree.get_children())

    # ------------------------------------------------------------------ #
    # Импорт/экспорт ручной таблицы в xlsx
    # ------------------------------------------------------------------ #
    def on_export_xlsx(self) -> None:
        """Сохранить ручную таблицу в xlsx (колонки azscode/start/end/trk)."""
        if self._is_busy():
            messagebox.showwarning(
                "Идёт обработка",
                "Дождитесь завершения или нажмите «Остановить».",
                parent=self,
            )
            return
        if not self.manual_rows:
            messagebox.showinfo("Экспорт в xlsx", "Нет строк для экспорта.", parent=self)
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path_text = filedialog.asksaveasfilename(
            title="Экспорт ручной таблицы в xlsx",
            initialdir=str(self.config.base_dir),
            initialfile=f"manual_{stamp}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Файлы Excel", "*.xlsx"), ("Все файлы", "*.*")],
            parent=self,
        )
        if not path_text:
            return
        path = Path(path_text)

        try:
            write_manual_xlsx(path, self.manual_rows, self.azs.number)
        except Exception as exc:  # XlsxError/OSError — показываем пользователю
            LOGGER.error("Ручное копирование: ошибка экспорта в %s: %s", path, exc)
            messagebox.showerror(
                "Экспорт в xlsx", f"Не удалось сохранить файл:\n{exc}", parent=self
            )
            return

        LOGGER.info(
            "Ручное копирование: экспортировано строк %s в %s",
            len(self.manual_rows),
            path,
        )
        messagebox.showinfo("Экспорт в xlsx", f"Сохранено: {path}", parent=self)
        if messagebox.askyesno(
            "Экспорт в xlsx", "Открыть папку с файлом?", parent=self
        ):
            self._open_directory(path.parent)

    def on_import_xlsx(self) -> None:
        """Выбрать xlsx и добавить строки из него в ручную таблицу."""
        path_text = filedialog.askopenfilename(
            title="Импорт ручной таблицы из xlsx",
            initialdir=str(self.config.base_dir),
            filetypes=[("Файлы Excel", "*.xlsx"), ("Все файлы", "*.*")],
            parent=self,
        )
        if not path_text:
            return
        self.import_from_xlsx(Path(path_text))

    def import_from_xlsx(self, path: Path) -> bool:
        """Импортировать строки ручного xlsx, добавив их к текущим.

        :param path: путь к xlsx
        :return: ``True``, если файл распознан и строки добавлены
        """
        if self._is_busy():
            messagebox.showwarning(
                "Идёт обработка",
                "Дождитесь завершения или нажмите «Остановить».",
                parent=self,
            )
            return False
        path = Path(path)
        try:
            kind = detect_xlsx_format(path)
        except XlsxError as exc:
            LOGGER.error("Ручное копирование: не удалось проверить %s: %s", path, exc)
            messagebox.showerror("Импорт из xlsx", str(exc), parent=self)
            return False

        if kind == "jobs":
            messagebox.showerror(
                "Импорт из xlsx",
                "Это файл заданий (docdate/azscode). Загрузите его на вкладке «Задания».",
                parent=self,
            )
            return False
        if kind != "manual":
            messagebox.showerror(
                "Импорт из xlsx",
                "Не удалось распознать формат файла. "
                "Ожидаются колонки azscode/start/end/trk.",
                parent=self,
            )
            return False

        try:
            imported = read_manual_xlsx(path)
        except XlsxError as exc:
            LOGGER.error("Ручное копирование: ошибка импорта %s: %s", path, exc)
            messagebox.showerror("Импорт из xlsx", str(exc), parent=self)
            return False

        for row in imported:
            row.status = f"ошибка: {row.error}" if row.error else "ожидание"
            self.manual_rows.append(row)

        self._clear_results()
        self._refresh_rows_tree()
        foreign = sum(
            1
            for row in imported
            if row.azscode is not None and row.azscode != self.azs.number
        )
        self.status_var.set(f"Импортировано строк: {len(imported)}")
        LOGGER.info(
            "Ручное копирование: импортировано строк %s из %s (для других АЗС: %s)",
            len(imported),
            path,
            foreign,
        )

        if foreign:
            messagebox.showinfo(
                "Импорт из xlsx",
                f"Импортировано {len(imported)} строк. Из них {foreign} строк для "
                "других АЗС — они будут пропущены при копировании.",
                parent=self,
            )
        else:
            messagebox.showinfo(
                "Импорт из xlsx",
                f"Импортировано {len(imported)} строк.",
                parent=self,
            )
        return True

    def _open_directory(self, path: Path) -> None:
        """Открыть каталог в Проводнике."""
        try:
            Path(path).mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))  # type: ignore[attr-defined]
        except OSError as exc:
            messagebox.showerror(
                "Ошибка", f"Не удалось открыть {path}: {exc}", parent=self
            )

    # ------------------------------------------------------------------ #
    # Запуск поиска/копирования
    # ------------------------------------------------------------------ #
    def on_find_all(self) -> None:
        """Найти файлы по всем строкам (без копирования)."""
        self._start_run(scan_only=True)

    def on_copy_all(self) -> None:
        """Скопировать файлы по всем строкам."""
        self._start_run(scan_only=False)

    def on_stop(self) -> None:
        """Запросить остановку фоновой обработки."""
        if self._is_busy():
            self._stop_event.set()
            self.status_var.set("Останавливаю...")
            LOGGER.info("Ручное копирование: запрошена остановка.")

    def on_open_output(self) -> None:
        """Открыть ``output_root`` текущей АЗС в Проводнике."""
        self._open_directory(Path(self.azs.output_root))

    def _is_busy(self) -> bool:
        """Идёт ли сейчас фоновая обработка."""
        return self._busy

    def refresh_expiry_state(self) -> None:
        """Заблокировать/разблокировать «Найти все»/«Копировать все» по сроку."""
        try:
            blocked = bool(self.app.expiry_blocked())  # type: ignore[attr-defined]
        except Exception:  # приложение без предохранителя — просто работаем
            blocked = False
        state = "disabled" if blocked or self._busy else "normal"
        self.find_button.configure(state=state)
        self.copy_button.configure(state=state)

    def on_config_changed(self) -> None:
        """Перечитать сдвиг часов и сбросить результаты (смена АЗС, ini)."""
        self.refresh_clock_label()
        self._clear_results()
        for row in self.manual_rows:
            if row.error:
                row.status = f"ошибка: {row.error}"
            elif row.azscode is not None and row.azscode != self.azs.number:
                row.status = "пропущено (другая АЗС)"
            else:
                row.status = "ожидание"
        self._refresh_rows_tree()
        self.refresh_expiry_state()

    def _start_run(self, scan_only: bool) -> None:
        """Запустить поиск или копирование по всем строкам в фоновом потоке."""
        if not self.app.expiry_allows():  # type: ignore[attr-defined]
            return
        if self._is_busy():
            messagebox.showwarning(
                "Идёт обработка",
                "Дождитесь завершения или нажмите «Остановить».",
                parent=self,
            )
            return
        if not self.manual_rows:
            messagebox.showinfo(
                "Ручное копирование", "Нет строк для обработки.", parent=self
            )
            return

        config = self.config
        azs = self.azs
        # Сдвиг часов берётся из settings.ini (редактируется на вкладке «Настройки»).
        drift = float(getattr(config.finder, "clock_drift_minutes", 0.0) or 0.0)
        finder_cfg = replace(config.finder, clock_drift_minutes=drift)

        rows = list(self.manual_rows)

        def resolved_trks(row: ManualRow) -> List[int]:
            """ТРК строки; пустой список = все ТРК текущей АЗС."""
            return row.trks or sorted(azs.trk_map.keys())

        for row in rows:
            if row.error:
                row.status = f"ошибка: {row.error}"
            elif row.azscode is not None and row.azscode != azs.number:
                row.status = "пропущено (другая АЗС)"
            else:
                row.status = "ожидание"
        self._refresh_rows_tree()
        self._clear_results()

        try:
            self._stop_event.clear()
            total_cells = max(1, sum(len(resolved_trks(row)) for row in rows))
            self.progress.configure(maximum=total_cells, value=0)
            mode_text = "Поиск" if scan_only else "Копирование"
            self.status_var.set(f"{mode_text}: 0 из {total_cells}")
            LOGGER.info(
                "Ручное копирование (%s): АЗС %s, строк %s, output_root=%s, сдвиг %s мин",
                mode_text,
                azs.number,
                len(rows),
                azs.output_root,
                format_minutes(drift),
            )
            LOGGER.info(
                "manual: before set_busy, rows=%s, trks_total=%s",
                len(rows),
                total_cells,
            )
            self._set_busy(True)
            LOGGER.info("manual: set_busy done, starting worker thread")

            def emit(kind: str, payload: object) -> None:
                self._event_queue.put((kind, payload))

            def run() -> Dict[str, object]:
                LOGGER.info(
                    "manual: run() entered, rows=%s, scan_only=%s",
                    len(rows),
                    scan_only,
                )
                video_roots = drives.resolve_video_roots(config, azs)
                include_audio = config.general.include_audio
                done_cells = 0
                total_found = 0
                total_copied = 0
                total_already = 0
                total_overwritten = 0
                total_errors = 0
                all_records: List[CopyRecord] = []

                for position, row in enumerate(rows, start=1):
                    if self._stop_event.is_set():
                        for rest in range(position, len(rows) + 1):
                            emit("row_status", (rest, "пропущено"))
                        break

                    if row.error:
                        emit("row_status", (position, f"ошибка: {row.error}"))
                        done_cells += max(1, len(resolved_trks(row)))
                        emit("progress", (done_cells, total_cells))
                        continue

                    if row.azscode is not None and row.azscode != azs.number:
                        emit("row_status", (position, "пропущено (другая АЗС)"))
                        done_cells += max(1, len(resolved_trks(row)))
                        emit("progress", (done_cells, total_cells))
                        continue

                    row_found = 0
                    row_records: List[CopyRecord] = []

                    for trk in resolved_trks(row):
                        if self._stop_event.is_set():
                            break
                        cameras = azs.cameras_for_trk(trk)
                        if not cameras:
                            continue

                        files: List[FoundFile] = list(
                            find_files_for_cameras(
                                azs=azs,
                                cameras=cameras,
                                start=row.start,
                                end=row.end,
                                finder_cfg=finder_cfg,
                                video_roots=video_roots,
                                include_audio=include_audio,
                            )
                        )
                        row_found += len(files)
                        total_found += len(files)

                        if files and scan_only:
                            for item in files:
                                record = CopyRecord(
                                    row_index=position,
                                    azs=azs.number,
                                    docdate=row.start,
                                    trk=trk,
                                    camera=item.camera,
                                    kind=item.kind,
                                    src=item.src,
                                    dst=destination_for(
                                        item.src, item.video_root, azs.output_root
                                    ),
                                    status=STATUS_WILL_COPY,
                                )
                                row_records.append(record)
                                all_records.append(record)
                                emit("record", (record, item.mtime_text))
                        elif files:
                            mtime_by_src = {
                                str(item.src).lower(): item.mtime_text for item in files
                            }

                            def on_file(
                                record: CopyRecord,
                                mapping: Dict[str, str] = mtime_by_src,
                            ) -> None:
                                key = str(record.src).lower() if record.src else ""
                                emit("record", (record, mapping.get(key, "—")))

                            copied = copy_found_files(
                                files=files,
                                azs=azs,
                                general_cfg=config.general,
                                stop_event=self._stop_event,
                                on_file=on_file,
                                row_index=position,
                                trk=trk,
                                docdate=row.start,
                            )
                            row_records.extend(copied)
                            all_records.extend(copied)

                        done_cells += 1
                        emit("progress", (done_cells, total_cells))

                    if scan_only:
                        status_text = (
                            f"найдено {row_found}" if row_found else "пропущено"
                        )
                    else:
                        copied_n = sum(
                            1 for rec in row_records if rec.status == STATUS_COPIED
                        )
                        overwritten_n = sum(
                            1
                            for rec in row_records
                            if rec.status == STATUS_SIZE_MISMATCH
                        )
                        already_n = sum(
                            1
                            for rec in row_records
                            if rec.status == STATUS_ALREADY_EXISTS
                        )
                        errors_n = sum(
                            1 for rec in row_records if rec.status == STATUS_ERROR
                        )
                        total_copied += copied_n
                        total_overwritten += overwritten_n
                        total_already += already_n
                        total_errors += errors_n
                        if errors_n:
                            status_text = "ошибка"
                        elif copied_n or overwritten_n or already_n:
                            status_text = (
                                f"скопировано {copied_n + overwritten_n + already_n}"
                            )
                        else:
                            status_text = "пропущено"
                    emit("row_status", (position, status_text))

                report_path: Optional[Path] = None
                if all_records:
                    report_path = self._write_report(azs.output_root, all_records)

                LOGGER.info(
                    "manual: run() finished, found=%s, copied=%s, errors=%s",
                    total_found,
                    total_copied,
                    total_errors,
                )
                return {
                    "scan_only": scan_only,
                    "rows": len(rows),
                    "found": total_found,
                    "copied": total_copied,
                    "already": total_already,
                    "overwritten": total_overwritten,
                    "errors": total_errors,
                    "report": report_path,
                    "stopped": self._stop_event.is_set(),
                }

            def worker() -> None:
                LOGGER.info(
                    "manual: worker thread started, name=%s",
                    threading.current_thread().name,
                )
                try:
                    self._event_queue.put(("done", run()))
                except Exception as exc:  # защита от падения фонового потока
                    LOGGER.exception(
                        "Ручное копирование: ошибка выполнения: %s", exc
                    )
                    self._event_queue.put(("error", str(exc)))

            self._worker = threading.Thread(
                target=worker, name="FuelVideoManualWorker", daemon=True
            )
            self._worker.start()
        except Exception as exc:
            LOGGER.exception("manual: ошибка запуска задачи: %s", exc)
            self._set_busy(False)
            self.status_var.set("Ошибка запуска — см. лог")
            messagebox.showerror("Ошибка запуска", str(exc), parent=self)
            return

    def _set_busy(self, busy: bool) -> None:
        """Заблокировать элементы управления на время выполнения."""
        try:
            self._busy = busy
            state = "disabled" if busy else "normal"
            for widget in (
                self.add_button,
                self.delete_button,
                self.clear_button,
                self.export_button,
                self.import_button,
            ):
                widget.configure(state=state)
            self.stop_button.configure(state="normal" if busy else "disabled")
            self.refresh_expiry_state()
        except Exception as exc:
            LOGGER.exception("manual: ошибка в _set_busy: %s", exc)

    # ------------------------------------------------------------------ #
    # Отчёт CSV
    # ------------------------------------------------------------------ #
    def _write_report(
        self, output_root: Path, records: Sequence[CopyRecord]
    ) -> Optional[Path]:
        """Записать отчёт ``manual_YYYYMMDD_HHMMSS.csv`` в ``output_root``."""
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = Path(output_root) / f"manual_{stamp}.csv"
        try:
            Path(output_root).mkdir(parents=True, exist_ok=True)
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow(MANUAL_REPORT_COLUMNS)
                for record in records:
                    writer.writerow(
                        [
                            record.row_index,
                            record.kind,
                            record.camera,
                            "" if record.src is None else str(record.src),
                            "" if record.dst is None else str(record.dst),
                            record.status,
                            record.error,
                        ]
                    )
        except OSError as exc:
            LOGGER.error(
                "Ручное копирование: не удалось записать отчёт %s: %s", path, exc
            )
            return None
        LOGGER.info("Ручное копирование: отчёт %s", path)
        return path

    # ------------------------------------------------------------------ #
    # Обработка сообщений фонового потока
    # ------------------------------------------------------------------ #
    @staticmethod
    def _display_status(status: str) -> str:
        """Русская подпись статуса записи."""
        if status == STATUS_WILL_COPY:
            return WILL_COPY_LABEL
        return status_label(status)

    def _append_record(self, record: CopyRecord, mtime_text: str) -> None:
        """Добавить запись в таблицу результатов."""
        status = self._display_status(record.status)
        if record.error:
            status = f"{status}: {record.error}"
        tag = "error" if record.status == STATUS_ERROR else ""
        self.results_tree.insert(
            "",
            "end",
            values=(
                record.row_index,
                record.kind,
                record.camera,
                mtime_text,
                str(record.src) if record.src else "—",
                str(record.dst) if record.dst else "—",
                status,
            ),
            tags=(tag,) if tag else (),
        )

    def _poll_events(self) -> None:
        """Обработать сообщения фонового потока (вызывается из GUI-потока)."""
        try:
            while True:
                kind, payload = self._event_queue.get_nowait()
                if kind == "record":
                    record, mtime_text = payload  # type: ignore[misc]
                    self._append_record(record, str(mtime_text))
                elif kind == "row_status":
                    position, text = payload  # type: ignore[misc]
                    self._update_row_status(int(position), str(text))
                elif kind == "progress":
                    done, total = payload  # type: ignore[misc]
                    self.progress.configure(maximum=max(1, total), value=done)
                    self.status_var.set(f"Обработано {done} из {total}")
                elif kind == "done":
                    self._on_done(dict(payload))  # type: ignore[arg-type]
                elif kind == "error":
                    self._on_error(str(payload))
        except queue.Empty:
            pass
        self.after(150, self._poll_events)

    def _on_done(self, summary: Dict[str, object]) -> None:
        """Завершить обработку: показать итог и путь к отчёту."""
        self._set_busy(False)

        scan_only = bool(summary.get("scan_only"))
        found = summary.get("found", 0)
        copied = summary.get("copied", 0)
        already = summary.get("already", 0)
        overwritten = summary.get("overwritten", 0)
        errors = summary.get("errors", 0)
        report = summary.get("report")
        stopped = bool(summary.get("stopped"))

        if scan_only:
            text = f"Поиск завершён: найдено файлов {found}"
        else:
            text = (
                f"Копирование завершено: скопировано {copied}, уже было {already}, "
                f"перезаписано {overwritten}, ошибок {errors}"
            )
        if stopped:
            text = "Остановлено пользователем. " + text
        if report is not None:
            text += f". Отчёт: {report}"
        self.status_var.set(text)

        try:
            self.progress.configure(value=float(self.progress.cget("maximum")))
        except (tk.TclError, ValueError):
            pass
        LOGGER.info("Ручное копирование: %s", text)

    def _on_error(self, message: str) -> None:
        """Сообщить об ошибке фоновой обработки."""
        self._set_busy(False)
        self.status_var.set("Обработка прервана ошибкой")
        messagebox.showerror("Ошибка обработки", message, parent=self)
