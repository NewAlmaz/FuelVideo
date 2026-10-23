"""Графический интерфейс приложения (tkinter/ttk).

Три вкладки:

* «Задания» — выбор xlsx, таблица строк, сканирование и копирование;
* «Настройки» — read-only отображение settings.ini/azs.ini и проверка путей;
* «Лог» — журнал работы приложения.

Копирование выполняется в фоновом потоке; обмен с GUI идёт через очереди и
``root.after``, поэтому окно не зависает.
"""

from __future__ import annotations

import logging
import os
import queue
import threading
import tkinter as tk
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Sequence, Tuple

import drives
from config import AppConfig, AzsConfig, ConfigError, get_base_dir, load_config
from expiry import (
    check_clock_rollback,
    expiry_text,
    is_expired,
    rollback_message,
    should_block,
    verify_unlock_password,
)
from license import UNLOCK_PASSWORD_HASH
from copier import (
    JobRunner,
    RowResult,
    destination_for,
    status_label,
    write_reports,
)
from video_finder import FoundFile
from xlsx_reader import RowData, XlsxError, read_rows

LOGGER = logging.getLogger(__name__)

TASKS_COLUMNS = ("azs", "docdate", "trk", "cameras", "start", "end", "video", "audio", "status")
TASKS_HEADINGS = {
    "azs": "АЗС",
    "docdate": "Дата/время заправки",
    "trk": "ТРК",
    "cameras": "Камеры",
    "start": "Начало",
    "end": "Конец",
    "video": "Видео",
    "audio": "Звук",
    "status": "Статус",
}
TASKS_WIDTHS = {
    "azs": 60,
    "docdate": 150,
    "trk": 50,
    "cameras": 110,
    "start": 150,
    "end": 150,
    "video": 70,
    "audio": 70,
    "status": 240,
}


#: Имя файла диагностики запуска (лежит рядом с ini-файлами, пишется в append-режиме).
STARTUP_DEBUG_FILE = "startup_debug.log"


def _startup_debug_path() -> Path:
    """Путь к файлу startup_debug.log."""
    return Path(get_base_dir()) / STARTUP_DEBUG_FILE


def log_startup(message: str) -> None:
    """Дописать строку в startup_debug.log (append).

    :param message: текст без отметки времени
    """
    try:
        with _startup_debug_path().open("a", encoding="utf-8") as stream:
            stream.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {message}\n")
    except Exception as exc:  # лог не должен ломать запуск программы
        LOGGER.error("Не удалось записать %s: %s", STARTUP_DEBUG_FILE, exc)


def log_startup_exception(prefix: str) -> None:
    """Дописать в startup_debug.log сообщение и текущий трейсбек."""
    try:
        with _startup_debug_path().open("a", encoding="utf-8") as stream:
            stream.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {prefix}\n")
            traceback.print_exc(file=stream)
    except Exception as exc:  # лог не должен ломать запуск программы
        LOGGER.error("Не удалось записать %s: %s", STARTUP_DEBUG_FILE, exc)


def show_error(parent: Optional[tk.Misc], title: str, message: str) -> None:
    """Показать модальное окно ошибки и записать его в startup_debug.log."""
    LOGGER.error("%s: %s", title, message)
    log_startup(f"error: {title}: {message}")
    try:
        messagebox.showerror(title, message, parent=parent)
    except Exception:
        log_startup_exception("error: не удалось показать messagebox")


def azs_option_label(azs: AzsConfig) -> str:
    """Подпись строки списка выбора АЗС, например «АЗС 38 — D:\\VideoExport\\AZS38».

    Строится только из номера секции и пути из azs.ini: диск НЕ опрашивается.
    """
    if azs.output_root_explicit:
        return f"АЗС {azs.number} — {azs.output_root}"
    return f"АЗС {azs.number} — (по умолчанию)"


def _reset_topmost(window: tk.Misc) -> None:
    """Снять признак «поверх всех окон» после первого показа диалога."""
    try:
        window.attributes("-topmost", False)
    except tk.TclError:
        pass


def ask_azs_dialog(parent: tk.Misc, options: Sequence[Tuple[int, str]]) -> Optional[int]:
    """Показать модальный диалог выбора АЗС.

    :param parent: уже созданное главное окно (может быть скрыто)
    :param options: список пар ``(номер АЗС, подпись)``
    :return: номер выбранной АЗС или ``None`` (кнопка «Выход», крестик, ошибка)
    """
    result: Optional[int] = None
    log_startup(f"dialog: opening, {len(options)} options")
    try:
        dlg = tk.Toplevel(parent)
        dlg.title("FuelVideo — выберите АЗС")
        # transient ставим только для ВИДИМОГО родителя: на Windows owned-окно
        # скрытого (withdraw) родителя никогда не мапится, и тогда
        # wait_visibility()/wait_window() зависают навсегда — окна не видно.
        # Модальность в этом случае обеспечивают grab_set() и wait_window().
        if parent.winfo_viewable():
            dlg.transient(parent)
        dlg.resizable(False, False)

        frame = ttk.Frame(dlg, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Выберите АЗС, с которой будет работать программа:",
        ).pack(anchor="w", pady=(0, 8))

        listbox = tk.Listbox(
            frame,
            height=min(10, max(3, len(options))),
            width=70,
            exportselection=False,
        )
        for _number, label in options:
            listbox.insert("end", label)
        listbox.pack(fill="both", expand=True)
        if options:
            listbox.selection_set(0)
            listbox.see(0)
            listbox.focus_set()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(10, 0))

        def on_ok(_event: Optional[tk.Event] = None) -> None:
            """Зафиксировать выбор пользователя."""
            nonlocal result
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Выбор АЗС", "Выберите АЗС из списка.", parent=dlg)
                return
            result = options[selection[0]][0]
            log_startup(f"dialog: user selected {result}")
            dlg.destroy()

        def on_cancel(_event: Optional[tk.Event] = None) -> None:
            """Закрыть диалог по кнопке «Выход»."""
            nonlocal result
            result = None
            log_startup("dialog: user cancelled")
            dlg.destroy()

        def on_window_close() -> None:
            """Закрыть диалог крестиком."""
            nonlocal result
            result = None
            log_startup("dialog: closed by window manager")
            dlg.destroy()

        ttk.Button(buttons, text="Выбрать", command=on_ok).pack(side="right")
        ttk.Button(buttons, text="Выход", command=on_cancel).pack(side="right", padx=(0, 8))

        dlg.protocol("WM_DELETE_WINDOW", on_window_close)
        dlg.bind("<Return>", on_ok)
        dlg.bind("<Escape>", on_cancel)
        listbox.bind("<Double-Button-1>", on_ok)

        # Порядок строго важен: сначала окно отрисовывается и показывается,
        # только потом ставится grab_set и запускаются wait_visibility/wait_window.
        dlg.update_idletasks()
        dlg.deiconify()
        dlg.lift()
        dlg.attributes("-topmost", True)
        dlg.after(100, lambda: _reset_topmost(dlg))
        dlg.grab_set()
        dlg.wait_visibility()
        dlg.wait_window()
    except Exception:
        log_startup_exception("dialog: exception while showing the AZS dialog")
        LOGGER.exception("Ошибка диалога выбора АЗС")
        try:
            messagebox.showerror(
                "Ошибка выбора АЗС",
                "Не удалось показать диалог выбора АЗС.\n"
                f"Подробности записаны в {STARTUP_DEBUG_FILE}.",
                parent=parent,
            )
        except Exception:
            log_startup_exception("dialog: messagebox failed")
        return None
    return result


class DetailsWindow(tk.Toplevel):
    """Окно деталей по строке: найденные файлы и пары «источник -> назначение»."""

    EXAMPLE_LIMIT = 30

    def __init__(
        self,
        master: tk.Misc,
        row: RowData,
        result: Optional[RowResult],
        azs: AzsConfig,
    ) -> None:
        super().__init__(master)
        self.title(f"Детали строки {row.index} — АЗС {azs.number}")
        self.geometry("1180x640")
        self.result = result

        info_lines = [
            f"Строка xlsx: {row.index}    АЗС: {row.azs}    ТРК: {row.trk}",
            f"Дата/время заправки: {row.docdate_text}    Карта: {row.cardcode or '—'}",
            f"Камеры: {', '.join(result.cameras) if result and result.cameras else '—'}",
        ]
        if result is not None:
            info_lines.append(f"Интервал: {result.start_text} .. {result.end_text}")
            info_lines.append(f"Итог: {result.summary_text}; статус: {result.status_text}")
            if result.error:
                info_lines.append(f"Ошибка: {result.error}")

        header = ttk.Label(self, text="\n".join(info_lines), justify="left", padding=(8, 8))
        header.pack(fill="x")

        columns = ("kind", "camera", "mtime", "src", "dst", "status")
        headings = {
            "kind": "Тип",
            "camera": "Камера",
            "mtime": "Время файла",
            "src": "Источник",
            "dst": "Назначение",
            "status": "Статус",
        }
        widths = {"kind": 60, "camera": 65, "mtime": 140, "src": 420, "dst": 420, "status": 180}

        tree_frame = ttk.Frame(self, padding=(8, 0))
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12)
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self,
            text=f"Примеры путей (до {self.EXAMPLE_LIMIT} на камеру) и пары «источник -> назначение»:",
            padding=(8, 6),
        ).pack(fill="x")

        text_frame = ttk.Frame(self, padding=(8, 0, 8, 8))
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, height=12, wrap="none")
        text_vsb = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=text_vsb.set)
        self.text.pack(side="left", fill="both", expand=True)
        text_vsb.pack(side="right", fill="y")

        self._fill(row, result, azs)
        self.text.configure(state="disabled")

    def _fill(self, row: RowData, result: Optional[RowResult], azs: AzsConfig) -> None:
        """Заполнить таблицу и текстовую область."""
        mtime_by_src: Dict[str, str] = {}
        if result is not None:
            for item in result.files:
                mtime_by_src[str(item.src).lower()] = item.mtime_text

        if result is not None and result.records:
            for record in result.records:
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        record.kind,
                        record.camera,
                        mtime_by_src.get(str(record.src).lower(), "—") if record.src else "—",
                        str(record.src) if record.src else "—",
                        str(record.dst) if record.dst else "—",
                        status_label(record.status) + (f" ({record.error})" if record.error else ""),
                    ),
                )
        elif result is not None and result.files:
            for item in result.files:
                planned = destination_for(item.src, item.video_root, azs.output_root)
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        item.kind,
                        item.camera,
                        item.mtime_text,
                        str(item.src),
                        str(planned),
                        "будет скопировано",
                    ),
                )
        else:
            self.tree.insert("", "end", values=("—", "—", "—", "файлы не найдены", "—", "не найдено"))

        if result is None or not result.files:
            self.text.insert("end", "Файлы за интервал не найдены.\n")
            if result is not None and result.error:
                self.text.insert("end", f"Причина: {result.error}\n")
            return

        grouped: Dict[str, List[FoundFile]] = {}
        for item in result.files:
            grouped.setdefault(item.camera, []).append(item)

        for camera in sorted(grouped):
            files = grouped[camera]
            self.text.insert("end", f"=== Камера {camera}: найдено {len(files)} файлов ===\n")
            for item in files[: self.EXAMPLE_LIMIT]:
                planned = destination_for(item.src, item.video_root, azs.output_root)
                self.text.insert(
                    "end",
                    f"[{item.kind}] {item.mtime_text}\n    {item.src}\n -> {planned}\n",
                )
            if len(files) > self.EXAMPLE_LIMIT:
                self.text.insert(
                    "end", f"    ... и ещё {len(files) - self.EXAMPLE_LIMIT} файлов\n"
                )
            self.text.insert("end", "\n")


class ExporterApp(ttk.Frame):
    """Содержимое главного окна внутри уже созданного ``root``.

    Класс НЕ создаёт собственный ``tk.Tk`` и НЕ вызывает ``mainloop()``:
    главное окно создаётся в ``main.py`` (``root = tk.Tk()``), сюда оно
    передаётся готовым, а ``mainloop()`` запускается один раз в конце.
    """

    def __init__(
        self,
        root: tk.Tk,
        config: AppConfig,
        azs: AzsConfig,
        ui_queue: "queue.Queue[str]",
    ) -> None:
        super().__init__(root)
        self.root = root
        self.config_data = config
        self.azs = azs
        self.ui_queue = ui_queue

        self.rows: List[RowData] = []
        self.results: List[RowResult] = []
        self._results_by_index: Dict[int, RowResult] = {}
        self._item_by_index: Dict[int, str] = {}
        self.xlsx_path: Optional[Path] = None
        self._worker: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._event_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self._reports: Tuple[Optional[Path], Optional[Path]] = (None, None)
        # Мягкий предохранитель срока действия (см. license.py / expiry.py):
        # разблокировка живёт только до конца текущей сессии.
        self._unlocked: bool = False
        self._clock_rollback: bool = False
        self._saved_date: Optional[date] = None
        self._rollback_notified: bool = False

        APP_VERSION = "0.1.0"
        self.root.title(f"FuelVideo — копирование видео и звука из СВН «Интеллект» v{APP_VERSION}")
        self.root.geometry("1200x720")
        self.root.minsize(980, 600)

        self._build_ui()
        self.pack(fill="both", expand=True)
        self._update_azs_label()
        # Заполнение вкладки «Настройки» обращается к дискам (в т.ч. сетевым),
        # поэтому выполняем его уже после того, как окно покажется на экране.
        self.root.after(100, self._rebuild_settings)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(120, self._poll_events)
        self.root.after(200, self._poll_log)

        self._apply_expiry_state()

    # ------------------------------------------------------------------ #
    # Построение интерфейса
    # ------------------------------------------------------------------ #
    def _build_ui(self) -> None:
        """Создать вкладки и элементы управления."""
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

        self.tasks_tab = ttk.Frame(self.notebook, padding=6)
        self.settings_tab = ttk.Frame(self.notebook, padding=6)
        self.log_tab = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tasks_tab, text="Задания")
        self.notebook.add(self.settings_tab, text="Настройки")
        self.notebook.add(self.log_tab, text="Лог")

        self._build_tasks_tab()
        self._build_settings_tab()
        self._build_log_tab()

    def _build_tasks_tab(self) -> None:
        """Вкладка «Задания»."""
        top = ttk.Frame(self.tasks_tab)
        top.pack(fill="x", pady=(0, 6))

        self.azs_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.azs_var, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(top, text="Сменить АЗС", command=self.on_switch_azs).pack(side="left", padx=6)
        self.unlock_button = ttk.Button(
            top,
            text="Ввести код разблокировки",
            command=self.on_unlock,
            state="disabled",
        )
        self.unlock_button.pack(side="left", padx=(0, 6))
        ttk.Button(top, text="Выбрать xlsx", command=self.on_choose_xlsx).pack(side="left", padx=(12, 6))

        self.xlsx_var = tk.StringVar(value="файл xlsx не выбран")
        ttk.Label(top, textvariable=self.xlsx_var).pack(side="left", padx=6)

        tree_frame = ttk.Frame(self.tasks_tab)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=TASKS_COLUMNS, show="headings", height=16)
        for column in TASKS_COLUMNS:
            self.tree.heading(column, text=TASKS_HEADINGS[column])
            anchor = "center" if column in ("azs", "trk", "video", "audio") else "w"
            self.tree.column(column, width=TASKS_WIDTHS[column], anchor=anchor)
        self.tree.tag_configure("error", foreground="#b00020")
        self.tree.tag_configure("notfound", foreground="#8a6d00")
        self.tree.tag_configure("ok", foreground="#0a7d24")

        tasks_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tasks_hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tasks_vsb.set, xscrollcommand=tasks_hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tasks_vsb.grid(row=0, column=1, sticky="ns")
        tasks_hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.on_row_details)

        buttons = ttk.Frame(self.tasks_tab)
        buttons.pack(fill="x", pady=(6, 0))

        self.scan_button = ttk.Button(buttons, text="Сканировать", command=self.on_scan)
        self.scan_button.pack(side="left")
        self.copy_button = ttk.Button(buttons, text="Копировать", command=self.on_copy)
        self.copy_button.pack(side="left", padx=6)
        self.stop_button = ttk.Button(buttons, text="Остановить", command=self.on_stop, state="disabled")
        self.stop_button.pack(side="left", padx=(0, 6))
        ttk.Button(
            buttons, text="Показать примеры файлов", command=self.on_row_details
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            buttons, text="Открыть папку результата", command=self.on_open_output
        ).pack(side="left")

        bottom = ttk.Frame(self.tasks_tab)
        bottom.pack(fill="x", pady=(6, 0))

        self.progress = ttk.Progressbar(bottom, mode="determinate", maximum=100)
        self.progress.pack(side="left", fill="x", expand=True)

        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(bottom, textvariable=self.status_var, padding=(8, 0)).pack(side="left")

    def _build_settings_tab(self) -> None:
        """Вкладка «Настройки»."""
        buttons = ttk.Frame(self.settings_tab)
        buttons.pack(fill="x", pady=(0, 6))
        ttk.Button(buttons, text="Перечитать ini", command=self.on_reload_config).pack(side="left")
        ttk.Button(buttons, text="Проверить пути", command=self.on_check_paths).pack(side="left", padx=6)
        ttk.Button(buttons, text="Пересканировать диски", command=self.on_rescan_drives).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(buttons, text="Открыть settings.ini", command=self.on_open_settings_ini).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(buttons, text="Открыть azs.ini", command=self.on_open_azs_ini).pack(side="left")

        text_frame = ttk.Frame(self.settings_tab)
        text_frame.pack(fill="both", expand=True)
        self.settings_text = tk.Text(text_frame, wrap="word", state="disabled")
        settings_vsb = ttk.Scrollbar(text_frame, orient="vertical", command=self.settings_text.yview)
        self.settings_text.configure(yscrollcommand=settings_vsb.set)
        self.settings_text.pack(side="left", fill="both", expand=True)
        settings_vsb.pack(side="right", fill="y")
        self.settings_text.tag_configure("head", font=("Segoe UI", 10, "bold"))
        self.settings_text.tag_configure("ok", foreground="#0a7d24")
        self.settings_text.tag_configure("bad", foreground="#b00020")

    def _build_log_tab(self) -> None:
        """Вкладка «Лог»."""
        buttons = ttk.Frame(self.log_tab)
        buttons.pack(fill="x", pady=(0, 6))
        ttk.Button(buttons, text="Очистить", command=self.on_clear_log).pack(side="left")

        log_frame = ttk.Frame(self.log_tab)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, wrap="none", state="disabled")
        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_vsb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_vsb.pack(side="right", fill="y")

    # ------------------------------------------------------------------ #
    # Обработчики кнопок
    # ------------------------------------------------------------------ #
    def on_choose_xlsx(self) -> None:
        """Выбрать файл xlsx и загрузить строки."""
        initial_dir = str(self.xlsx_path.parent) if self.xlsx_path else str(Path.home())
        path_text = filedialog.askopenfilename(
            title="Выберите файл xlsx со списком заправок",
            initialdir=initial_dir,
            filetypes=[("Файлы Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if not path_text:
            return
        self.load_xlsx(Path(path_text))

    def load_xlsx(self, path: Path) -> None:
        """Прочитать xlsx и заполнить таблицу строк."""
        try:
            rows = read_rows(path)
        except XlsxError as exc:
            LOGGER.error("Не удалось прочитать xlsx: %s", exc)
            messagebox.showerror("Ошибка чтения xlsx", str(exc), parent=self)
            return

        self.xlsx_path = Path(path)
        self.rows = rows
        self.xlsx_var.set(str(path))
        self.results = []
        self._results_by_index.clear()
        self._reports = (None, None)
        self._populate_tree()
        self.status_var.set(f"Загружено строк: {len(rows)}")
        LOGGER.info("Загружен файл %s, строк: %s", path, len(rows))

    def on_switch_azs(self) -> None:
        """Открыть диалог смены АЗС."""
        if self._is_busy():
            messagebox.showwarning("Идёт обработка", "Дождитесь завершения или нажмите «Остановить».", parent=self)
            return
        options = [(azs.number, azs_option_label(azs)) for azs in self.config_data.azs_list]
        chosen_number = ask_azs_dialog(self.root, options)
        if chosen_number is None:
            return
        chosen = self.config_data.find_azs(chosen_number)
        if chosen is None:
            messagebox.showwarning(
                "Секция АЗС не найдена",
                f"В azs.ini нет секции [AZS_{chosen_number}].",
                parent=self,
            )
            return
        if chosen.number == self.azs.number:
            return
        self.set_azs(chosen)
        if self.rows:
            self.results = []
            self._results_by_index.clear()
            self._populate_tree()
        self.status_var.set(f"Выбрана АЗС {chosen.number}")

    def set_azs(self, azs: AzsConfig) -> None:
        """Переключить текущую АЗС и обновить интерфейс."""
        self.azs = azs
        drives.reset_cache()
        self._update_azs_label()
        self._rebuild_settings()
        self.root.title(f"FuelVideo — АЗС {azs.number}")
        LOGGER.info(
            "Выбрана АЗС: %s; output_root=%s; video_roots=%s",
            azs.number,
            azs.output_root,
            ", ".join(str(item) for item in azs.video_roots) if azs.video_roots else "auto",
        )

    def _expiry_allows(self) -> bool:
        """Разрешено ли действие мягким предохранителем срока действия.

        Перепроверяет даты и откат часов непосредственно перед запуском задачи.

        :return: ``True``, если работа разрешена; иначе показывает окно и
            возвращает ``False``
        """
        rollback, saved_date = check_clock_rollback()
        self._clock_rollback = rollback
        self._saved_date = saved_date
        if should_block(rollback, saved_date) and not self._unlocked:
            if is_expired():
                text = (
                    f"Срок действия программы истёк ({expiry_text()}). "
                    "Копирование заблокировано."
                )
            else:
                text = "Обнаружен откат системных часов. Копирование заблокировано."
            messagebox.showwarning(
                "Срок истёк",
                f"{text}\nВведите код разблокировки.",
                parent=self,
            )
            return False
        return True

    def on_scan(self) -> None:
        """Запустить сканирование без копирования."""
        if not self._expiry_allows():
            return
        self._start_job(scan_only=True)

    def on_copy(self) -> None:
        """Запустить копирование с подтверждением."""
        if not self._expiry_allows():
            return
        if self._is_busy():
            messagebox.showwarning("Идёт обработка", "Дождитесь завершения текущей задачи.", parent=self)
            return
        if not self.rows:
            messagebox.showwarning("Нет данных", "Сначала выберите файл xlsx.", parent=self)
            return
        answer = messagebox.askyesno(
            "Копирование",
            f"Скопировать файлы за {len(self.rows)} строк xlsx?\n\n"
            f"АЗС: {self.azs.number}\nКаталог результата: {self.azs.output_root}\n"
            f"Режим: {self.config_data.general.copy_mode}",
            parent=self,
        )
        if not answer:
            return
        self._start_job(scan_only=False)

    def on_stop(self) -> None:
        """Запросить остановку обработки."""
        if self._is_busy():
            self._stop_event.set()
            self.status_var.set("Останавливаю...")
            LOGGER.info("Запрошена остановка обработки.")

    def on_open_output(self) -> None:
        """Открыть каталог результата в проводнике."""
        target = Path(self.azs.output_root)
        try:
            target.mkdir(parents=True, exist_ok=True)
            os.startfile(str(target))  # type: ignore[attr-defined]
        except OSError as exc:
            messagebox.showerror("Ошибка", f"Не удалось открыть {target}: {exc}", parent=self)

    def on_row_details(self, _event: Optional[tk.Event] = None) -> None:
        """Открыть окно деталей по выбранной строке."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Детали", "Выберите строку в таблице.", parent=self)
            return
        item = selection[0]
        index = None
        for row_index, item_id in self._item_by_index.items():
            if item_id == item:
                index = row_index
                break
        if index is None:
            return
        row = next((item for item in self.rows if item.index == index), None)
        if row is None:
            return
        DetailsWindow(self, row, self._results_by_index.get(index), self.azs)

    def on_reload_config(self) -> None:
        """Перечитать ini-файлы."""
        try:
            new_config = load_config(self.config_data.base_dir)
        except ConfigError as exc:
            LOGGER.error("Ошибка настроек: %s", exc)
            messagebox.showerror("Ошибка настроек", str(exc), parent=self)
            return
        self.config_data = new_config
        azs = new_config.find_azs(self.azs.number)
        if azs is None:
            messagebox.showwarning(
                "Секция АЗС не найдена",
                f"В azs.ini больше нет секции [AZS_{self.azs.number}]. "
                "Выберите другую АЗС кнопкой «Сменить АЗС».",
                parent=self,
            )
            azs = new_config.azs_list[0]
        self.azs = azs
        self._update_azs_label()
        self._rebuild_settings()
        self.status_var.set("Настройки перечитаны")
        LOGGER.info("Настройки перечитаны из ini.")

    def on_check_paths(self) -> None:
        """Проверить существование video_roots и output_root."""
        self._rebuild_settings()
        roots = drives.resolve_video_roots(self.config_data, self.azs)
        existing = [root for root in roots if root.is_dir()]
        output_exists = Path(self.azs.output_root).exists()
        message = (
            f"Корни видеоархива: {len(existing)} из {len(roots)} доступны.\n"
            f"Каталог результата {self.azs.output_root}: "
            f"{'существует' if output_exists else 'будет создан при копировании'}."
        )
        LOGGER.info("Проверка путей: %s", message.replace("\n", " "))
        messagebox.showinfo("Проверка путей", message, parent=self)

    def on_rescan_drives(self) -> None:
        """Пересканировать диски на наличие каталогов VIDEO."""
        drives.reset_cache()
        roots = drives.discover_video_roots(
            subdir=self.config_data.general.video_root_subdir, force=True
        )
        self._rebuild_settings()
        self.status_var.set(f"Обнаружено корней VIDEO: {len(roots)}")
        LOGGER.info("Повторное сканирование дисков: найдено %s корней VIDEO.", len(roots))

    def on_open_settings_ini(self) -> None:
        """Открыть settings.ini в ассоциированной программе."""
        self._open_path(self.config_data.settings_path)

    def on_open_azs_ini(self) -> None:
        """Открыть azs.ini в ассоциированной программе."""
        self._open_path(self.config_data.azs_path)

    def on_clear_log(self) -> None:
        """Очистить окно лога."""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ------------------------------------------------------------------ #
    # Запуск обработки
    # ------------------------------------------------------------------ #
    def _start_job(self, scan_only: bool) -> None:
        """Запустить фоновую обработку строк xlsx."""
        if self._is_busy():
            messagebox.showwarning("Идёт обработка", "Дождитесь завершения текущей задачи.", parent=self)
            return
        if not self.rows:
            messagebox.showwarning("Нет данных", "Сначала выберите файл xlsx.", parent=self)
            return

        self._stop_event.clear()
        self.results = []
        self._results_by_index.clear()
        self._reports = (None, None)
        self._populate_tree()

        self.progress.configure(maximum=max(1, len(self.rows)), value=0)
        mode_text = "Сканирование" if scan_only else "Копирование"
        self.status_var.set(f"{mode_text}: 0 из {len(self.rows)}")
        LOGGER.info(
            "%s: АЗС %s, строк %s, output_root=%s, интервал %s",
            mode_text,
            self.azs.number,
            len(self.rows),
            self.azs.output_root,
            self.config_data.general.interval_text,
        )

        self._set_buttons_enabled(False)
        self.stop_button.configure(state="normal")

        config = self.config_data
        azs = self.azs
        rows = list(self.rows)
        runner = JobRunner(config, azs)

        def worker() -> None:
            try:
                results = runner.run(
                    rows,
                    scan_only=scan_only,
                    stop_event=self._stop_event,
                    on_row=lambda result: self._event_queue.put(("row", result)),
                    on_progress=lambda done, total: self._event_queue.put(
                        ("progress", (done, total))
                    ),
                )
                self._event_queue.put(("done", results))
            except Exception as exc:  # защита от падения фонового потока
                LOGGER.exception("Ошибка обработки: %s", exc)
                self._event_queue.put(("error", str(exc)))

        self._worker = threading.Thread(target=worker, name="FuelVideoWorker", daemon=True)
        self._worker.start()

    def _is_busy(self) -> bool:
        """Идёт ли сейчас фоновая обработка."""
        return self._worker is not None and self._worker.is_alive()

    def _expiry_blocked(self) -> bool:
        """Заблокирована ли работа мягким предохранителем прямо сейчас."""
        return should_block(self._clock_rollback, self._saved_date) and not self._unlocked

    def _set_buttons_enabled(self, enabled: bool) -> None:
        """Заблокировать/разблокировать кнопки запуска.

        Блокировка по сроку действия имеет приоритет: после завершения задачи
        кнопки не включаются, если предохранитель сработал.
        """
        state = "normal" if enabled and not self._expiry_blocked() else "disabled"
        self.scan_button.configure(state=state)
        self.copy_button.configure(state=state)

    # ------------------------------------------------------------------ #
    # Мягкий предохранитель срока действия
    # ------------------------------------------------------------------ #
    def _apply_expiry_state(self) -> None:
        """Применить состояние предохранителя к кнопкам и строке статуса.

        Даты проверяются здесь же (часы могли откатить между запусками),
        поэтому метод вызывается и при старте, и после разблокировки.

        Откат часов сам по себе не блокирует: блокировка включается только если
        реальный срок вышел или до отката дата уже была после ``EXPIRY_DATE``.
        """
        rollback, saved_date = check_clock_rollback()
        self._clock_rollback = rollback
        self._saved_date = saved_date
        blocked = should_block(rollback, saved_date) and not self._unlocked

        if blocked:
            self.copy_button.configure(state="disabled")
            self.scan_button.configure(state="disabled")
            self.unlock_button.configure(state="normal")
            if is_expired():
                self.status_var.set(
                    f"Срок действия истёк ({expiry_text()}). Копирование заблокировано."
                )
            else:
                self.status_var.set(
                    "Обнаружен откат системных часов. Копирование заблокировано."
                )
            LOGGER.warning(
                "expiry: блокировка (expired=%s, clock_rollback=%s, saved=%s)",
                is_expired(),
                rollback,
                saved_date,
            )
        else:
            self.copy_button.configure(state="normal")
            self.scan_button.configure(state="normal")
            self.unlock_button.configure(state="disabled")
            if rollback:
                self.status_var.set(
                    rollback_message(rollback, saved_date) + " Работа продолжается."
                )
            elif self._unlocked:
                self.status_var.set("Разблокировано до конца сессии")
            else:
                self.status_var.set("Готов к работе")
            LOGGER.info(
                "expiry: активен %s, clock_rollback=%s, saved=%s",
                expiry_text(),
                rollback,
                saved_date,
            )

        # Предупреждение об откате часов — один раз за запуск (не при смене АЗС).
        if rollback and not self._rollback_notified:
            self._rollback_notified = True
            self.root.after(
                300, lambda: self._notify_rollback(blocked, rollback, saved_date)
            )

    def _notify_rollback(
        self, blocked: bool, rollback: bool, saved_date: Optional[date]
    ) -> None:
        """Показать пользователю предупреждение об обнаруженном откате часов."""
        message = rollback_message(rollback, saved_date)
        try:
            if blocked:
                messagebox.showwarning(
                    "Откат часов",
                    message + "\nКопирование заблокировано.",
                    parent=self,
                )
            else:
                messagebox.showinfo(
                    "Откат часов",
                    message + "\nРабота продолжается, но проверьте дату на компьютере.",
                    parent=self,
                )
        except tk.TclError as exc:  # сообщение не должно ломать главное окно
            LOGGER.error("Не удалось показать предупреждение об откате часов: %s", exc)

    def on_unlock(self) -> None:
        """Запросить код разблокировки (модальный диалог)."""
        if not UNLOCK_PASSWORD_HASH:
            messagebox.showerror(
                "Разблокировка недоступна",
                "В сборке не задан пароль разблокировки.",
                parent=self,
            )
            return

        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("FuelVideo — разблокировка")
            if self.root.winfo_viewable():
                dialog.transient(self.root)
            dialog.resizable(False, False)

            frame = ttk.Frame(dialog, padding=12)
            frame.pack(fill="both", expand=True)
            ttk.Label(
                frame,
                text="Введите код разблокировки:",
            ).pack(anchor="w", pady=(0, 6))

            password_var = tk.StringVar()
            entry = ttk.Entry(frame, textvariable=password_var, show="*", width=40)
            entry.pack(fill="x")
            entry.focus_set()

            def try_unlock() -> None:
                """Проверить код и при успехе снять блокировку до конца сессии."""
                if verify_unlock_password(password_var.get()):
                    self._unlocked = True
                    self._clock_rollback = False
                    LOGGER.info("expiry: разблокировано по паролю")
                    _close()
                    self._apply_expiry_state()
                    messagebox.showinfo(
                        "Разблокировка",
                        "Программа разблокирована до конца сессии.",
                        parent=self,
                    )
                else:
                    LOGGER.warning("expiry: неверный код")
                    messagebox.showerror(
                        "Ошибка",
                        "Неверный код разблокировки.",
                        parent=dialog,
                    )
                    entry.select_range(0, "end")
                    entry.focus_set()

            def _close() -> None:
                """Закрыть диалог, сняв захват ввода."""
                try:
                    dialog.grab_release()
                except tk.TclError:
                    pass
                dialog.destroy()

            buttons = ttk.Frame(frame)
            buttons.pack(fill="x", pady=(10, 0))
            ttk.Button(buttons, text="Разблокировать", command=try_unlock).pack(side="left")
            ttk.Button(buttons, text="Отмена", command=_close).pack(side="left", padx=6)

            entry.bind("<Return>", lambda _event: try_unlock())
            dialog.protocol("WM_DELETE_WINDOW", _close)
            dialog.grab_set()
            dialog.wait_window()
        except tk.TclError as exc:  # диалог не должен ломать главное окно
            LOGGER.error("Не удалось открыть диалог разблокировки: %s", exc)

    # ------------------------------------------------------------------ #
    # Опрос очередей
    # ------------------------------------------------------------------ #
    def _poll_events(self) -> None:
        """Обработать сообщения фонового потока."""
        try:
            while True:
                kind, payload = self._event_queue.get_nowait()
                if kind == "row" and isinstance(payload, RowResult):
                    self._apply_row_result(payload)
                elif kind == "progress":
                    done, total = payload  # type: ignore[misc]
                    self.progress.configure(maximum=max(1, total), value=done)
                    self.status_var.set(f"Обработано {done} из {total}")
                elif kind == "done":
                    self._on_job_done(list(payload))  # type: ignore[arg-type]
                elif kind == "error":
                    self._on_job_error(str(payload))
        except queue.Empty:
            pass
        self.root.after(120, self._poll_events)

    def _poll_log(self) -> None:
        """Перелить строки лога из очереди в текстовое поле."""
        messages: List[str] = []
        try:
            while True:
                messages.append(self.ui_queue.get_nowait())
        except queue.Empty:
            pass
        if messages:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", "\n".join(messages) + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(200, self._poll_log)

    # ------------------------------------------------------------------ #
    # Обновление таблицы
    # ------------------------------------------------------------------ #
    def _populate_tree(self) -> None:
        """Заполнить таблицу строками xlsx в статусе «ожидание»."""
        self.tree.delete(*self.tree.get_children())
        self._item_by_index.clear()
        for row in self.rows:
            item = self.tree.insert("", "end", values=self._row_values(row, None), tags=())
            self._item_by_index[row.index] = item

    def _row_values(self, row: RowData, result: Optional[RowResult]) -> Tuple[object, ...]:
        """Значения колонок таблицы для строки xlsx."""
        cameras = ""
        start_text = "—"
        end_text = "—"
        if row.azs is not None and row.trk is not None and not row.error:
            configured = self.azs.cameras_for_trk(int(row.trk))
            cameras = ",".join(configured) if configured else "нет настройки"

        if result is not None:
            cameras = result.cameras_text or cameras
            start_text = result.start_text
            end_text = result.end_text
            video = str(result.found_video)
            audio = str(result.found_audio)
            status = result.status_text
            if result.error:
                status = f"{status}: {result.error}"
        elif row.docdate is not None:
            start_text = (
                row.docdate - timedelta(minutes=self.config_data.general.minutes_before)
            ).strftime("%d.%m.%Y %H:%M:%S")
            end_text = (
                row.docdate + timedelta(minutes=self.config_data.general.minutes_after)
            ).strftime("%d.%m.%Y %H:%M:%S")
            video = "-"
            audio = "-"
            status = "ожидание"
            if row.error:
                status = f"ошибка: {row.error}"
        else:
            video = "-"
            audio = "-"
            status = f"ошибка: {row.error}" if row.error else "ожидание"

        return (
            "" if row.azs is None else row.azs,
            row.docdate_text,
            "" if row.trk is None else row.trk,
            cameras,
            start_text,
            end_text,
            video,
            audio,
            status,
        )

    def _apply_row_result(self, result: RowResult) -> None:
        """Обновить строку таблицы по результату обработки."""
        self._results_by_index[result.row.index] = result
        item = self._item_by_index.get(result.row.index)
        if item is None:
            return
        tag = ""
        if result.status == "error":
            tag = "error"
        elif result.status == "not_found":
            tag = "notfound"
        elif result.status in ("copied", "found"):
            tag = "ok"
        self.tree.item(item, values=self._row_values(result.row, result), tags=(tag,) if tag else ())

    def _on_job_done(self, results: List[RowResult]) -> None:
        """Завершить обработку: записать отчёты и показать итог."""
        self.results = results
        self._set_buttons_enabled(True)
        self.stop_button.configure(state="disabled")

        records_path, summary_path = write_reports(self.azs.output_root, results)
        self._reports = (records_path, summary_path)

        total_copied = sum(item.copied for item in results)
        total_already = sum(item.already_exists for item in results)
        total_overwritten = sum(item.overwritten for item in results)
        total_not_found = sum(item.not_found for item in results)
        errors = [item for item in results if item.status == "error"]

        summary = (
            f"Строк обработано: {len(results)}\n"
            f"Скопировано файлов: {total_copied}\n"
            f"Уже было: {total_already}\n"
            f"Перезаписано (размер отличался): {total_overwritten}\n"
            f"Не найдено файлов: {total_not_found}\n"
            f"Строк с ошибками: {len(errors)}"
        )
        if records_path is not None:
            summary += f"\n\nОтчёт: {records_path}"
        if summary_path is not None:
            summary += f"\nСводка: {summary_path}"

        self.status_var.set(
            f"Готово. Скопировано {total_copied}, уже было {total_already}, "
            f"не найдено {total_not_found}, ошибок {len(errors)}"
        )
        LOGGER.info(
            "Обработка завершена: строк %s, скопировано %s, уже было %s, перезаписано %s, "
            "не найдено %s, ошибок %s",
            len(results),
            total_copied,
            total_already,
            total_overwritten,
            total_not_found,
            len(errors),
        )
        messagebox.showinfo("Обработка завершена", summary, parent=self)

    def _on_job_error(self, message: str) -> None:
        """Сообщить об ошибке фоновой обработки."""
        self._set_buttons_enabled(True)
        self.stop_button.configure(state="disabled")
        self.status_var.set("Обработка прервана ошибкой")
        messagebox.showerror("Ошибка обработки", message, parent=self)

    # ------------------------------------------------------------------ #
    # Вкладка «Настройки»
    # ------------------------------------------------------------------ #
    def _update_azs_label(self) -> None:
        """Обновить подпись «Текущая АЗС»."""
        self.azs_var.set(f"Текущая АЗС: {self.azs.number}")

    def _rebuild_settings(self) -> None:
        """Перерисовать вкладку «Настройки» (read-only)."""
        general = self.config_data.general
        finder = self.config_data.finder
        text = self.settings_text
        text.configure(state="normal")
        text.delete("1.0", "end")

        def head(value: str) -> None:
            text.insert("end", value + "\n", "head")

        def line(value: str) -> None:
            text.insert("end", "  " + value + "\n")

        head("Общие настройки (settings.ini)")
        line(f"current_azs = {general.current_azs}")
        line(f"output_root = {general.output_root}")
        line(f"minutes_before = {general.minutes_before}")
        line(f"minutes_after = {general.minutes_after}")
        line(f"Интервал: -{general.minutes_before} мин до, +{general.minutes_after} мин после")
        line(f"copy_mode = {general.copy_mode}")
        line(f"threads = {general.threads}")
        line(f"video_roots = {general.video_roots_raw}")
        line(f"video_root_subdir = {general.video_root_subdir}")
        line(f"include_audio = {general.include_audio}")
        text.insert("end", "\n")

        head("Параметры поиска ([finder])")
        line(f"path_date_regex = {finder.path_date_regex}")
        line(f"path_date_filter = {finder.path_date_filter}")
        line(f"path_date_margin_hours = {finder.path_date_margin_hours}")
        line(f"time_source = {finder.time_source}")
        line(f"video_suffix_video = {finder.video_suffix_video}")
        line(f"video_suffix_audio = {finder.video_suffix_audio}")
        line(f"segment_duration_sec = {finder.segment_duration_sec}")
        line(f"mtime_fix_sec = {finder.mtime_fix_sec}")
        line(f"skip_dirs = {', '.join(finder.skip_dirs)}")
        text.insert("end", "\n")

        head(f"Текущая АЗС: {self.azs.number}")
        self._insert_path_line("output_root", Path(self.azs.output_root), is_dir=True)
        if self.azs.video_roots:
            line("video_roots (из azs.ini):")
            for root in self.azs.video_roots:
                self._insert_path_line("  корень", root, is_dir=True, indent="    ")
        else:
            line(f"video_roots = {general.video_roots_raw} (наследуется из settings.ini)")
        text.insert("end", "\n")

        head("Обнаруженные каталоги VIDEO")
        roots = drives.resolve_video_roots(self.config_data, self.azs)
        if not roots:
            text.insert("end", "  не найдено ни одного каталога VIDEO\n", "bad")
        for root in roots:
            ok = False
            try:
                ok = Path(root).is_dir()
            except OSError:
                ok = False
            text.insert("end", f"  {root} ", "ok" if ok else "bad")
            text.insert("end", "[OK]\n" if ok else "[НЕТ]\n")
        text.insert("end", "\n")

        head(f"Маппинг ТРК -> камеры (АЗС {self.azs.number})")
        for trk, cameras in sorted(self.azs.trk_map.items()):
            line(f"ТРК {trk}: {', '.join(cameras)}")
        text.insert("end", "\n")

        head("Файлы настроек")
        line(str(self.config_data.settings_path))
        line(str(self.config_data.azs_path))

        text.configure(state="disabled")

    def _insert_path_line(
        self, title: str, path: Path, is_dir: bool = True, indent: str = "  "
    ) -> None:
        """Вставить строку с путём и цветной пометкой его существования."""
        try:
            exists = path.is_dir() if is_dir else path.exists()
        except OSError:
            exists = False
        self.settings_text.insert("end", f"{indent}{title} = {path} ")
        self.settings_text.insert("end", "[OK]\n" if exists else "[НЕТ]\n", "ok" if exists else "bad")

    # ------------------------------------------------------------------ #
    # Служебное
    # ------------------------------------------------------------------ #
    def _open_path(self, path: Path) -> None:
        """Открыть файл или каталог средствами ОС."""
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except OSError as exc:
            messagebox.showerror("Ошибка", f"Не удалось открыть {path}: {exc}", parent=self)

    def _on_close(self) -> None:
        """Закрытие окна с проверкой активной задачи."""
        if self._is_busy():
            answer = messagebox.askyesno(
                "Идёт обработка",
                "Копирование ещё не завершено. Прервать и выйти?",
                parent=self,
            )
            if not answer:
                return
            self._stop_event.set()
        LOGGER.info("Приложение закрыто пользователем.")
        log_startup("startup: main window closed by user")
        self.root.destroy()


# --------------------------------------------------------------------------- #
# Запуск главного окна на уже созданном root
# --------------------------------------------------------------------------- #
def apply_azs(root: tk.Misc, azs: AzsConfig) -> None:
    """Загрузить в главное окно выбранную АЗС: output_root, video_roots, trk_N.

    Ничего не рисует и не опрашивает диски — только фиксирует выбор,
    который затем подхватит :func:`build_main_window`.

    :param root: главное окно, созданное в ``main.py``
    :param azs: выбранная секция ``[AZS_<номер>]``
    """
    root._fuelvideo_azs = azs  # type: ignore[attr-defined]
    drives.reset_cache()
    root.title(f"FuelVideo — АЗС {azs.number}")
    LOGGER.info(
        "Применена АЗС: %s; output_root=%s; video_roots=%s; ТРК: %s",
        azs.number,
        azs.output_root,
        ", ".join(str(item) for item in azs.video_roots) if azs.video_roots else "auto",
        ", ".join(str(trk) for trk in sorted(azs.trk_map)) or "нет",
    )
    log_startup(f"startup: azs applied {azs.number}")


def build_main_window(
    root: tk.Misc,
    config: AppConfig,
    ui_queue: "queue.Queue[str]",
) -> ExporterApp:
    """Собрать главное окно внутри уже созданного ``root`` (ровно один раз).

    :param root: главное окно, созданное в ``main.py``
    :param config: полные настройки
    :param ui_queue: очередь строк лога для вкладки «Лог»
    :return: экземпляр :class:`ExporterApp`
    """
    existing = getattr(root, "_fuelvideo_app", None)
    if isinstance(existing, ExporterApp):
        return existing
    azs = getattr(root, "_fuelvideo_azs", None)
    if azs is None:
        azs = config.azs_list[0]
    app = ExporterApp(root, config, azs, ui_queue)
    root._fuelvideo_app = app  # type: ignore[attr-defined]
    log_startup("startup: main window built")
    return app
