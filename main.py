"""Точка входа приложения FuelVideo.

Режимы запуска::

    python main.py                 # обычный запуск GUI
    python main.py --selftest      # самотестирование без копирования
    python main.py --selftest --xlsx "C:\\path\\to\\report.xlsx"
    python main.py --xlsx report.xlsx   # GUI сразу с загруженным файлом

В режиме ``--selftest`` программа проверяет настройки, показывает найденные
каталоги ``X:\\VIDEO`` и по каждой строке xlsx печатает список камер и
количество найденных файлов, ничего не копируя.
"""

from __future__ import annotations

import argparse
import logging
import queue
import sys
from pathlib import Path
from typing import List, Optional, Sequence

import drives
import expiry
from config import (
    AppConfig,
    ConfigError,
    get_base_dir,
    load_config,
    parse_current_azs,
    setup_logging,
)
from copier import STATUS_ERROR, STATUS_SKIPPED_OTHER_AZS, JobRunner, status_label
from xlsx_reader import RowData, XlsxError, read_rows

LOGGER = logging.getLogger("main")


# --------------------------------------------------------------------------- #
# Разбор аргументов командной строки
# --------------------------------------------------------------------------- #
def build_arg_parser() -> argparse.ArgumentParser:
    """Создать парсер аргументов командной строки."""
    parser = argparse.ArgumentParser(
        prog="FuelVideo",
        description="Копирование видео и звука из СВН «Интеллект» по данным из Excel.",
    )
    parser.add_argument(
        "--selftest",
        action="store_true",
        help="выполнить самотестирование (поиск файлов без копирования) и выйти",
    )
    parser.add_argument(
        "--xlsx",
        metavar="FILE",
        default=None,
        help="путь к файлу xlsx со списком заправок",
    )
    parser.add_argument(
        "--config-dir",
        metavar="DIR",
        default=None,
        help="каталог, где лежат settings.ini и azs.ini (по умолчанию рядом с exe)",
    )
    return parser


def find_default_xlsx(base_dir: Path) -> Optional[Path]:
    """Найти первый xlsx рядом с программой или в текущем каталоге."""
    candidates: List[Path] = []
    for folder in (base_dir, Path.cwd()):
        try:
            candidates.extend(sorted(folder.glob("*.xlsx")))
        except OSError:
            continue
    for candidate in candidates:
        if not candidate.name.startswith("~$"):
            return candidate
    return None


# --------------------------------------------------------------------------- #
# Самотестирование
# --------------------------------------------------------------------------- #
def print_config_summary(config: AppConfig) -> None:
    """Напечатать сводку по настройкам."""
    general = config.general
    finder = config.finder
    print("=== Настройки ===")
    print(f"Каталог программы : {config.base_dir}")
    print(f"settings.ini      : {config.settings_path}")
    print(f"azs.ini           : {config.azs_path}")
    print(f"current_azs       : {general.current_azs}")
    print(f"output_root общий : {general.output_root}")
    print(f"Интервал          : -{general.minutes_before} / +{general.minutes_after} мин")
    print(f"copy_mode/threads : {general.copy_mode} / {general.threads}")
    print(f"include_audio     : {general.include_audio}")
    print(f"video_roots       : {general.video_roots_raw}")
    print(f"video_root_subdir : {general.video_root_subdir}")
    print(
        "Поиск             : time_source={}, segment={} c, margin={} ч, mtime_fix={} c".format(
            finder.time_source,
            finder.segment_duration_sec,
            finder.path_date_margin_hours,
            finder.mtime_fix_sec,
        )
    )
    drift = float(getattr(finder, "clock_drift_minutes", 0.0) or 0.0)
    if abs(drift) < 1e-9:
        drift_text = "без сдвига"
    elif drift > 0:
        drift_text = f"часы видео-сервера отстают на {drift:g} мин (к mtime прибавляется сдвиг)"
    else:
        drift_text = (
            f"часы видео-сервера спешат на {abs(drift):g} мин (из mtime вычитается сдвиг)"
        )
    print(f"clock_drift       : {drift:g} мин ({drift_text})")
    print(f"skip_dirs         : {', '.join(finder.skip_dirs)}")
    print()


def print_azs_summary(config: AppConfig) -> None:
    """Напечатать секции АЗС и реальные корни видеоархива."""
    print("=== Секции АЗС ===")
    for azs in config.azs_list:
        roots = drives.resolve_video_roots(config, azs)
        roots_text = ", ".join(str(item) for item in roots) or "нет"
        print(f"AZS_{azs.number}: output_root={azs.output_root}")
        print(f"    video_roots: {roots_text}")
        for root in roots:
            state = "есть" if Path(root).is_dir() else "НЕТ"
            print(f"      {root} — {state}")
        for trk, cameras in sorted(azs.trk_map.items()):
            print(f"    ТРК {trk}: {', '.join(cameras)}")
    print()


def run_selftest(config: AppConfig, xlsx_path: Optional[Path]) -> int:
    """Выполнить самотестирование и вернуть код возврата.

    :param config: прочитанные настройки
    :param xlsx_path: путь к xlsx (может быть ``None``)
    :return: 0, если хотя бы одна строка обработана без ошибок, иначе 1
    """
    print("=== FuelVideo: самотестирование ===")
    print()
    rollback, saved_date = expiry.check_clock_rollback()
    print(f"Срок действия: {expiry.expiry_text()}", end="")
    if expiry.is_expired():
        print(" (ИСТЁК)")
    else:
        print()
    if rollback:
        if saved_date is not None:
            print(f"Обнаружен откат часов: до отката была дата {saved_date:%d.%m.%Y}")
        else:
            print("Обнаружен откат часов.")
    print()

    parsed_current = parse_current_azs(config.general.current_azs)
    print_config_summary(config)

    print("=== Обнаружение дисков ===")
    discovered = drives.discover_video_roots(subdir=config.general.video_root_subdir)
    if discovered:
        print("Найдены каталоги VIDEO:")
        for root in discovered:
            print(f"  {root}")
    else:
        print(
            f"Каталоги {config.general.video_root_subdir} не найдены ни на одном диске "
            "(это предупреждение, не ошибка)."
        )
    print()

    print_azs_summary(config)

    if parsed_current is not None:
        print(f"Режим current_azs: работа без диалога с АЗС {parsed_current}.")
    elif len(config.azs_list) == 1:
        print(f"Режим current_azs=auto: в azs.ini одна секция — АЗС {config.azs_list[0].number}.")
    else:
        print("Режим current_azs=auto: секций несколько — при запуске GUI будет диалог выбора.")
    print()

    if xlsx_path is None:
        print("Файл xlsx не указан (--xlsx) и не найден рядом с программой.")
        print("Проверка строк невозможна: код возврата 1.")
        return 1

    try:
        rows = read_rows(xlsx_path)
    except XlsxError as exc:
        print(f"Ошибка чтения xlsx: {exc}")
        return 1

    print(f"=== Строки xlsx ({xlsx_path.name}): {len(rows)} ===")
    if not rows:
        print("В файле нет строк с данными.")
        return 1

    known_numbers = set(config.azs_numbers)
    success_rows = 0
    processed_rows = 0

    for azs in config.azs_list:
        azs_rows = [row for row in rows if row.azs == azs.number]
        if not azs_rows:
            continue
        print(f"--- АЗС {azs.number}: строк {len(azs_rows)} ---")
        runner = JobRunner(config, azs)
        results = runner.run(azs_rows, scan_only=True)
        for result in results:
            row = result.row
            processed_rows += 1
            if result.status not in (STATUS_ERROR, STATUS_SKIPPED_OTHER_AZS):
                success_rows += 1
            print(
                "Строка {index}: АЗС={azs} ТРК={trk} docdate={docdate} "
                "интервал={start}..{end} камеры=[{cameras}] видео={video} аудио={audio} — {status}".format(
                    index=row.index,
                    azs=row.azs,
                    trk=row.trk,
                    docdate=row.docdate_text,
                    start=result.start_text,
                    end=result.end_text,
                    cameras=result.cameras_text or "—",
                    video=result.found_video,
                    audio=result.found_audio,
                    status=status_label(result.status),
                )
            )
            if result.error:
                print(f"    Причина: {result.error}")
        print()

    unknown_rows = [
        row for row in rows if row.azs is not None and row.azs not in known_numbers
    ]
    if unknown_rows:
        numbers = sorted({row.azs for row in unknown_rows})
        print(
            "Строки с АЗС без секции в azs.ini: "
            f"{', '.join(str(number) for number in numbers if number is not None)} "
            f"(строк {len(unknown_rows)})."
        )
        for row in unknown_rows[:5]:
            print(f"    строка {row.index}: azscode={row.azs}")
        if len(unknown_rows) > 5:
            print(f"    ... и ещё {len(unknown_rows) - 5} строк")

    print()
    print(f"Обработано строк: {processed_rows}, без ошибок: {success_rows}")
    if success_rows:
        print("Код возврата 0.")
        return 0
    print("Ни одна строка не обработана без ошибок. Код возврата 1.")
    return 1


# --------------------------------------------------------------------------- #
# Запуск GUI
# --------------------------------------------------------------------------- #
def show_error_dialog(title: str, message: str) -> None:
    """Показать модальное окно ошибки (если доступен tkinter)."""
    print(f"{title}: {message}", file=sys.stderr)
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:  # pragma: no cover - окружение без tkinter
        return
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:  # pragma: no cover
        pass


def resolve_azs_or_none(root: "tk.Misc", config: AppConfig) -> Optional[int]:
    """Определить АЗС для работы.

    Правила:

    * ``current_azs=<число>`` — секция должна существовать, диалог не показывается;
    * ``current_azs=auto`` и одна секция — возвращается её номер;
    * ``current_azs=auto`` и ни одной секции — сообщение об ошибке и ``None``;
    * ``current_azs=auto`` и несколько секций — модальный диалог выбора.

    :param root: уже созданное (скрытое) главное окно — родитель для диалога
    :param config: прочитанные настройки
    :return: номер АЗС или ``None``, если работать не с чем
    """
    import gui  # импорт внутри функции: --selftest работает и без дисплея

    try:
        try:
            current = parse_current_azs(config.general.current_azs)
        except ConfigError as exc:
            gui.show_error(root, "Ошибка настроек", str(exc))
            return None

        if current is not None:
            if config.find_azs(current) is None:
                available = (
                    ", ".join(f"AZS_{number}" for number in config.azs_numbers) or "нет ни одной"
                )
                gui.show_error(
                    root,
                    "Секция АЗС не найдена",
                    f"В azs.ini нет секции [AZS_{current}] (current_azs={current}).\n"
                    f"Доступные секции: {available}.",
                )
                return None
            LOGGER.info("current_azs=%s: диалог выбора АЗС не требуется.", current)
            gui.log_startup(f"startup: current_azs={current}, dialog skipped")
            return current

        sections = list(config.azs_list)
        if len(sections) == 1:
            only = sections[0].number
            LOGGER.info("current_azs=auto: в azs.ini одна секция — АЗС %s.", only)
            gui.log_startup(f"startup: auto with a single section -> AZS {only}")
            return only

        if not sections:
            gui.show_error(
                root,
                "Нет секций АЗС",
                "В azs.ini нет ни одной секции вида [AZS_<номер>].\n"
                "Добавьте секцию своей АЗС и запустите программу снова.",
            )
            return None

        options = [(azs.number, gui.azs_option_label(azs)) for azs in sections]
        chosen = gui.ask_azs_dialog(root, options)
        if chosen is None:
            LOGGER.info("Диалог выбора АЗС закрыт без выбора, выход.")
            gui.log_startup("startup: AZS dialog cancelled, exit")
            return None
        gui.log_startup(f"startup: AZS dialog result {chosen}")
        return chosen
    except Exception as exc:  # любая неожиданная ошибка не должна вешать процесс
        LOGGER.exception("Ошибка выбора АЗС: %s", exc)
        gui.log_startup_exception("startup: exception in resolve_azs_or_none")
        gui.show_error(
            root,
            "Ошибка выбора АЗС",
            f"{exc}\n\nПодробности записаны в startup_debug.log.",
        )
        return None


def run_gui(config: AppConfig, ui_queue: "queue.Queue[str]", xlsx_path: Optional[Path]) -> int:
    """Запустить графический интерфейс.

    Порядок строго фиксирован:

    1. создаётся ``root`` и сразу прячется;
    2. выбирается АЗС (диалог / auto / номер из настроек);
    3. к ``root`` применяется выбранная АЗС и собирается главное окно;
    4. окно показывается и запускается единственный ``mainloop()``.

    :param config: настройки
    :param ui_queue: очередь для строк лога
    :param xlsx_path: заранее выбранный файл xlsx (может быть ``None``)
    """
    import tkinter as tk

    import gui  # импорт внутри функции: --selftest работает и без дисплея

    root = tk.Tk()
    root.withdraw()  # прячем ГЛАВНОЕ окно, пока выбираем АЗС
    clock_rollback, saved_date = expiry.check_clock_rollback()
    expired = expiry.is_expired()
    log_expiry = LOGGER.warning if expired else LOGGER.info
    log_expiry(
        "expiry: %s, expired=%s, clock_rollback=%s (saved=%s)",
        expiry.expiry_text(),
        expired,
        clock_rollback,
        saved_date,
    )
    log_path = gui._startup_debug_path()
    gui.log_startup(
        f"startup: app launch, current_azs={config.general.current_azs!r}, "
        f"sections={config.azs_numbers}, debug_log={log_path}"
    )

    azs_number = resolve_azs_or_none(root, config)  # см. п.2
    if azs_number is None:
        LOGGER.info("АЗС не выбрана, выход.")
        gui.log_startup("startup: no AZS selected, exit")
        root.destroy()
        return 0

    try:
        azs = config.get_azs(azs_number)
    except ConfigError as exc:
        gui.show_error(root, "Ошибка настроек", str(exc))
        root.destroy()
        return 2

    gui.apply_azs(root, azs)  # загрузить video_roots/output_root/trk_N
    app = gui.build_main_window(root, config, ui_queue)  # собрать главное окно

    if xlsx_path is not None:
        app.load_xlsx(xlsx_path)

    def log_resolution() -> None:
        """Записать в лог итоговые корни видеоархива (может опрашивать диски)."""
        try:
            roots = drives.resolve_video_roots(config, app.azs)
            roots_text = ", ".join(str(item) for item in roots) or "нет"
            LOGGER.info(
                "Старт. current_azs=%s; выбрана АЗС: %s; video_roots=%s; output_root=%s; интервал: %s",
                config.general.current_azs,
                app.azs.number,
                roots_text,
                app.azs.output_root,
                config.general.interval_text,
            )
            gui.log_startup(f"startup: video_roots for AZS {app.azs.number}: {roots_text}")
        except Exception:
            LOGGER.exception("Не удалось определить корни видеоархива.")
            gui.log_startup_exception("startup: failed to resolve video_roots")

    root.deiconify()  # ПОКАЗАТЬ главное окно
    root.after(100, log_resolution)
    gui.log_startup("startup: main window shown, entering mainloop")
    root.mainloop()  # единственный mainloop в программе
    return 0


# --------------------------------------------------------------------------- #
# Точка входа
# --------------------------------------------------------------------------- #
def main(argv: Optional[Sequence[str]] = None) -> int:
    """Главная функция приложения."""
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    base_dir = Path(args.config_dir).resolve() if args.config_dir else get_base_dir()
    ui_queue: "queue.Queue[str]" = queue.Queue()
    log_path = setup_logging(base_dir, ui_queue=None if args.selftest else ui_queue)
    LOGGER.info("Лог файла: %s", log_path)

    try:
        config = load_config(base_dir)
    except ConfigError as exc:
        LOGGER.error("Ошибка настроек: %s", exc)
        if args.selftest:
            print(f"Ошибка настроек: {exc}", file=sys.stderr)
        else:
            show_error_dialog("Ошибка настроек", str(exc))
        return 2

    xlsx_path: Optional[Path] = None
    if args.xlsx:
        xlsx_path = Path(args.xlsx).expanduser()
        if not xlsx_path.is_absolute():
            xlsx_path = (Path.cwd() / xlsx_path).resolve()
        if not xlsx_path.exists() and args.selftest:
            print(f"Файл xlsx не найден: {xlsx_path}", file=sys.stderr)
            return 1
    elif args.selftest:
        xlsx_path = find_default_xlsx(base_dir)
        if xlsx_path is not None:
            print(f"Автоматически найден файл xlsx: {xlsx_path}")

    if args.selftest:
        return run_selftest(config, xlsx_path)

    return run_gui(config, ui_queue, xlsx_path)


if __name__ == "__main__":
    sys.exit(main())
