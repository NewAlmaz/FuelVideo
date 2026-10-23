"""Чтение, валидация и хранение настроек приложения.

Модуль отвечает за:

* определение каталога, рядом с которым лежат ini-файлы и папка ``logs``;
* создание ``settings.ini`` и ``azs.ini`` из встроенных шаблонов при первом запуске;
* разбор и валидацию всех параметров (никаких «тихих» дефолтов в коде);
* настройку логирования (файл рядом с exe + необязательная очередь для GUI).

Все ошибки настроек поднимаются как :class:`ConfigError` с русским текстом.
"""

from __future__ import annotations

import configparser
import logging
import queue
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

LOGGER = logging.getLogger(__name__)

#: Имя секции АЗС: ``[AZS_75]``
AZS_SECTION_RE = re.compile(r"^AZS_(?P<num>\d+)$")
#: Ключ привязки ТРК: ``trk_1=07,08``
TRK_KEY_RE = re.compile(r"^trk_(?P<num>\d+)$")

VALID_COPY_MODES = ("copy", "move")
VALID_TIME_SOURCES = ("mtime", "ctime")
BOOL_TRUE = ("1", "true", "yes", "on", "да")
BOOL_FALSE = ("0", "false", "no", "off", "нет")

#: Имя параметра-подстановки номера камеры в настройках суффиксов.
CAMERA_PLACEHOLDER = "NN"


class ConfigError(Exception):
    """Ошибка в настройках. Текст исключения показывается пользователю."""


DEFAULT_SETTINGS_INI = r"""[general]
; Текущая АЗС:
;   auto    - выбрать при старте из секций [AZS_*] в azs.ini
;   <число> - работать с этой АЗС без диалога
current_azs=auto

; Куда копировать (если в [AZS_<номер>] не задан свой output_root)
output_root=D:\VideoExport

; Интервал вокруг docdate из xlsx:
;   start = docdate - minutes_before
;   end   = docdate + minutes_after
minutes_before=2
minutes_after=10

; copy - копировать, move - переносить (исходник удаляется после успеха)
copy_mode=copy
; Потоков копирования (1 = последовательно)
threads=2

; auto = искать X:\VIDEO\ на всех локальных дисках
; или явные пути через ; : D:\VIDEO;E:\VIDEO
video_roots=auto
; Имя подкаталога с видео на диске
video_root_subdir=VIDEO

; Копировать ли звук (.wNN) вместе с видео (._NN)
include_audio=true

[finder]
; Регулярное выражение для имени час-каталога (DD-MM-YY HH)
path_date_regex=(?P<d>\d{2})-(?P<m>\d{2})-(?P<y>\d{2})\s+(?P<H>\d{2})
; Включить фильтр по дате/часу из имени каталога
path_date_filter=true
; Запас по часам при отборе час-каталогов
path_date_margin_hours=1
; Откуда брать время файла: mtime или ctime
time_source=mtime
; Суффиксы: NN - номер камеры (01..99)
video_suffix_video=._NN
video_suffix_audio=.wNN
; Длительность одного файла-сегмента, сек (0 = неизвестна)
segment_duration_sec=60
; Поправка к времени файла, сек (может быть отрицательной)
mtime_fix_sec=0
; Сдвиг часов видео-сервера относительно реального времени, минут.
;   >0  — часы на видео-сервере ОТСТАЮТ (mtime меньше реального; прибавляем N минут)
;   <0  — часы СПЕШАТ (mtime больше реального; вычитаем N минут)
;   0   — без сдвига
; Применяется и к xlsx-режиму, и к ручному копированию.
; Для точной подстройки в секундах есть mtime_fix_sec.
clock_drift_minutes=0
; Служебные каталоги первого уровня, которые обходить нельзя
skip_dirs=META,PROTECTED,INDEX_DATA
"""

DEFAULT_AZS_INI = r"""; Секция = одна АЗС. Номера камер 01..99 через запятую.
; Одна ТРК может сниматься несколькими камерами: trk_1=07,08

[AZS_75]
; video_roots не задан -> используется [general].video_roots (auto)
output_root=D:\VideoExport\AZS75
trk_1=07,08
trk_2=09
trk_3=10,11
trk_4=12
trk_5=13
trk_6=14

; [AZS_76]
; video_roots=E:\VIDEO
; output_root=D:\VideoExport\AZS76
; trk_1=01,02
; trk_2=03
"""


# --------------------------------------------------------------------------- #
# Дат-классы настроек
# --------------------------------------------------------------------------- #
@dataclass
class GeneralConfig:
    """Секция ``[general]`` файла settings.ini."""

    current_azs: str
    output_root: Path
    minutes_before: int
    minutes_after: int
    copy_mode: str
    threads: int
    video_roots_raw: str
    video_root_subdir: str
    include_audio: bool

    @property
    def interval_text(self) -> str:
        """Человекочитаемое описание интервала, например «-2/+10 мин»."""
        return f"-{self.minutes_before}/+{self.minutes_after} мин"


@dataclass
class FinderConfig:
    """Секция ``[finder]`` файла settings.ini — параметры поиска файлов."""

    path_date_regex: str
    path_date_filter: bool
    path_date_margin_hours: float
    time_source: str
    video_suffix_video: str
    video_suffix_audio: str
    segment_duration_sec: float
    mtime_fix_sec: float
    clock_drift_minutes: float
    skip_dirs: List[str]
    compiled_path_date_regex: Optional["re.Pattern[str]"] = None

    @property
    def skip_dirs_upper(self) -> List[str]:
        """Служебные каталоги в верхнем регистре для быстрого сравнения."""
        return [name.upper() for name in self.skip_dirs]

    @property
    def total_offset_sec(self) -> float:
        """Суммарный сдвиг времени файла, сек.

        Учитывает и точную поправку ``mtime_fix_sec``, и сдвиг часов
        видео-сервера ``clock_drift_minutes`` (минуты переводятся в секунды).
        """
        return self.mtime_fix_sec + self.clock_drift_minutes * 60.0


@dataclass
class AzsConfig:
    """Секция ``[AZS_<номер>]`` файла azs.ini."""

    number: int
    output_root: Path
    video_roots: Optional[List[Path]] = None
    trk_map: Dict[int, List[str]] = field(default_factory=dict)
    output_root_explicit: bool = False

    @property
    def title(self) -> str:
        """Короткое описание для списков выбора, например «АЗС 75 — D:\\...»."""
        return f"АЗС {self.number} — {self.output_root}"

    def cameras_for_trk(self, trk: int) -> Optional[List[str]]:
        """Список камер для номера ТРК или ``None``, если настройки нет."""
        return self.trk_map.get(int(trk))


@dataclass
class AppConfig:
    """Все настройки приложения целиком."""

    base_dir: Path
    settings_path: Path
    azs_path: Path
    general: GeneralConfig
    finder: FinderConfig
    azs_list: List[AzsConfig] = field(default_factory=list)

    def get_azs(self, number: int) -> AzsConfig:
        """Вернуть настройки АЗС по номеру или поднять :class:`ConfigError`."""
        for azs in self.azs_list:
            if azs.number == int(number):
                return azs
        available = ", ".join(f"AZS_{a.number}" for a in self.azs_list) or "нет ни одной"
        raise ConfigError(
            f"В файле azs.ini нет секции [AZS_{number}]. Доступные секции: {available}."
        )

    def find_azs(self, number: int) -> Optional[AzsConfig]:
        """Вернуть настройки АЗС по номеру или ``None``."""
        for azs in self.azs_list:
            if azs.number == int(number):
                return azs
        return None

    @property
    def azs_numbers(self) -> List[int]:
        """Список номеров АЗС из azs.ini."""
        return [azs.number for azs in self.azs_list]


# --------------------------------------------------------------------------- #
# Пути и шаблоны ini
# --------------------------------------------------------------------------- #
def get_base_dir() -> Path:
    """Каталог, рядом с которым лежат ini-файлы и папка ``logs``."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def ensure_ini_files(base_dir: Path) -> Tuple[Path, Path]:
    """Создать ini-файлы рядом с exe, если их ещё нет.

    Источники в порядке приоритета:

    1. уже существующий внешний файл — не трогаем;
    2. файл, упакованный в exe (``--add-data``, распакован в ``sys._MEIPASS``);
    3. встроенная строка-шаблон (:data:`DEFAULT_SETTINGS_INI` / :data:`DEFAULT_AZS_INI`).

    :param base_dir: каталог рядом с exe (или с исходниками)
    :return: кортеж (путь к settings.ini, путь к azs.ini)
    """
    base_dir = Path(base_dir)
    settings_path = base_dir / "settings.ini"
    azs_path = base_dir / "azs.ini"

    def materialize(name: str, target: Path, fallback_text: str) -> None:
        if target.exists():
            # Внешний файл уже есть — не трогаем.
            LOGGER.info("Используется существующий файл: %s", target)
            return
        bundled = Path(getattr(sys, "_MEIPASS", base_dir)) / name
        try:
            bundled_exists = bundled.exists() and bundled.resolve() != target.resolve()
        except OSError:
            bundled_exists = False
        if bundled_exists:
            shutil.copy2(bundled, target)
            LOGGER.info("Создан %s из упакованного шаблона: %s", name, target)
            return
        target.write_text(fallback_text, encoding="utf-8")
        LOGGER.info("Создан %s из встроенной строки: %s", name, target)

    materialize("settings.ini", settings_path, DEFAULT_SETTINGS_INI)
    materialize("azs.ini", azs_path, DEFAULT_AZS_INI)
    return settings_path, azs_path


# --------------------------------------------------------------------------- #
# Вспомогательные чтения значений
# --------------------------------------------------------------------------- #
def _read_parser(path: Path) -> configparser.ConfigParser:
    """Прочитать ini-файл (UTF-8 с BOM или cp1251)."""
    parser = configparser.ConfigParser(interpolation=None)
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        try:
            text = path.read_text(encoding="cp1251")
        except OSError as exc:
            raise ConfigError(f"Не удалось прочитать файл {path.name}: {exc}") from exc
    except OSError as exc:
        raise ConfigError(f"Не удалось прочитать файл {path.name}: {exc}") from exc

    try:
        parser.read_string(text, source=str(path))
    except configparser.Error as exc:
        raise ConfigError(f"Ошибка разбора {path.name}: {exc}") from exc
    return parser


def _get(parser: configparser.ConfigParser, section: str, key: str, source: str) -> str:
    """Получить обязательный строковый параметр."""
    if not parser.has_section(section):
        raise ConfigError(f"В файле {source} отсутствует секция [{section}].")
    if not parser.has_option(section, key):
        raise ConfigError(
            f"В файле {source} в секции [{section}] отсутствует параметр «{key}»."
        )
    return parser.get(section, key).strip()


def _get_int(
    parser: configparser.ConfigParser,
    section: str,
    key: str,
    source: str,
    minimum: Optional[int] = None,
) -> int:
    """Получить обязательный целочисленный параметр."""
    raw = _get(parser, section, key, source)
    try:
        value = int(raw)
    except ValueError as exc:
        raise ConfigError(
            f"В файле {source} параметр [{section}].{key} должен быть целым числом, получено: «{raw}»."
        ) from exc
    if minimum is not None and value < minimum:
        raise ConfigError(
            f"В файле {source} параметр [{section}].{key} должен быть >= {minimum}, получено: {value}."
        )
    return value


def _get_float(
    parser: configparser.ConfigParser,
    section: str,
    key: str,
    source: str,
    minimum: Optional[float] = None,
) -> float:
    """Получить обязательный числовой параметр (допускается запятая)."""
    raw = _get(parser, section, key, source).replace(",", ".")
    try:
        value = float(raw)
    except ValueError as exc:
        raise ConfigError(
            f"В файле {source} параметр [{section}].{key} должен быть числом, получено: «{raw}»."
        ) from exc
    if minimum is not None and value < minimum:
        raise ConfigError(
            f"В файле {source} параметр [{section}].{key} должен быть >= {minimum}, получено: {value}."
        )
    return value


def _get_bool(parser: configparser.ConfigParser, section: str, key: str, source: str) -> bool:
    """Получить обязательный логический параметр (да/нет, true/false, 1/0)."""
    raw = _get(parser, section, key, source).lower()
    if raw in BOOL_TRUE:
        return True
    if raw in BOOL_FALSE:
        return False
    raise ConfigError(
        f"В файле {source} параметр [{section}].{key} должен быть true или false, получено: «{raw}»."
    )


def normalize_camera(value: str) -> str:
    """Привести номер камеры к виду ``07`` (01..99)."""
    text = str(value).strip()
    if not text.isdigit():
        raise ConfigError(f"Номер камеры должен быть числом от 1 до 99, получено: «{value}».")
    number = int(text)
    if not 1 <= number <= 99:
        raise ConfigError(f"Номер камеры должен быть в диапазоне 01..99, получено: {number}.")
    return f"{number:02d}"


def parse_camera_list(value: str, where: str) -> List[str]:
    """Разобрать список камер ``07,08`` -> ``['07', '08']``."""
    cameras: List[str] = []
    for chunk in value.replace(";", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        try:
            camera = normalize_camera(chunk)
        except ConfigError as exc:
            raise ConfigError(f"{where}: {exc}") from exc
        if camera not in cameras:
            cameras.append(camera)
    if not cameras:
        raise ConfigError(f"{where}: список камер пуст.")
    return cameras


def parse_roots(text: str) -> List[Path]:
    """Разобрать список путей ``D:\\VIDEO;E:\\VIDEO`` в список Path."""
    roots: List[Path] = []
    for chunk in text.replace("|", ";").split(";"):
        chunk = chunk.strip().strip('"')
        if not chunk:
            continue
        path = Path(chunk)
        if path not in roots:
            roots.append(path)
    return roots


# --------------------------------------------------------------------------- #
# Разбор секций
# --------------------------------------------------------------------------- #
def _load_general(parser: configparser.ConfigParser, source: str) -> GeneralConfig:
    """Разобрать и провалидировать секцию ``[general]``."""
    if not parser.has_section("general"):
        raise ConfigError(f"В файле {source} отсутствует секция [general].")

    current_azs = _get(parser, "general", "current_azs", source)
    output_root = Path(_get(parser, "general", "output_root", source))
    minutes_before = _get_int(parser, "general", "minutes_before", source, minimum=0)
    minutes_after = _get_int(parser, "general", "minutes_after", source, minimum=0)
    copy_mode = _get(parser, "general", "copy_mode", source).lower()
    threads = _get_int(parser, "general", "threads", source, minimum=1)
    video_roots_raw = _get(parser, "general", "video_roots", source)
    video_root_subdir = _get(parser, "general", "video_root_subdir", source)
    include_audio = _get_bool(parser, "general", "include_audio", source)

    if copy_mode not in VALID_COPY_MODES:
        raise ConfigError(
            f"В файле {source} параметр [general].copy_mode должен быть copy или move, "
            f"получено: «{copy_mode}»."
        )
    if minutes_after > 120:
        LOGGER.warning(
            "minutes_after=%s больше 120 минут — копироваться будет очень много файлов.",
            minutes_after,
        )
    if not video_roots_raw:
        raise ConfigError(f"В файле {source} параметр [general].video_roots пуст.")

    return GeneralConfig(
        current_azs=current_azs,
        output_root=output_root,
        minutes_before=minutes_before,
        minutes_after=minutes_after,
        copy_mode=copy_mode,
        threads=threads,
        video_roots_raw=video_roots_raw,
        video_root_subdir=video_root_subdir,
        include_audio=include_audio,
    )


def _load_finder(parser: configparser.ConfigParser, source: str) -> FinderConfig:
    """Разобрать и провалидировать секцию ``[finder]``."""
    if not parser.has_section("finder"):
        raise ConfigError(f"В файле {source} отсутствует секция [finder].")

    regex_text = _get(parser, "finder", "path_date_regex", source)
    try:
        compiled = re.compile(regex_text)
    except re.error as exc:
        raise ConfigError(
            f"В файле {source} параметр [finder].path_date_regex не компилируется: {exc}"
        ) from exc

    missing_groups = [name for name in ("d", "m", "y", "H") if name not in compiled.groupindex]
    if missing_groups:
        raise ConfigError(
            "Параметр [finder].path_date_regex должен содержать именованные группы "
            f"(?P<d>), (?P<m>), (?P<y>), (?P<H>). Не найдены: {', '.join(missing_groups)}."
        )

    path_date_filter = _get_bool(parser, "finder", "path_date_filter", source)
    margin_hours = _get_float(parser, "finder", "path_date_margin_hours", source, minimum=0.0)
    time_source = _get(parser, "finder", "time_source", source).lower()
    if time_source not in VALID_TIME_SOURCES:
        raise ConfigError(
            f"В файле {source} параметр [finder].time_source должен быть mtime или ctime, "
            f"получено: «{time_source}»."
        )

    suffix_video = _get(parser, "finder", "video_suffix_video", source)
    suffix_audio = _get(parser, "finder", "video_suffix_audio", source)
    for key, value in (("video_suffix_video", suffix_video), ("video_suffix_audio", suffix_audio)):
        if CAMERA_PLACEHOLDER not in value:
            raise ConfigError(
                f"В файле {source} параметр [finder].{key} должен содержать подстановку "
                f"номера камеры «{CAMERA_PLACEHOLDER}», получено: «{value}»."
            )

    segment_duration = _get_float(parser, "finder", "segment_duration_sec", source, minimum=0.0)
    mtime_fix = _get_float(parser, "finder", "mtime_fix_sec", source)
    # Сдвиг часов может быть отрицательным, поэтому minimum не задаём.
    clock_drift = _get_float(parser, "finder", "clock_drift_minutes", source)
    skip_dirs_raw = _get(parser, "finder", "skip_dirs", source)
    skip_dirs = [name.strip() for name in skip_dirs_raw.replace(";", ",").split(",") if name.strip()]

    return FinderConfig(
        path_date_regex=regex_text,
        path_date_filter=path_date_filter,
        path_date_margin_hours=margin_hours,
        time_source=time_source,
        video_suffix_video=suffix_video,
        video_suffix_audio=suffix_audio,
        segment_duration_sec=segment_duration,
        mtime_fix_sec=mtime_fix,
        clock_drift_minutes=clock_drift,
        skip_dirs=skip_dirs,
        compiled_path_date_regex=compiled,
    )


def _load_azs_list(
    parser: configparser.ConfigParser, source: str, general: GeneralConfig
) -> List[AzsConfig]:
    """Разобрать все секции ``[AZS_<номер>]`` файла azs.ini."""
    result: List[AzsConfig] = []
    seen: Dict[int, str] = {}

    for section in parser.sections():
        match = AZS_SECTION_RE.match(section)
        if not match:
            LOGGER.warning(
                "Секция [%s] в файле %s пропущена: имя должно иметь вид AZS_<номер>.",
                section,
                source,
            )
            continue

        number = int(match.group("num"))
        if number in seen:
            raise ConfigError(
                f"В файле {source} номер АЗС {number} встречается дважды: "
                f"секции [{seen[number]}] и [{section}]."
            )
        seen[number] = section

        video_roots: Optional[List[Path]] = None
        if parser.has_option(section, "video_roots"):
            raw_roots = parser.get(section, "video_roots").strip()
            if raw_roots:
                video_roots = parse_roots(raw_roots)
                if not video_roots:
                    raise ConfigError(
                        f"В файле {source} параметр [{section}].video_roots не содержит путей."
                    )

        explicit_root = False
        if parser.has_option(section, "output_root"):
            raw_output = parser.get(section, "output_root").strip()
            if raw_output:
                output_root = Path(raw_output)
                explicit_root = True
            else:
                output_root = general.output_root / f"AZS{number}"
        else:
            output_root = general.output_root / f"AZS{number}"

        trk_map: Dict[int, List[str]] = {}
        for key, value in parser.items(section):
            trk_match = TRK_KEY_RE.match(key.strip())
            if not trk_match:
                continue
            trk_number = int(trk_match.group("num"))
            trk_map[trk_number] = parse_camera_list(value, f"[{section}].{key}")

        if not trk_map:
            raise ConfigError(
                f"В файле {source} в секции [{section}] нет ни одной привязки trk_N=<камеры>. "
                "Пример: trk_1=07,08"
            )

        result.append(
            AzsConfig(
                number=number,
                output_root=output_root,
                video_roots=video_roots,
                trk_map=dict(sorted(trk_map.items())),
                output_root_explicit=explicit_root,
            )
        )

    if not result:
        raise ConfigError(
            f"В файле {source} нет ни одной секции вида [AZS_<номер>]. "
            "Добавьте секцию своей АЗС, например [AZS_75]."
        )

    result.sort(key=lambda item: item.number)
    return result


def parse_current_azs(value: str) -> Optional[int]:
    """Разобрать ``[general].current_azs``: ``None`` для auto, иначе номер АЗС."""
    text = str(value).strip().lower()
    if text == "auto":
        return None
    if re.fullmatch(r"\d+", text):
        return int(text)
    raise ConfigError(
        "Параметр [general].current_azs должен быть «auto» или номером АЗС "
        f"(например 75), получено: «{value}»."
    )


def load_config(base_dir: Optional[Path] = None, create_files: bool = True) -> AppConfig:
    """Прочитать и провалидировать оба ini-файла.

    :param base_dir: каталог рядом с exe; по умолчанию вычисляется автоматически
    :param create_files: создавать ini-файлы из шаблонов при отсутствии
    :raises ConfigError: при любой проблеме настроек
    """
    base_dir = Path(base_dir) if base_dir is not None else get_base_dir()
    if create_files:
        settings_path, azs_path = ensure_ini_files(base_dir)
    else:
        settings_path = base_dir / "settings.ini"
        azs_path = base_dir / "azs.ini"
        if not settings_path.exists() or not azs_path.exists():
            raise ConfigError(
                f"Не найдены файлы settings.ini и/или azs.ini в каталоге {base_dir}."
            )

    settings_parser = _read_parser(settings_path)
    azs_parser = _read_parser(azs_path)

    general = _load_general(settings_parser, settings_path.name)
    finder = _load_finder(settings_parser, settings_path.name)
    azs_list = _load_azs_list(azs_parser, azs_path.name, general)

    config = AppConfig(
        base_dir=base_dir,
        settings_path=settings_path,
        azs_path=azs_path,
        general=general,
        finder=finder,
        azs_list=azs_list,
    )

    # Проверяем, что current_azs валиден и указывает на существующую секцию.
    current = parse_current_azs(general.current_azs)
    if current is not None:
        config.get_azs(current)

    return config


# --------------------------------------------------------------------------- #
# Логирование
# --------------------------------------------------------------------------- #
class QueueLogHandler(logging.Handler):
    """Обработчик, складывающий готовые строки лога в очередь (для GUI)."""

    def __init__(self, ui_queue: "queue.Queue[str]") -> None:
        super().__init__()
        self.ui_queue = ui_queue

    def emit(self, record: logging.LogRecord) -> None:
        """Положить отформатированную запись в очередь."""
        try:
            self.ui_queue.put_nowait(self.format(record))
        except Exception:  # pragma: no cover - защита от падения логирования
            self.handleError(record)


def setup_logging(
    base_dir: Path,
    ui_queue: Optional["queue.Queue[str]"] = None,
    level: int = logging.INFO,
) -> Path:
    """Настроить логирование в файл ``logs/app_YYYYMMDD.log`` и (опционально) в GUI.

    :param base_dir: каталог рядом с exe
    :param ui_queue: очередь, куда дублируются строки лога для вкладки «Лог»
    :param level: уровень логирования
    :return: путь к файлу лога
    """
    logs_dir = Path(base_dir) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / f"app_{datetime.now():%Y%m%d}.log"

    root = logging.getLogger()
    root.setLevel(level)
    for handler in list(root.handlers):
        root.removeHandler(handler)

    file_formatter = logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s: %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(file_formatter)
    root.addHandler(file_handler)

    if sys.stdout is not None and not getattr(sys, "frozen", False):
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(file_formatter)
        root.addHandler(stream_handler)

    if ui_queue is not None:
        ui_handler = QueueLogHandler(ui_queue)
        ui_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", "%H:%M:%S"))
        root.addHandler(ui_handler)

    logging.getLogger("openpyxl").setLevel(logging.WARNING)
    return log_path
