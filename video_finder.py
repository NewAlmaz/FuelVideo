"""Поиск видео- и аудиофайлов камер за заданный интервал времени.

Алгоритм (см. README, раздел «Как настроить под свою структуру VIDEO»):

1. Вычисляется «грубый» интервал по каталогам: ``start/end ± path_date_margin_hours``.
2. Перебираются каталоги первого уровня в ``video_root``; служебные
   (``META``, ``PROTECTED``, ``INDEX_DATA``) пропускаются.
3. Из имени каталога извлекаются дата и час по ``path_date_regex``; каталоги,
   не пересекающиеся с грубым интервалом, не открываются.
4. Файлы читаются без рекурсии; камера и тип определяются только по суффиксу
   (``._NN`` — видео, ``.wNN`` — звук).
5. Время файла берётся из ``mtime`` плюс суммарный сдвиг
   (``mtime_fix_sec`` + ``clock_drift_minutes`` * 60), и файл попадает в
   результат, если его сегмент пересекается с точным интервалом ``[start; end]``.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from config import CAMERA_PLACEHOLDER, FinderConfig

LOGGER = logging.getLogger(__name__)

KIND_VIDEO = "video"
KIND_AUDIO = "audio"

#: Размер «мягкого» окна вокруг интервала, если длительность сегмента неизвестна.
UNKNOWN_SEGMENT_MARGIN = timedelta(minutes=5)


@dataclass
class FoundFile:
    """Найденный файл камеры."""

    src: Path
    video_root: Path
    camera: str
    kind: str
    timestamp: Optional[datetime]
    size: int = 0

    @property
    def mtime_text(self) -> str:
        """Время файла в человекочитаемом виде."""
        if self.timestamp is None:
            return "время неизвестно"
        return self.timestamp.strftime("%d.%m.%Y %H:%M:%S")


@dataclass
class SearchStats:
    """Статистика обхода каталогов для строки лога."""

    dirs_visited: int = 0
    dirs_skipped_by_date: int = 0
    dirs_skipped_service: int = 0
    files_checked: int = 0
    files_matched: int = 0

    def merge(self, other: "SearchStats") -> None:
        """Прибавить статистику другого обхода."""
        self.dirs_visited += other.dirs_visited
        self.dirs_skipped_by_date += other.dirs_skipped_by_date
        self.dirs_skipped_service += other.dirs_skipped_service
        self.files_checked += other.files_checked
        self.files_matched += other.files_matched


@dataclass
class SearchResult:
    """Результат поиска по строке xlsx."""

    files: List[FoundFile] = field(default_factory=list)
    stats: SearchStats = field(default_factory=SearchStats)

    def by_camera(self) -> Dict[str, List[FoundFile]]:
        """Сгруппировать найденные файлы по номеру камеры."""
        grouped: Dict[str, List[FoundFile]] = {}
        for item in self.files:
            grouped.setdefault(item.camera, []).append(item)
        return grouped

    def count(self, kind: str) -> int:
        """Количество найденных файлов указанного типа."""
        return sum(1 for item in self.files if item.kind == kind)


def _overlaps(
    a_start: datetime, a_end: datetime, b_start: datetime, b_end: datetime
) -> bool:
    """Пересекаются ли два интервала (включая границы)."""
    return a_start <= b_end and b_start <= a_end


def parse_hour_dir(name: str, finder_cfg: FinderConfig) -> Optional[datetime]:
    """Извлечь дату/час из имени каталога вида ``12-10-26 10``.

    :param name: имя каталога первого уровня
    :param finder_cfg: параметры поиска
    :return: начало часа или ``None``, если имя не соответствует шаблону
    """
    pattern = finder_cfg.compiled_path_date_regex
    if pattern is None:
        return None
    match = pattern.search(name)
    if not match:
        return None

    groups = match.groupdict()
    try:
        day = int(groups["d"])
        month = int(groups["m"])
        year_text = groups["y"]
        hour = int(groups["H"])
    except (KeyError, TypeError, ValueError):
        LOGGER.warning("Не удалось разобрать дату из имени каталога «%s».", name)
        return None

    year = int(year_text)
    if year < 100:
        year += 2000

    try:
        return datetime(year, month, day, hour)
    except ValueError:
        LOGGER.warning(
            "Имя каталога «%s» похоже на дату, но дата некорректна (%s-%s-%s %s ч).",
            name,
            day,
            month,
            year,
            hour,
        )
        return None


def build_suffix_map(
    cameras: Sequence[str], finder_cfg: FinderConfig, include_audio: bool
) -> List[Tuple[str, str, str]]:
    """Построить список ``(суффикс, тип, камера)`` для сопоставления файлов.

    :param cameras: номера камер в формате ``07``
    :param finder_cfg: параметры поиска
    :param include_audio: учитывать ли аудиофайлы
    """
    mapping: List[Tuple[str, str, str]] = []
    for camera in cameras:
        video_suffix = finder_cfg.video_suffix_video.replace(CAMERA_PLACEHOLDER, camera)
        mapping.append((video_suffix.lower(), KIND_VIDEO, camera))
        if include_audio:
            audio_suffix = finder_cfg.video_suffix_audio.replace(CAMERA_PLACEHOLDER, camera)
            mapping.append((audio_suffix.lower(), KIND_AUDIO, camera))
    return mapping


def _file_timestamp(path: Path, finder_cfg: FinderConfig) -> Optional[datetime]:
    """Время файла: ``mtime`` или ``ctime`` плюс суммарный сдвиг.

    Суммарный сдвиг — это ``mtime_fix_sec`` (точная подстройка в секундах)
    плюс ``clock_drift_minutes`` * 60 (сдвиг часов видео-сервера).
    См. :attr:`config.FinderConfig.total_offset_sec`.
    """
    try:
        stat = path.stat()
    except OSError as exc:
        LOGGER.warning("Не удалось прочитать атрибуты файла %s: %s", path, exc)
        return None

    base = stat.st_mtime if finder_cfg.time_source == "mtime" else stat.st_ctime
    try:
        stamp = datetime.fromtimestamp(base)
    except (OverflowError, OSError, ValueError) as exc:
        LOGGER.warning("Не удалось преобразовать время файла %s: %s", path, exc)
        return None
    return stamp + timedelta(seconds=finder_cfg.total_offset_sec)


def _file_matches_interval(
    timestamp: Optional[datetime],
    start: datetime,
    end: datetime,
    finder_cfg: FinderConfig,
) -> bool:
    """Попадает ли файл в интервал (по началу сегмента и его длительности)."""
    if timestamp is None:
        # Время неизвестно — лучше скопировать лишний файл, чем пропустить нужный.
        return True

    if finder_cfg.segment_duration_sec > 0:
        segment_end = timestamp + timedelta(seconds=finder_cfg.segment_duration_sec)
        return _overlaps(timestamp, segment_end, start, end)

    # Длительность сегмента неизвестна: используем «мягкое» окно.
    return (start - UNKNOWN_SEGMENT_MARGIN) <= timestamp <= (end + UNKNOWN_SEGMENT_MARGIN)


def find_files(
    video_roots: Sequence[Path],
    cameras: Sequence[str],
    start: datetime,
    end: datetime,
    finder_cfg: FinderConfig,
    include_audio: bool = True,
    log_examples: bool = False,
    example_limit: int = 5,
) -> SearchResult:
    """Найти файлы указанных камер, попадающие в интервал ``[start; end]``.

    :param video_roots: корни видеоархива (``X:\\VIDEO``)
    :param cameras: номера камер в формате ``07``
    :param start: начало интервала
    :param end: конец интервала
    :param finder_cfg: параметры поиска
    :param include_audio: искать ли аудиофайлы
    :param log_examples: вывести в лог примеры найденных путей (для проверки настроек)
    :param example_limit: сколько примеров на камеру выводить
    """
    result = SearchResult()
    suffix_map = build_suffix_map(cameras, finder_cfg, include_audio)
    if not suffix_map:
        LOGGER.warning("Для поиска не задано ни одной камеры.")
        return result

    coarse_lo = start - timedelta(hours=finder_cfg.path_date_margin_hours)
    coarse_hi = end + timedelta(hours=finder_cfg.path_date_margin_hours)
    skip_upper = finder_cfg.skip_dirs_upper
    seen_paths: set = set()

    for raw_root in video_roots:
        root = Path(raw_root)
        try:
            if not root.is_dir():
                LOGGER.warning("Каталог видеоархива не найден: %s", root)
                continue
            entries = sorted(root.iterdir())
        except OSError as exc:
            LOGGER.warning("Не удалось прочитать каталог %s: %s", root, exc)
            continue

        hour_dirs = [entry for entry in entries if entry.is_dir()]
        if not hour_dirs:
            LOGGER.warning(
                "В каталоге %s нет ни одного час-каталога — похоже, путь указан неверно.",
                root,
            )
            continue

        for entry in hour_dirs:
            name = entry.name
            if name.upper() in skip_upper:
                result.stats.dirs_skipped_service += 1
                continue

            hour_start = parse_hour_dir(name, finder_cfg)
            if finder_cfg.path_date_filter and hour_start is not None:
                hour_end = hour_start + timedelta(hours=1)
                if not _overlaps(hour_start, hour_end, coarse_lo, coarse_hi):
                    result.stats.dirs_skipped_by_date += 1
                    continue

            result.stats.dirs_visited += 1
            try:
                files = sorted(entry.iterdir())
            except OSError as exc:
                LOGGER.warning("Не удалось прочитать каталог %s: %s", entry, exc)
                continue

            for file_entry in files:
                try:
                    if not file_entry.is_file():
                        continue
                except OSError:
                    continue

                result.stats.files_checked += 1
                name_lower = file_entry.name.lower()

                matched_kind = ""
                matched_camera = ""
                for suffix, kind, camera in suffix_map:
                    if name_lower.endswith(suffix):
                        matched_kind = kind
                        matched_camera = camera
                        break
                if not matched_kind:
                    continue

                key = str(file_entry).lower()
                if key in seen_paths:
                    continue

                timestamp = _file_timestamp(file_entry, finder_cfg)
                if timestamp is None:
                    LOGGER.warning(
                        "Время файла не определено, файл включён в результат: %s", file_entry
                    )
                elif not _file_matches_interval(timestamp, start, end, finder_cfg):
                    continue

                try:
                    size = file_entry.stat().st_size
                except OSError:
                    size = 0

                seen_paths.add(key)
                result.stats.files_matched += 1
                result.files.append(
                    FoundFile(
                        src=file_entry,
                        video_root=root,
                        camera=matched_camera,
                        kind=matched_kind,
                        timestamp=timestamp,
                        size=size,
                    )
                )

    if log_examples:
        log_search_examples(result, cameras, example_limit)

    return result


def log_search_examples(
    result: SearchResult, cameras: Sequence[str], limit: int = 5
) -> None:
    """Записать в лог до ``limit`` примеров найденных путей по каждой камере."""
    grouped = result.by_camera()
    for camera in cameras:
        files = grouped.get(camera, [])
        if not files:
            LOGGER.info("Камера %s: файлов не найдено.", camera)
            continue
        LOGGER.info("Камера %s: найдено файлов %s, примеры (до %s):", camera, len(files), limit)
        for item in files[:limit]:
            LOGGER.info("  %s  %s  [%s]", item.mtime_text, item.src, item.kind)
