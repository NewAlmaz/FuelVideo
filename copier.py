"""Копирование найденных файлов, отчёт CSV и обработка строк xlsx.

Принцип копирования — «зеркало источника»: файл сохраняет относительный путь
относительно ``video_root``::

    g:\\VIDEO\\12-10-26 10\\A._07  ->  D:\\VideoExport\\AZS75\\12-10-26 10\\A._07

Никаких подстановок ``{azs}``/``{date}``/``{cardcode}`` в пути нет.

Повторный запуск по тому же xlsx безопасен: существующие файлы того же размера
пропускаются (``already_exists``), перезаписываются только при расхождении
размеров (``size_mismatch_overwritten``).
"""

from __future__ import annotations

import csv
import logging
import shutil
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import drives
from config import AppConfig, AzsConfig, ConfigError, FinderConfig, GeneralConfig
from video_finder import FoundFile, SearchResult, SearchStats, find_files
from xlsx_reader import RowData

LOGGER = logging.getLogger(__name__)

STATUS_PENDING = "pending"
STATUS_FOUND = "found"
STATUS_COPIED = "copied"
STATUS_ALREADY_EXISTS = "already_exists"
STATUS_SIZE_MISMATCH = "size_mismatch_overwritten"
STATUS_NOT_FOUND = "not_found"
STATUS_ERROR = "error"
STATUS_SKIPPED_OTHER_AZS = "skipped_other_azs"
STATUS_STOPPED = "stopped"

#: Русские подписи статусов для GUI и отчётов.
STATUS_LABELS: Dict[str, str] = {
    STATUS_PENDING: "ожидание",
    STATUS_FOUND: "найдено (без копирования)",
    STATUS_COPIED: "скопировано",
    STATUS_ALREADY_EXISTS: "уже было",
    STATUS_SIZE_MISMATCH: "перезаписано (размер отличался)",
    STATUS_NOT_FOUND: "не найдено",
    STATUS_ERROR: "ошибка",
    STATUS_SKIPPED_OTHER_AZS: "пропущено (другая АЗС)",
    STATUS_STOPPED: "остановлено",
}


def status_label(status: str) -> str:
    """Русская подпись статуса."""
    return STATUS_LABELS.get(status, status)


@dataclass
class CopyRecord:
    """Одна запись отчёта: файл (или отсутствующий файл) по конкретной камере."""

    row_index: int
    azs: Optional[int]
    docdate: Optional[datetime]
    trk: Optional[int]
    camera: str
    kind: str
    src: Optional[Path]
    dst: Optional[Path]
    status: str
    error: str = ""

    def as_csv_row(self) -> List[str]:
        """Строка для CSV-отчёта."""
        return [
            "" if self.azs is None else str(self.azs),
            "" if self.docdate is None else self.docdate.strftime("%d.%m.%Y %H:%M:%S"),
            "" if self.trk is None else str(self.trk),
            self.camera,
            self.kind,
            "" if self.src is None else str(self.src),
            "" if self.dst is None else str(self.dst),
            self.status,
            self.error,
        ]


@dataclass
class RowResult:
    """Результат обработки одной строки xlsx."""

    row: RowData
    cameras: List[str] = field(default_factory=list)
    start: Optional[datetime] = None
    end: Optional[datetime] = None
    files: List[FoundFile] = field(default_factory=list)
    records: List[CopyRecord] = field(default_factory=list)
    status: str = STATUS_PENDING
    error: str = ""
    found_video: int = 0
    found_audio: int = 0
    copied: int = 0
    already_exists: int = 0
    overwritten: int = 0
    not_found: int = 0

    @property
    def status_text(self) -> str:
        """Русская подпись статуса строки."""
        return status_label(self.status)

    @property
    def cameras_text(self) -> str:
        """Камеры через запятую."""
        return ",".join(self.cameras)

    @property
    def start_text(self) -> str:
        """Начало интервала в человекочитаемом виде."""
        return self.start.strftime("%d.%m.%Y %H:%M:%S") if self.start else "—"

    @property
    def end_text(self) -> str:
        """Конец интервала в человекочитаемом виде."""
        return self.end.strftime("%d.%m.%Y %H:%M:%S") if self.end else "—"

    @property
    def summary_text(self) -> str:
        """Краткая сводка по строке для лога и вкладки «Задания»."""
        return (
            f"найдено {self.found_video + self.found_audio}"
            f" (видео {self.found_video}, аудио {self.found_audio}), "
            f"скопировано {self.copied}, уже было {self.already_exists}, "
            f"не найдено {self.not_found}"
        )


def destination_for(src: Path, video_root: Path, output_root: Path) -> Path:
    """Куда будет скопирован файл: ``output_root`` + путь относительно ``video_root``."""
    try:
        relative = Path(src).relative_to(Path(video_root))
    except ValueError:
        LOGGER.warning(
            "Файл %s не находится внутри %s, структура каталогов будет потеряна.", src, video_root
        )
        relative = Path(Path(src).name)
    return Path(output_root) / relative


def copy_one(
    src: Path,
    video_root: Path,
    output_root: Path,
    mode: str = "copy",
) -> Tuple[str, Path, str]:
    """Скопировать (или перенести) один файл с сохранением структуры.

    :param src: исходный файл
    :param video_root: корень видеоархива (для вычисления относительного пути)
    :param output_root: каталог результата
    :param mode: ``copy`` или ``move``
    :return: кортеж (статус, путь назначения, текст ошибки)
    """
    src = Path(src)
    dst = destination_for(src, video_root, output_root)

    try:
        if src.resolve() == dst.resolve():
            return (
                STATUS_ERROR,
                dst,
                "Путь назначения совпадает с источником — проверьте output_root.",
            )
    except OSError:
        pass

    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return STATUS_ERROR, dst, f"Не удалось создать каталог {dst.parent}: {exc}"

    status = STATUS_COPIED
    try:
        if dst.exists():
            try:
                dst_size = dst.stat().st_size
                src_size = src.stat().st_size
            except OSError as exc:
                return STATUS_ERROR, dst, f"Не удалось прочитать размер файла: {exc}"

            if dst_size == src_size:
                LOGGER.debug("Файл уже скопирован ранее: %s", dst)
                return STATUS_ALREADY_EXISTS, dst, ""

            LOGGER.warning(
                "Размеры файлов отличаются, файл будет перезаписан: %s (%s != %s байт)",
                dst,
                dst_size,
                src_size,
            )
            shutil.copy2(src, dst)
            status = STATUS_SIZE_MISMATCH
        else:
            shutil.copy2(src, dst)
    except OSError as exc:
        return STATUS_ERROR, dst, f"Ошибка копирования {src} -> {dst}: {exc}"

    if mode == "move" and status in (STATUS_COPIED, STATUS_SIZE_MISMATCH):
        try:
            src.unlink()
        except OSError as exc:
            LOGGER.warning("Не удалось удалить исходный файл %s: %s", src, exc)
            return status, dst, f"Скопировано, но исходник не удалён: {exc}"

    return status, dst, ""


# --------------------------------------------------------------------------- #
# Общие операции для xlsx-режима и ручного копирования
# --------------------------------------------------------------------------- #
def search_files_for_cameras(
    azs: AzsConfig,
    cameras: Sequence[str],
    start: datetime,
    end: datetime,
    finder_cfg: FinderConfig,
    video_roots: Sequence[Path],
    include_audio: bool = True,
    log_examples: bool = False,
) -> SearchResult:
    """Найти файлы камер за точный интервал ``[start; end]`` (сo статистикой).

    Общая функция для вкладки «Задания» (xlsx) и вкладки «Ручное копирование».
    Никаких ``minutes_before``/``minutes_after`` здесь нет: интервал передаётся
    готовым. Сдвиг часов видео-сервера берётся из ``finder_cfg``
    (``mtime_fix_sec`` + ``clock_drift_minutes`` * 60).

    :param azs: текущая АЗС (для лога)
    :param cameras: номера камер в формате ``07``
    :param start: начало интервала
    :param end: конец интервала
    :param finder_cfg: параметры поиска
    :param video_roots: корни видеоархива
    :param include_audio: искать ли аудиофайлы
    :param log_examples: вывести в лог примеры найденных путей
    :return: :class:`video_finder.SearchResult` (файлы + статистика обхода)
    """
    if not cameras or not video_roots:
        return SearchResult()

    LOGGER.debug(
        "Поиск: АЗС %s, камеры %s, интервал %s..%s, сдвиг %.1f c",
        azs.number,
        ",".join(cameras),
        start.strftime("%d.%m.%Y %H:%M:%S"),
        end.strftime("%d.%m.%Y %H:%M:%S"),
        finder_cfg.total_offset_sec,
    )

    return find_files(
        video_roots=video_roots,
        cameras=cameras,
        start=start,
        end=end,
        finder_cfg=finder_cfg,
        include_audio=include_audio,
        log_examples=log_examples,
    )


def find_files_for_cameras(
    azs: AzsConfig,
    cameras: Sequence[str],
    start: datetime,
    end: datetime,
    finder_cfg: FinderConfig,
    video_roots: Sequence[Path],
    include_audio: bool = True,
    log_examples: bool = False,
) -> List[FoundFile]:
    """Найти файлы камер за точный интервал ``[start; end]``.

    Обёртка над :func:`search_files_for_cameras` для случаев, когда нужен
    только список файлов (ручное копирование).

    :return: список найденных файлов
    """
    return search_files_for_cameras(
        azs=azs,
        cameras=cameras,
        start=start,
        end=end,
        finder_cfg=finder_cfg,
        video_roots=video_roots,
        include_audio=include_audio,
        log_examples=log_examples,
    ).files


def copy_found_files(
    files: Sequence[FoundFile],
    azs: AzsConfig,
    general_cfg: GeneralConfig,
    stop_event: Optional[object] = None,
    on_file: Optional[Callable[[CopyRecord], None]] = None,
    on_progress: Optional[Callable[[int, int], None]] = None,
    row_index: int = 0,
    trk: Optional[int] = None,
    docdate: Optional[datetime] = None,
) -> List[CopyRecord]:
    """Скопировать найденные файлы «зеркалом источника».

    Общая функция для xlsx-режима и ручного копирования: вся логика
    ``copy2`` / ``already_exists`` / ``duplicate_in_run`` /
    ``size_mismatch_overwritten`` живёт здесь (:func:`copy_one`).

    :param files: найденные файлы
    :param azs: текущая АЗС (используется ``output_root``)
    :param general_cfg: общие настройки (``copy_mode``, ``threads``)
    :param stop_event: объект с методом ``is_set()`` для остановки
    :param on_file: колбэк после каждого файла, получает :class:`CopyRecord`
    :param on_progress: колбэк ``(обработано, всего)``
    :param row_index: номер строки для отчёта (строка xlsx или ручная строка)
    :param trk: номер ТРК для отчёта (если известен)
    :param docdate: время заправки/начало интервала для отчёта
    :return: записи отчёта по каждому файлу
    """
    output_root = Path(azs.output_root)
    mode = general_cfg.copy_mode
    threads = max(1, int(general_cfg.threads))
    items = list(files)
    total = len(items)
    records: List[CopyRecord] = []
    # Внутри одного прогона один и тот же файл не копируется дважды
    # (например, если ТРК 1 и ТРК 2 ссылаются на одну камеру).
    seen_src: set = set()

    def stopped() -> bool:
        return stop_event is not None and bool(getattr(stop_event, "is_set", lambda: False)())

    def make_record(
        item: FoundFile, status: str, dst: Optional[Path], error: str
    ) -> CopyRecord:
        return CopyRecord(
            row_index=row_index,
            azs=azs.number,
            docdate=docdate,
            trk=trk,
            camera=item.camera,
            kind=item.kind,
            src=item.src,
            dst=dst,
            status=status,
            error=error,
        )

    def handle(item: FoundFile, outcome: Tuple[str, Path, str]) -> None:
        status, dst, error = outcome
        if status == STATUS_SIZE_MISMATCH:
            LOGGER.warning("Перезаписан файл с другим размером: %s", dst)
        key = str(item.src).lower()
        if status in (STATUS_COPIED, STATUS_SIZE_MISMATCH) and key in seen_src:
            LOGGER.debug("Файл уже скопирован в этом прогоне: %s", item.src)
            status, error = STATUS_ALREADY_EXISTS, ""
        if status in (STATUS_COPIED, STATUS_SIZE_MISMATCH, STATUS_ALREADY_EXISTS):
            seen_src.add(key)
        record = make_record(item, status, dst, error)
        records.append(record)
        if on_progress is not None:
            on_progress(len(records), total)
        if on_file is not None:
            on_file(record)

    try:
        pending = [item for item in items]
        if threads > 1 and total > 1:
            with ThreadPoolExecutor(max_workers=threads) as executor:
                pairs = []
                for item in pending:
                    if stopped():
                        break
                    future = executor.submit(
                        copy_one, item.src, item.video_root, output_root, mode
                    )
                    pairs.append((item, future))
                for item, future in pairs:
                    try:
                        outcome = future.result()
                    except Exception as exc:  # защита от неожиданных ошибок потока
                        outcome = (
                            STATUS_ERROR,
                            destination_for(item.src, item.video_root, output_root),
                            f"Ошибка копирования: {exc}",
                        )
                    handle(item, outcome)
        else:
            for item in pending:
                if stopped():
                    break
                try:
                    outcome = copy_one(item.src, item.video_root, output_root, mode)
                except Exception as exc:  # защита от неожиданных ошибок
                    outcome = (
                        STATUS_ERROR,
                        destination_for(item.src, item.video_root, output_root),
                        f"Ошибка копирования: {exc}",
                    )
                handle(item, outcome)
    except OSError as exc:
        LOGGER.error("Ошибка копирования: %s", exc)

    return records


class JobRunner:
    """Обработка строк xlsx для одной выбранной АЗС."""

    def __init__(self, config: AppConfig, azs: AzsConfig) -> None:
        self.config = config
        self.azs = azs
        self.total_stats = SearchStats()
        self._examples_logged = False

    # ------------------------------------------------------------------ #
    # Публичные методы
    # ------------------------------------------------------------------ #
    def get_video_roots(self, force_rescan: bool = False) -> List[Path]:
        """Корни видеоархива для текущей АЗС."""
        return drives.resolve_video_roots(self.config, self.azs, force_rescan=force_rescan)

    def run(
        self,
        rows: Sequence[RowData],
        scan_only: bool = False,
        stop_event: Optional[object] = None,
        on_row: Optional[Callable[[RowResult], None]] = None,
        on_progress: Optional[Callable[[int, int], None]] = None,
    ) -> List[RowResult]:
        """Обработать все строки xlsx.

        :param rows: строки xlsx
        :param scan_only: только поиск без копирования
        :param stop_event: объект с методом ``is_set()`` для остановки
        :param on_row: колбэк, вызываемый после каждой строки
        :param on_progress: колбэк ``(обработано, всего)``
        :return: список результатов по строкам
        """
        results: List[RowResult] = []
        total = len(rows)
        video_roots = self.get_video_roots()
        LOGGER.info(
            "АЗС %s: output_root=%s; video_roots=%s; интервал %s",
            self.azs.number,
            self.azs.output_root,
            ", ".join(str(root) for root in video_roots) or "нет",
            self.config.general.interval_text,
        )

        for position, row in enumerate(rows, start=1):
            if stop_event is not None and getattr(stop_event, "is_set", lambda: False)():
                LOGGER.info("Обработка остановлена пользователем на строке %s.", row.index)
                for skipped in rows[position - 1 :]:
                    results.append(self._make_stopped_result(skipped))
                    if on_row is not None:
                        on_row(results[-1])
                break

            result = self.process_row(row, video_roots, scan_only=scan_only, stop_event=stop_event)
            results.append(result)
            if on_row is not None:
                on_row(result)
            if on_progress is not None:
                on_progress(position, total)

        LOGGER.info(
            "Обойдено каталогов: %s, отсечено фильтром по дате/часу: %s, проверено файлов: %s, отобрано: %s",
            self.total_stats.dirs_visited,
            self.total_stats.dirs_skipped_by_date,
            self.total_stats.files_checked,
            self.total_stats.files_matched,
        )
        return results

    def process_row(
        self,
        row: RowData,
        video_roots: Sequence[Path],
        scan_only: bool = False,
        stop_event: Optional[object] = None,
    ) -> RowResult:
        """Обработать одну строку xlsx: найти файлы и (при необходимости) скопировать."""
        result = RowResult(row=row)

        if row.error:
            result.status = STATUS_ERROR
            result.error = f"Строка {row.index}: {row.error}"
            LOGGER.warning(result.error)
            return result

        if row.azs is not None and int(row.azs) != self.azs.number:
            result.status = STATUS_SKIPPED_OTHER_AZS
            result.error = (
                f"Строка {row.index}: АЗС {row.azs} не совпадает с выбранной АЗС {self.azs.number}"
            )
            LOGGER.warning(result.error)
            return result

        cameras = self.azs.cameras_for_trk(int(row.trk)) if row.trk is not None else None
        if not cameras:
            result.status = STATUS_ERROR
            result.error = (
                f"Строка {row.index}: в azs.ini для АЗС {self.azs.number} "
                f"не задан список камер для ТРК {row.trk}."
            )
            LOGGER.warning(result.error)
            return result

        result.cameras = list(cameras)
        assert row.docdate is not None  # проверено выше

        # Никакой логики по литрам (realdose): интервал зависит только от ini.
        result.start = row.docdate - timedelta(minutes=self.config.general.minutes_before)
        result.end = row.docdate + timedelta(minutes=self.config.general.minutes_after)

        if not video_roots:
            result.status = STATUS_ERROR
            result.error = "Не найдено ни одного каталога VIDEO — проверьте настройки video_roots."
            LOGGER.warning(result.error)
            return result

        search = search_files_for_cameras(
            azs=self.azs,
            cameras=cameras,
            start=result.start,
            end=result.end,
            finder_cfg=self.config.finder,
            video_roots=video_roots,
            include_audio=self.config.general.include_audio,
            log_examples=not self._examples_logged,
        )
        self._examples_logged = True
        self.total_stats.merge(search.stats)

        result.files = search.files
        result.found_video = search.count("video")
        result.found_audio = search.count("audio")

        LOGGER.info(
            "AZS=%s, docdate=%s, интервал=%s..%s, открыто час-каталогов=%s, "
            "найдено видео=%s, аудио=%s",
            self.azs.number,
            row.docdate.strftime("%d.%m.%Y %H:%M:%S"),
            result.start.strftime("%d.%m.%Y %H:%M:%S"),
            result.end.strftime("%d.%m.%Y %H:%M:%S"),
            search.stats.dirs_visited,
            result.found_video,
            result.found_audio,
        )

        if scan_only:
            result.status = STATUS_FOUND if search.files else STATUS_NOT_FOUND
            return result

        self._copy_row_files(result, stop_event=stop_event)
        self._finalize(result)
        return result

    # ------------------------------------------------------------------ #
    # Внутренние методы
    # ------------------------------------------------------------------ #
    def _copy_row_files(self, result: RowResult, stop_event: Optional[object] = None) -> None:
        """Скопировать найденные файлы строки и заполнить записи отчёта."""
        records = copy_found_files(
            files=result.files,
            azs=self.azs,
            general_cfg=self.config.general,
            stop_event=stop_event,
            row_index=result.row.index,
            trk=result.row.trk,
            docdate=result.row.docdate,
        )
        result.records.extend(records)

        # Отметить камеры, по которым вообще не нашлось файлов нужного типа.
        present = {(item.camera, item.kind) for item in result.files}
        for camera in result.cameras:
            for kind in ("video", "audio"):
                if kind == "audio" and not self.config.general.include_audio:
                    continue
                if (camera, kind) in present:
                    continue
                result.records.append(
                    CopyRecord(
                        row_index=result.row.index,
                        azs=result.row.azs,
                        docdate=result.row.docdate,
                        trk=result.row.trk,
                        camera=camera,
                        kind=kind,
                        src=None,
                        dst=None,
                        status=STATUS_NOT_FOUND,
                        error="файлы за интервал не найдены",
                    )
                )

    def _finalize(self, result: RowResult) -> None:
        """Подсчитать итоги строки и выставить общий статус."""
        result.copied = sum(1 for rec in result.records if rec.status == STATUS_COPIED)
        result.already_exists = sum(
            1 for rec in result.records if rec.status == STATUS_ALREADY_EXISTS
        )
        result.overwritten = sum(
            1 for rec in result.records if rec.status == STATUS_SIZE_MISMATCH
        )
        result.not_found = sum(1 for rec in result.records if rec.status == STATUS_NOT_FOUND)
        errors = [rec for rec in result.records if rec.status == STATUS_ERROR]

        if errors:
            result.status = STATUS_ERROR
            result.error = "; ".join(
                f"{rec.camera}/{rec.kind}: {rec.error}" for rec in errors[:3]
            )
            if len(errors) > 3:
                result.error += f" (и ещё {len(errors) - 3})"
        elif result.copied or result.overwritten:
            result.status = STATUS_COPIED
        elif result.already_exists:
            result.status = STATUS_ALREADY_EXISTS
        else:
            result.status = STATUS_NOT_FOUND

        LOGGER.info("Строка %s: %s", result.row.index, result.summary_text)

    def _make_stopped_result(self, row: RowData) -> RowResult:
        """Заглушка результата для строк, не обработанных из-за остановки."""
        result = RowResult(row=row, status=STATUS_STOPPED, error="обработка остановлена")
        return result


# --------------------------------------------------------------------------- #
# Отчёт CSV
# --------------------------------------------------------------------------- #
REPORT_COLUMNS = [
    "azs",
    "docdate",
    "trk",
    "camera",
    "kind",
    "src",
    "dst",
    "status",
    "error",
]

SUMMARY_COLUMNS = [
    "azs",
    "docdate",
    "trk",
    "cameras",
    "start",
    "end",
    "found_video",
    "found_audio",
    "copied",
    "already_exists",
    "overwritten",
    "not_found",
    "status",
    "error",
]


def write_reports(
    output_root: Path, results: Sequence[RowResult], tag_prefix: str = "report"
) -> Tuple[Optional[Path], Optional[Path]]:
    """Записать два CSV-отчёта: по файлам и сводку по строкам xlsx.

    :param output_root: каталог, куда писать отчёты
    :param results: результаты обработки строк
    :param tag_prefix: префикс имени файла
    :return: кортеж (путь к отчёту по файлам, путь к сводке)
    """
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    records_path = Path(output_root) / f"{tag_prefix}_{stamp}.csv"
    summary_path = Path(output_root) / f"{tag_prefix}_{stamp}_summary.csv"

    try:
        Path(output_root).mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        LOGGER.error("Не удалось создать каталог для отчёта %s: %s", output_root, exc)
        return None, None

    try:
        with records_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(REPORT_COLUMNS)
            for result in results:
                for record in result.records:
                    writer.writerow(record.as_csv_row())

        with summary_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(SUMMARY_COLUMNS)
            for result in results:
                writer.writerow(
                    [
                        "" if result.row.azs is None else str(result.row.azs),
                        result.row.docdate_text,
                        "" if result.row.trk is None else str(result.row.trk),
                        result.cameras_text,
                        result.start_text,
                        result.end_text,
                        result.found_video,
                        result.found_audio,
                        result.copied,
                        result.already_exists,
                        result.overwritten,
                        result.not_found,
                        result.status,
                        result.error,
                    ]
                )
    except OSError as exc:
        LOGGER.error("Не удалось записать отчёт: %s", exc)
        return None, None

    LOGGER.info("Отчёт по файлам: %s", records_path)
    LOGGER.info("Сводка по строкам: %s", summary_path)
    return records_path, summary_path
