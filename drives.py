"""Автообнаружение локальных дисков и каталогов ``VIDEO``.

Функции модуля определяют, где лежит видеоархив СВН «Интеллект»:

* :func:`list_available_drives` — перебор букв A..Z;
* :func:`discover_video_roots` — поиск ``X:\\VIDEO`` на всех доступных дисках
  (результат кэшируется на время сессии, пока его не сбросят);
* :func:`resolve_video_roots` — итоговый список корней для выбранной АЗС:
  явные пути из ``azs.ini`` имеют приоритет над общим параметром ``auto``.
"""

from __future__ import annotations

import logging
import string
from pathlib import Path
from typing import List, Optional, Sequence

from config import AppConfig, AzsConfig, parse_roots

LOGGER = logging.getLogger(__name__)

#: Кэш результатов сканирования дисков и имя подкаталога, для которого он построен.
_CACHE: Optional[List[Path]] = None
_CACHE_SUBDIR: Optional[str] = None


def list_available_drives() -> List[str]:
    """Вернуть список доступных корней дисков вида ``['C:\\\\', 'D:\\\\']``."""
    drives: List[str] = []
    for letter in string.ascii_uppercase:
        root = Path(f"{letter}:\\")
        try:
            if root.exists():
                drives.append(f"{letter}:\\")
        except OSError:
            continue
    return drives


def discover_video_roots(subdir: str = "VIDEO", force: bool = False) -> List[Path]:
    """Найти все каталоги ``X:\\<subdir>`` на доступных дисках.

    Результат кэшируется до вызова с ``force=True``.

    :param subdir: имя подкаталога (по умолчанию VIDEO)
    :param force: принудительно повторить сканирование
    :return: список найденных корней видеоархива
    """
    global _CACHE, _CACHE_SUBDIR

    if _CACHE is not None and _CACHE_SUBDIR == subdir and not force:
        return list(_CACHE)

    roots: List[Path] = []
    for drive in list_available_drives():
        candidate = Path(drive) / subdir
        try:
            if candidate.is_dir():
                roots.append(candidate)
        except OSError as exc:  # например, диск отключился во время проверки
            LOGGER.warning("Не удалось проверить каталог %s: %s", candidate, exc)

    _CACHE = roots
    _CACHE_SUBDIR = subdir

    LOGGER.info(
        "Обнаружено дисков с %s: %s: [%s]",
        subdir,
        len(roots),
        ", ".join(str(item) for item in roots) if roots else "нет",
    )
    if not roots:
        LOGGER.warning(
            "Каталоги %s не найдены ни на одном диске. Проверьте, подключены ли диски с видеоархивом.",
            subdir,
        )
    return list(roots)


def reset_cache() -> None:
    """Сбросить кэш обнаруженных корней (кнопка «Пересканировать диски»)."""
    global _CACHE, _CACHE_SUBDIR
    _CACHE = None
    _CACHE_SUBDIR = None


def resolve_video_roots(
    config: AppConfig,
    azs: Optional[AzsConfig] = None,
    force_rescan: bool = False,
) -> List[Path]:
    """Определить список корней видеоархива для выбранной АЗС.

    Приоритет:

    1. ``[AZS_<номер>].video_roots`` из azs.ini (если задан явно);
    2. ``[general].video_roots`` со значением ``auto`` — автообнаружение;
    3. ``[general].video_roots`` со списком явных путей через ``;``.

    :param config: полные настройки приложения
    :param azs: настройки выбранной АЗС (может быть ``None``)
    :param force_rescan: повторить сканирование дисков
    :return: список существующих (в идеале) корней видеоархива
    """
    if azs is not None and azs.video_roots:
        LOGGER.info(
            "Источник видео: явные пути из azs.ini для АЗС %s: %s",
            azs.number,
            ", ".join(str(item) for item in azs.video_roots),
        )
        return list(azs.video_roots)

    raw = config.general.video_roots_raw.strip()
    subdir = config.general.video_root_subdir

    if raw.lower() == "auto":
        roots = discover_video_roots(subdir=subdir, force=force_rescan)
        return roots

    roots = parse_roots(raw)
    LOGGER.info(
        "Источник видео: явные пути из settings.ini: %s",
        ", ".join(str(item) for item in roots) or "нет",
    )
    return roots


def check_roots(roots: Sequence[Path]) -> List[tuple]:
    """Проверить существование корней видеоархива.

    :param roots: список путей
    :return: список кортежей ``(путь, существует_ли)``
    """
    result = []
    for root in roots:
        try:
            exists = Path(root).is_dir()
        except OSError:
            exists = False
        result.append((Path(root), exists))
    return result
