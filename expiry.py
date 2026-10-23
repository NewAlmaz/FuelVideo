"""Проверка срока действия, пароля разблокировки и отката системных часов."""

from __future__ import annotations

import hashlib
import logging
import os
from datetime import date, datetime
from pathlib import Path
from typing import Optional, Tuple

from license import EXPIRY_DATE, UNLOCK_PASSWORD_HASH

LOGGER = logging.getLogger(__name__)

LAST_SEEN_FILE = "last_seen.dat"


def get_last_seen_path() -> Path:
    """Путь к файлу last_seen.dat в %LOCALAPPDATA%\\FuelVideo."""
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    folder = Path(base) / "FuelVideo"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / LAST_SEEN_FILE


def is_expired(today: date | None = None) -> bool:
    """Истёк ли срок действия (сегодня > EXPIRY_DATE)."""
    today = today or date.today()
    return today > EXPIRY_DATE


def expiry_text() -> str:
    """Строка для логов и GUI, например «до 23.02.2027»."""
    return f"до {EXPIRY_DATE:%d.%m.%Y}"


def verify_unlock_password(password: str) -> bool:
    """Проверить пароль разблокировки по SHA-256 хэшу."""
    if not UNLOCK_PASSWORD_HASH:
        return False
    digest = hashlib.sha256(password.encode("utf-8")).hexdigest().lower()
    return digest == UNLOCK_PASSWORD_HASH.lower()


def check_clock_rollback() -> Tuple[bool, Optional[date]]:
    """Проверить, откатывали ли системные часы назад.

    Сравнивает текущую дату с последней сохранённой в ``last_seen.dat``.

    :return: пара ``(обнаружен_откат, сохранённая_дата)``:

        * файла не было — файл создан, ``(False, None)``;
        * ``today < saved`` — откат, ``(True, saved)`` (файл НЕ перезаписывается);
        * ``today > saved`` — файл обновлён, ``(False, saved)``;
        * ``today == saved`` — ``(False, saved)``;
        * ошибка чтения/разбора — ``(False, None)`` и warning в лог.
    """
    path = get_last_seen_path()
    today = date.today()
    try:
        if path.exists():
            raw = path.read_text(encoding="utf-8").strip()
            saved = datetime.strptime(raw, "%Y-%m-%d").date()
            if today < saved:
                LOGGER.warning(
                    "expiry: обнаружен откат часов (today=%s < saved=%s)", today, saved
                )
                return True, saved
            if today > saved:
                path.write_text(today.isoformat(), encoding="utf-8")
            return False, saved
        path.write_text(today.isoformat(), encoding="utf-8")
        LOGGER.info("expiry: last_seen.dat создан (%s)", today)
        return False, None
    except (OSError, ValueError) as exc:
        LOGGER.warning("expiry: не удалось проверить last_seen.dat: %s", exc)
        return False, None


def should_block(rollback_detected: bool, saved_date: Optional[date]) -> bool:
    """Нужно ли блокировать работу по сроку действия / откату часов.

    Откат часов сам по себе блокировкой НЕ является: если реальный срок ещё не
    вышел, программа работает. Блокировка включается, только если

    * реальный срок уже вышел (``is_expired()``), либо
    * часы откатили, и до отката дата уже была после ``EXPIRY_DATE``
      (то есть откат явно сделан, чтобы обойти окончание срока).
    """
    if is_expired():
        return True
    if rollback_detected and saved_date is not None and saved_date > EXPIRY_DATE:
        return True
    return False


def rollback_message(rollback_detected: bool, saved_date: Optional[date]) -> str:
    """Текст предупреждения об откате часов (пустая строка, если отката нет)."""
    if not rollback_detected:
        return ""
    if saved_date is None:
        return "Обнаружен откат системных часов."
    return f"Обнаружен откат системных часов (до отката была дата {saved_date:%d.%m.%Y})."
