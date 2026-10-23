"""Проверка срока действия, пароля разблокировки и отката системных часов."""

from __future__ import annotations

import hashlib
import logging
import os
from datetime import date, datetime
from pathlib import Path

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


def check_clock_rollback() -> bool:
    """Вернуть True, если системные часы откатили назад.

    Сравнивает текущую дату с последней сохранённой в last_seen.dat.
    Если файла нет — создаёт и возвращает False.
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
                return True
            if today > saved:
                path.write_text(today.isoformat(), encoding="utf-8")
            return False
        path.write_text(today.isoformat(), encoding="utf-8")
        LOGGER.info("expiry: last_seen.dat создан (%s)", today)
        return False
    except (OSError, ValueError) as exc:
        LOGGER.warning("expiry: не удалось проверить last_seen.dat: %s", exc)
        return False
