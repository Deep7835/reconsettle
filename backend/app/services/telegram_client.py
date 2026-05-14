"""Minimal synchronous wrapper around the Telegram Bot HTTP API.

We deliberately avoid `python-telegram-bot` to keep deps small — the bot API
is just HTTPS calls and httpx handles them fine. All methods raise
TelegramError on non-200 responses with the description from Telegram.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional

import httpx

from app.config import settings


class TelegramError(RuntimeError):
    pass


class TelegramNotConfigured(TelegramError):
    pass


API_BASE = "https://api.telegram.org"


def _token() -> str:
    if not settings.telegram_bot_token:
        raise TelegramNotConfigured(
            "Set TELEGRAM_BOT_TOKEN in .env (create the bot via @BotFather on Telegram)."
        )
    return settings.telegram_bot_token


def _request(method: str, params: Optional[dict] = None, files: Optional[dict] = None, timeout: float = 30.0) -> dict:
    url = f"{API_BASE}/bot{_token()}/{method}"
    try:
        with httpx.Client(timeout=timeout) as client:
            r = client.post(url, data=params, files=files)
    except httpx.HTTPError as e:
        raise TelegramError(f"Network error talking to Telegram: {e}") from e
    if r.status_code != 200:
        try:
            j = r.json()
            desc = j.get("description") or r.text
        except Exception:
            desc = r.text
        raise TelegramError(f"Telegram {method} failed ({r.status_code}): {desc}")
    data = r.json()
    if not data.get("ok"):
        raise TelegramError(f"Telegram {method} returned !ok: {data.get('description')}")
    return data["result"]


# ---------- Public methods ----------


@dataclass
class TelegramUser:
    id: int
    is_bot: bool
    first_name: str
    username: Optional[str]


def get_me() -> TelegramUser:
    """Verify the bot token. Cheapest call to test connectivity."""
    r = _request("getMe")
    return TelegramUser(
        id=r["id"],
        is_bot=r.get("is_bot", True),
        first_name=r.get("first_name", ""),
        username=r.get("username"),
    )


def get_updates(offset: Optional[int] = None, timeout_seconds: int = 0, limit: int = 50) -> list[dict]:
    """Long-poll for updates. With timeout_seconds=0 this returns immediately."""
    params = {"timeout": timeout_seconds, "limit": limit}
    if offset is not None:
        params["offset"] = offset
    return _request("getUpdates", params=params, timeout=timeout_seconds + 10)


def send_message(
    chat_id: str,
    text: str,
    parse_mode: Optional[str] = "Markdown",
    disable_web_page_preview: bool = True,
) -> dict:
    params = {
        "chat_id": chat_id,
        "text": text,
        "disable_web_page_preview": str(disable_web_page_preview).lower(),
    }
    if parse_mode and parse_mode != "none":
        params["parse_mode"] = parse_mode
    return _request("sendMessage", params=params)


def send_document(
    chat_id: str,
    filename: str,
    data: bytes,
    caption: Optional[str] = None,
    parse_mode: Optional[str] = "Markdown",
) -> dict:
    params = {"chat_id": chat_id}
    if caption:
        params["caption"] = caption[:1024]
        if parse_mode and parse_mode != "none":
            params["parse_mode"] = parse_mode
    files = {"document": (filename, data, "application/octet-stream")}
    return _request("sendDocument", params=params, files=files, timeout=60.0)


def get_file(file_id: str) -> dict:
    """Resolve a file_id to file_path so we can download it."""
    return _request("getFile", params={"file_id": file_id})


def download_file(file_path: str, timeout: float = 60.0) -> bytes:
    """Download a previously resolved file (per Bot API: <=20 MB)."""
    url = f"{API_BASE}/file/bot{_token()}/{file_path}"
    try:
        with httpx.Client(timeout=timeout) as client:
            r = client.get(url)
            r.raise_for_status()
            return r.content
    except httpx.HTTPError as e:
        raise TelegramError(f"Failed to download {file_path}: {e}") from e
