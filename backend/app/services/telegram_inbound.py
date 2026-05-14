"""Inbound poller for Telegram: pull new messages, download any Excel/CSV files,
and feed them through the same excel_parser pipeline as Gmail (BankUpload /
SourceUpload).

State (last seen update_id) lives in `data/telegram_state.json` so we never
re-process the same message twice.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from sqlalchemy.orm import Session

from app.config import settings
from app.models import BankUpload, SourceUpload, TelegramRun
from app.services import excel_parser, telegram_client


logger = logging.getLogger("settleops.telegram")


ALLOWED_EXTENSIONS = (".xlsx", ".xls", ".csv")


@dataclass
class InboundResult:
    run: TelegramRun
    saved_files: List[str]


# ---------- State persistence ----------


def _state_path() -> Path:
    return Path(settings.telegram_state_file)


def _load_last_update_id() -> Optional[int]:
    p = _state_path()
    if not p.exists():
        return None
    try:
        return int(json.loads(p.read_text(encoding="utf-8")).get("last_update_id"))
    except Exception:
        return None


def _save_last_update_id(update_id: int) -> None:
    p = _state_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({"last_update_id": update_id}), encoding="utf-8")


# ---------- Chat allowlist ----------


def _chat_allowed(chat_id_str: str) -> bool:
    allowlist = settings.telegram_inbound_chat_ids
    if not allowlist:
        return True  # empty allowlist = accept all
    return chat_id_str in {str(c) for c in allowlist}


# ---------- Core sync ----------


def sync(db: Session, triggered_by: str) -> InboundResult:
    """Pull updates, download attachments, parse, persist as Upload rows."""
    kind = "inbound_scheduler" if triggered_by == "scheduler" else "inbound_manual"
    started = datetime.utcnow()
    run = TelegramRun(
        kind=kind,
        status="ok",
        started_at=started,
        messages_seen=0,
        attachments_saved=0,
        uploads_created=0,
        payload={"allowlist": settings.telegram_inbound_chat_ids},
    )
    db.add(run)
    db.flush()

    saved_files: list[str] = []

    try:
        last = _load_last_update_id()
        offset = (last + 1) if last is not None else None
        updates = telegram_client.get_updates(offset=offset, timeout_seconds=0, limit=50)
    except telegram_client.TelegramNotConfigured as e:
        run.status = "error"
        run.error_message = str(e)
        run.finished_at = datetime.utcnow()
        db.commit()
        db.refresh(run)
        return InboundResult(run=run, saved_files=saved_files)
    except telegram_client.TelegramError as e:
        run.status = "error"
        run.error_message = str(e)
        run.finished_at = datetime.utcnow()
        db.commit()
        db.refresh(run)
        return InboundResult(run=run, saved_files=saved_files)

    max_update_id: Optional[int] = None
    seen_chat_ids: set[str] = set()

    for upd in updates:
        update_id = upd.get("update_id")
        if update_id is not None and (max_update_id is None or update_id > max_update_id):
            max_update_id = update_id

        msg = upd.get("message") or upd.get("channel_post")
        if not msg:
            continue
        run.messages_seen += 1

        chat = msg.get("chat") or {}
        chat_id = str(chat.get("id") or "")
        seen_chat_ids.add(chat_id)
        if not _chat_allowed(chat_id):
            continue

        document = msg.get("document")
        if not document:
            continue

        filename = document.get("file_name") or f"telegram_{document.get('file_unique_id', 'file')}.bin"
        if not filename.lower().endswith(ALLOWED_EXTENSIONS):
            continue

        try:
            meta = telegram_client.get_file(document["file_id"])
            data = telegram_client.download_file(meta["file_path"])
        except telegram_client.TelegramError as e:
            logger.warning("Telegram download failed for %s: %s", filename, e)
            continue

        run.attachments_saved += 1
        try:
            _store_attachment(db, filename, data, source="telegram")
            run.uploads_created += 1
            saved_files.append(filename)
        except Exception as e:  # noqa: BLE001
            logger.warning("Couldn't parse %s from Telegram: %s", filename, e)

    if max_update_id is not None:
        _save_last_update_id(max_update_id)

    if seen_chat_ids:
        run.chat_id = ",".join(sorted(seen_chat_ids))

    run.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(run)
    return InboundResult(run=run, saved_files=saved_files)


# ---------- Internal: file routing (mirrors Gmail's heuristic) ----------


def _is_bank_filename(name: str) -> bool:
    name = name.lower()
    return any(k in name for k in ("bank", "settle", "mid"))


def _store_attachment(db: Session, filename: str, data: bytes, source: str) -> None:
    if _is_bank_filename(filename):
        parsed = excel_parser.parse_bank_excel(data)
        db.add(
            BankUpload(
                filename=filename,
                sha256=excel_parser.file_sha256(data),
                source=source,
                row_count=len(parsed.rows),
                total_amount=round(parsed.total_amount, 2),
                rows_json=parsed.rows,
            )
        )
    else:
        parsed = excel_parser.parse_source_excel(data)
        db.add(
            SourceUpload(
                filename=filename,
                sha256=excel_parser.file_sha256(data),
                source=source,
                row_count=len(parsed.rows),
                total_amount=round(parsed.total_amount, 2),
                rows_json=parsed.rows,
            )
        )
