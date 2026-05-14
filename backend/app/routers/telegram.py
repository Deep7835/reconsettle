"""/api/telegram — config + bot identity + inbound sync + send message/report + runs."""

from __future__ import annotations

from datetime import datetime
from typing import List

from fastapi import APIRouter, Depends, HTTPException, Query, Response, status
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app.config import settings
from app.db import get_db
from app.models import TelegramRecipient, TelegramRun
from app.schemas.telegram import (
    RecipientCreate,
    RecipientOut,
    RecipientUpdate,
    TelegramConfig,
    TelegramConfigUpdate,
    TelegramMe,
    TelegramRunOut,
    TelegramSendMessage,
    TelegramSendReport,
    TelegramSyncResult,
)
from app.services import report_builder, scheduler, telegram_client, telegram_inbound


router = APIRouter(prefix="/api/telegram", tags=["telegram"])


# ---------- Config ----------


@router.get("/config", response_model=TelegramConfig)
def get_config() -> TelegramConfig:
    return TelegramConfig(
        has_bot_token=bool(settings.telegram_bot_token),
        report_chat_id=settings.telegram_report_chat_id,
        inbound_chat_ids=settings.telegram_inbound_chat_ids,
        sync_interval_minutes=settings.telegram_sync_interval_minutes,
        digest_cron=settings.telegram_digest_cron,
    )


@router.patch("/config", response_model=TelegramConfig)
def update_config(payload: TelegramConfigUpdate) -> TelegramConfig:
    """Runtime config update. Persists for this process only — edit .env for permanence."""
    if payload.bot_token is not None:
        settings.telegram_bot_token = payload.bot_token
    if payload.report_chat_id is not None:
        settings.telegram_report_chat_id = payload.report_chat_id
    if payload.inbound_chat_ids is not None:
        settings.telegram_inbound_chat_ids = payload.inbound_chat_ids
    if payload.sync_interval_minutes is not None:
        settings.telegram_sync_interval_minutes = payload.sync_interval_minutes
        scheduler.update_telegram_job(payload.sync_interval_minutes)
    if payload.digest_cron is not None:
        settings.telegram_digest_cron = payload.digest_cron
        scheduler.update_telegram_digest(payload.digest_cron)
    return get_config()


# ---------- Bot identity ----------


@router.get("/me", response_model=TelegramMe)
def whoami() -> TelegramMe:
    try:
        u = telegram_client.get_me()
    except telegram_client.TelegramNotConfigured as e:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, str(e)) from e
    except telegram_client.TelegramError as e:
        raise HTTPException(status.HTTP_502_BAD_GATEWAY, str(e)) from e
    return TelegramMe(id=u.id, is_bot=u.is_bot, first_name=u.first_name, username=u.username)


# ---------- Inbound ----------


@router.post("/sync", response_model=TelegramSyncResult)
def sync_now(db: Session = Depends(get_db)) -> TelegramSyncResult:
    result = telegram_inbound.sync(db=db, triggered_by="manual")
    return TelegramSyncResult(
        run=TelegramRunOut.model_validate(result.run),
        saved_files=result.saved_files,
    )


@router.post("/state/reset")
def reset_inbound_state() -> dict:
    """Forget last_update_id so the next sync re-pulls everything Telegram still has cached.

    Telegram retains updates for ~24h. Useful when you set up the bot, sent
    messages with privacy mode on (so they were never delivered to the bot),
    or just want to re-process the recent backlog.
    """
    from pathlib import Path

    from app.config import settings
    path = Path(settings.telegram_state_file)
    existed = path.exists()
    if existed:
        path.unlink()
    return {"reset": True, "had_state": existed, "state_file": str(path)}


@router.get("/runs", response_model=List[TelegramRunOut])
def list_runs(limit: int = Query(30, ge=1, le=200), db: Session = Depends(get_db)) -> List[TelegramRunOut]:
    rows = db.query(TelegramRun).order_by(TelegramRun.id.desc()).limit(limit).all()
    return [TelegramRunOut.model_validate(r) for r in rows]


# ---------- Outbound ----------


def _resolve_chat(override: str | None) -> str:
    chat = override or settings.telegram_report_chat_id
    if not chat:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            "No chat_id provided and TELEGRAM_REPORT_CHAT_ID is empty.",
        )
    return chat


@router.post("/send/message", response_model=TelegramRunOut)
def send_message(payload: TelegramSendMessage, db: Session = Depends(get_db)) -> TelegramRunOut:
    chat_id = _resolve_chat(payload.chat_id)
    started = datetime.utcnow()
    run = TelegramRun(
        kind="outbound_message",
        status="ok",
        started_at=started,
        chat_id=chat_id,
        payload={"text_len": len(payload.text)},
    )
    db.add(run)
    db.flush()
    try:
        telegram_client.send_message(
            chat_id=chat_id, text=payload.text, parse_mode=payload.parse_mode
        )
    except telegram_client.TelegramError as e:
        run.status = "error"
        run.error_message = str(e)
    run.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(run)
    if run.status == "error":
        raise HTTPException(status.HTTP_502_BAD_GATEWAY, run.error_message or "Send failed")
    return TelegramRunOut.model_validate(run)


@router.post("/send/report", response_model=TelegramRunOut)
def send_report(payload: TelegramSendReport, db: Session = Depends(get_db)) -> TelegramRunOut:
    """Send a markdown daily digest + optional Excel attachment to Telegram."""
    chat_id = _resolve_chat(payload.chat_id)
    started = datetime.utcnow()
    run = TelegramRun(
        kind="outbound_report",
        status="ok",
        started_at=started,
        chat_id=chat_id,
        payload={"include_excel": payload.include_excel},
    )
    db.add(run)
    db.flush()

    text = report_builder.build_digest_text(db)
    try:
        telegram_client.send_message(chat_id=chat_id, text=text, parse_mode="Markdown")
        if payload.include_excel:
            filename, data = report_builder.build_entries_xlsx(db)
            caption = payload.caption or f"Settlements export — {datetime.utcnow().date().isoformat()}"
            telegram_client.send_document(
                chat_id=chat_id, filename=filename, data=data, caption=caption
            )
    except telegram_client.TelegramError as e:
        run.status = "error"
        run.error_message = str(e)
    run.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(run)
    if run.status == "error":
        raise HTTPException(status.HTTP_502_BAD_GATEWAY, run.error_message or "Send failed")
    return TelegramRunOut.model_validate(run)


# ---------- Recipient contact book ----------


@router.get("/recipients", response_model=List[RecipientOut])
def list_recipients(db: Session = Depends(get_db)) -> List[RecipientOut]:
    rows = db.query(TelegramRecipient).order_by(TelegramRecipient.name).all()
    return [RecipientOut.model_validate(r) for r in rows]


@router.post("/recipients", response_model=RecipientOut, status_code=status.HTTP_201_CREATED)
def add_recipient(payload: RecipientCreate, db: Session = Depends(get_db)) -> RecipientOut:
    r = TelegramRecipient(
        name=payload.name.strip(),
        chat_id=payload.chat_id.strip(),
        description=(payload.description or "").strip() or None,
        is_active=payload.is_active,
        digest_cron=(payload.digest_cron or "").strip(),
        include_excel=payload.include_excel,
    )
    db.add(r)
    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            f"A recipient with chat_id {payload.chat_id} already exists.",
        )
    db.refresh(r)
    scheduler.update_recipient_digest_job(r)
    return RecipientOut.model_validate(r)


@router.patch("/recipients/{recipient_id}", response_model=RecipientOut)
def update_recipient(recipient_id: int, payload: RecipientUpdate, db: Session = Depends(get_db)) -> RecipientOut:
    r = db.get(TelegramRecipient, recipient_id)
    if not r:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Recipient not found")
    if payload.name is not None:
        r.name = payload.name.strip()
    if payload.chat_id is not None:
        r.chat_id = payload.chat_id.strip()
    if payload.description is not None:
        r.description = payload.description.strip() or None
    if payload.is_active is not None:
        r.is_active = payload.is_active
    if payload.digest_cron is not None:
        r.digest_cron = payload.digest_cron.strip()
    if payload.include_excel is not None:
        r.include_excel = payload.include_excel
    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        raise HTTPException(status.HTTP_409_CONFLICT, "chat_id conflicts with an existing recipient.")
    db.refresh(r)
    scheduler.update_recipient_digest_job(r)
    return RecipientOut.model_validate(r)


@router.delete("/recipients/{recipient_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_recipient(recipient_id: int, db: Session = Depends(get_db)) -> Response:
    r = db.get(TelegramRecipient, recipient_id)
    if not r:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Recipient not found")
    scheduler.remove_recipient_digest_job(r.id)
    db.delete(r)
    db.commit()
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.post("/recipients/{recipient_id}/send/report", response_model=TelegramRunOut)
def send_report_to_recipient(recipient_id: int, db: Session = Depends(get_db)) -> TelegramRunOut:
    r = db.get(TelegramRecipient, recipient_id)
    if not r:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Recipient not found")
    if not r.is_active:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "Recipient is inactive")
    started = datetime.utcnow()
    run = TelegramRun(
        kind="outbound_report",
        status="ok",
        started_at=started,
        chat_id=r.chat_id,
        payload={"recipient_id": r.id, "recipient_name": r.name, "include_excel": r.include_excel},
    )
    db.add(run)
    db.flush()
    text = report_builder.build_digest_text(db)
    try:
        telegram_client.send_message(chat_id=r.chat_id, text=text, parse_mode="Markdown")
        if r.include_excel:
            filename, data = report_builder.build_entries_xlsx(db)
            telegram_client.send_document(
                chat_id=r.chat_id,
                filename=filename,
                data=data,
                caption=f"SettleOps digest — {datetime.utcnow().date().isoformat()}",
            )
    except telegram_client.TelegramError as e:
        run.status = "error"
        run.error_message = str(e)
    run.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(run)
    if run.status == "error":
        raise HTTPException(status.HTTP_502_BAD_GATEWAY, run.error_message or "Send failed")
    return TelegramRunOut.model_validate(run)
