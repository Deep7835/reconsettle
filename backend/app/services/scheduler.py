"""APScheduler — runs Gmail sync, Telegram inbound polling, and an optional
daily Telegram digest, all inside the FastAPI process.
"""

from __future__ import annotations

import logging
from typing import Optional

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from app.config import settings


logger = logging.getLogger("settleops.scheduler")

_scheduler: Optional[BackgroundScheduler] = None

_GMAIL_JOB_ID = "gmail_sync"
_TELEGRAM_INBOUND_JOB_ID = "telegram_inbound"
_TELEGRAM_DIGEST_JOB_ID = "telegram_digest"


def get_scheduler() -> BackgroundScheduler:
    global _scheduler
    if _scheduler is None:
        _scheduler = BackgroundScheduler(timezone="UTC")
    return _scheduler


def start() -> None:
    """Start the scheduler and register all jobs based on current settings."""
    s = get_scheduler()
    if not s.running:
        s.start()
    # Lazy imports so we don't pull in optional deps at module load
    from app.services import gmail_fetcher, telegram_inbound  # noqa: F401

    update_gmail_job(settings.gmail_sync_interval_minutes)
    update_telegram_job(settings.telegram_sync_interval_minutes)
    update_telegram_digest(settings.telegram_digest_cron)
    load_recipient_digest_jobs()


def shutdown() -> None:
    s = get_scheduler()
    if s.running:
        s.shutdown(wait=False)


# ---------- Gmail ----------


def update_gmail_job(interval_minutes: int) -> None:
    s = get_scheduler()
    if s.get_job(_GMAIL_JOB_ID):
        s.remove_job(_GMAIL_JOB_ID)
    if interval_minutes <= 0:
        logger.info("Gmail sync disabled (interval <= 0)")
        return
    s.add_job(
        _run_gmail_sync,
        "interval",
        minutes=interval_minutes,
        id=_GMAIL_JOB_ID,
        max_instances=1,
        coalesce=True,
    )
    logger.info("Scheduled Gmail sync every %s minutes", interval_minutes)


def _run_gmail_sync() -> None:
    from app.db import SessionLocal
    from app.routers.gmail import run_gmail_sync  # type: ignore

    db = SessionLocal()
    try:
        try:
            run_gmail_sync(db=db, triggered_by="scheduler")
            logger.info("Scheduled Gmail sync completed")
        except PermissionError:
            logger.info("Skipping Gmail sync: not authorized yet")
        except Exception as e:  # noqa: BLE001
            logger.exception("Scheduled Gmail sync failed: %s", e)
    finally:
        db.close()


# ---------- Telegram inbound ----------


def update_telegram_job(interval_minutes: int) -> None:
    s = get_scheduler()
    if s.get_job(_TELEGRAM_INBOUND_JOB_ID):
        s.remove_job(_TELEGRAM_INBOUND_JOB_ID)
    if interval_minutes <= 0:
        logger.info("Telegram inbound disabled (interval <= 0)")
        return
    if not settings.telegram_bot_token:
        logger.info("Telegram inbound skipped: no bot token configured yet")
        return
    s.add_job(
        _run_telegram_inbound,
        "interval",
        minutes=interval_minutes,
        id=_TELEGRAM_INBOUND_JOB_ID,
        max_instances=1,
        coalesce=True,
    )
    logger.info("Scheduled Telegram inbound poll every %s minutes", interval_minutes)


def _run_telegram_inbound() -> None:
    from app.db import SessionLocal
    from app.services import telegram_inbound

    db = SessionLocal()
    try:
        try:
            telegram_inbound.sync(db=db, triggered_by="scheduler")
        except Exception as e:  # noqa: BLE001
            logger.exception("Scheduled Telegram inbound failed: %s", e)
    finally:
        db.close()


# ---------- Telegram digest ----------


def update_telegram_digest(cron_expr: str) -> None:
    s = get_scheduler()
    if s.get_job(_TELEGRAM_DIGEST_JOB_ID):
        s.remove_job(_TELEGRAM_DIGEST_JOB_ID)
    if not cron_expr.strip():
        logger.info("Telegram digest disabled (no cron expression)")
        return
    if not settings.telegram_bot_token or not settings.telegram_report_chat_id:
        logger.info("Telegram digest skipped: bot token or report chat not configured")
        return
    try:
        trigger = CronTrigger.from_crontab(cron_expr.strip())
    except ValueError as e:
        logger.warning("Invalid TELEGRAM_DIGEST_CRON %r: %s", cron_expr, e)
        return
    s.add_job(
        _run_telegram_digest,
        trigger=trigger,
        id=_TELEGRAM_DIGEST_JOB_ID,
        max_instances=1,
        coalesce=True,
    )
    logger.info("Scheduled Telegram digest with cron '%s'", cron_expr)


def _run_telegram_digest() -> None:
    from datetime import datetime

    from app.db import SessionLocal
    from app.models import TelegramRun
    from app.services import report_builder, telegram_client

    db = SessionLocal()
    try:
        started = datetime.utcnow()
        run = TelegramRun(
            kind="outbound_digest",
            status="ok",
            started_at=started,
            chat_id=settings.telegram_report_chat_id,
            payload={"cron": settings.telegram_digest_cron},
        )
        db.add(run)
        db.flush()
        try:
            text = report_builder.build_digest_text(db)
            telegram_client.send_message(
                chat_id=settings.telegram_report_chat_id,
                text=text,
                parse_mode="Markdown",
            )
            filename, data = report_builder.build_entries_xlsx(db)
            telegram_client.send_document(
                chat_id=settings.telegram_report_chat_id,
                filename=filename,
                data=data,
                caption=f"SettleOps digest — {datetime.utcnow().date().isoformat()}",
            )
        except Exception as e:  # noqa: BLE001
            run.status = "error"
            run.error_message = str(e)
            logger.exception("Telegram digest failed: %s", e)
        run.finished_at = datetime.utcnow()
        db.commit()
    finally:
        db.close()


# ---------- Per-recipient digest jobs ----------


def _recipient_job_id(recipient_id: int) -> str:
    return f"telegram_recipient_digest_{recipient_id}"


def load_recipient_digest_jobs() -> None:
    """On startup: register a cron job for each active recipient that has digest_cron set."""
    from app.db import SessionLocal
    from app.models import TelegramRecipient

    db = SessionLocal()
    try:
        rows = (
            db.query(TelegramRecipient)
            .filter(TelegramRecipient.is_active.is_(True))
            .filter(TelegramRecipient.digest_cron != "")
            .all()
        )
        for r in rows:
            update_recipient_digest_job(r)
    finally:
        db.close()


def update_recipient_digest_job(recipient) -> None:
    """Re-register (or remove) the cron job for a single recipient.

    Called from CRUD endpoints right after a recipient is created/updated/deleted.
    """
    s = get_scheduler()
    job_id = _recipient_job_id(recipient.id)
    if s.get_job(job_id):
        s.remove_job(job_id)

    if not recipient.is_active or not recipient.digest_cron.strip():
        return
    if not settings.telegram_bot_token:
        logger.info("Recipient %s digest skipped: no bot token configured", recipient.id)
        return
    try:
        trigger = CronTrigger.from_crontab(recipient.digest_cron.strip())
    except ValueError as e:
        logger.warning("Invalid cron for recipient %s (%r): %s", recipient.id, recipient.digest_cron, e)
        return
    s.add_job(
        _run_recipient_digest,
        trigger=trigger,
        id=job_id,
        max_instances=1,
        coalesce=True,
        args=[recipient.id],
    )
    logger.info(
        "Scheduled digest for recipient %s (%s) with cron '%s'",
        recipient.id, recipient.name, recipient.digest_cron,
    )


def remove_recipient_digest_job(recipient_id: int) -> None:
    s = get_scheduler()
    job_id = _recipient_job_id(recipient_id)
    if s.get_job(job_id):
        s.remove_job(job_id)


def _run_recipient_digest(recipient_id: int) -> None:
    """Background job — sends digest text + Excel to one specific recipient."""
    from datetime import datetime

    from app.db import SessionLocal
    from app.models import TelegramRecipient, TelegramRun
    from app.services import report_builder, telegram_client

    db = SessionLocal()
    try:
        r = db.get(TelegramRecipient, recipient_id)
        if not r or not r.is_active:
            logger.info("Recipient %s missing/inactive; skipping digest", recipient_id)
            return
        started = datetime.utcnow()
        run = TelegramRun(
            kind="outbound_digest",
            status="ok",
            started_at=started,
            chat_id=r.chat_id,
            payload={"recipient_id": r.id, "recipient_name": r.name, "cron": r.digest_cron},
        )
        db.add(run)
        db.flush()
        try:
            text = report_builder.build_digest_text(db)
            telegram_client.send_message(chat_id=r.chat_id, text=text, parse_mode="Markdown")
            if r.include_excel:
                filename, data = report_builder.build_entries_xlsx(db)
                telegram_client.send_document(
                    chat_id=r.chat_id,
                    filename=filename,
                    data=data,
                    caption=f"SettleOps digest — {datetime.utcnow().date().isoformat()}",
                )
        except Exception as e:  # noqa: BLE001
            run.status = "error"
            run.error_message = str(e)
            logger.exception("Recipient %s digest failed: %s", r.id, e)
        run.finished_at = datetime.utcnow()
        db.commit()
    finally:
        db.close()
