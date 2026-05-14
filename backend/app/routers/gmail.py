"""/api/gmail — OAuth flow + manual sync + run history + config."""

from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query, Request, status
from fastapi.responses import RedirectResponse
from sqlalchemy.orm import Session

from app.config import settings
from app.db import get_db
from app.models import BankUpload, GmailRun, SourceUpload
from app.schemas.gmail import (
    GmailAuthStart,
    GmailConfig,
    GmailConfigUpdate,
    GmailMessageSummary,
    GmailRunOut,
    GmailSyncResult,
)
from app.services import excel_parser, gmail_fetcher, scheduler


router = APIRouter(prefix="/api/gmail", tags=["gmail"])


# ---------- Config ----------


@router.get("/config", response_model=GmailConfig)
def get_config() -> GmailConfig:
    return GmailConfig(
        search_query=settings.gmail_search_query,
        sync_interval_minutes=settings.gmail_sync_interval_minutes,
        has_credentials=gmail_fetcher.has_credentials_file(),
        has_token=gmail_fetcher.has_token_file(),
    )


@router.patch("/config", response_model=GmailConfig)
def update_config(payload: GmailConfigUpdate) -> GmailConfig:
    """In-memory config update. Persists for the lifetime of the process.

    For permanent changes, update the .env file. Scheduler is rescheduled on change.
    """
    if payload.search_query is not None:
        settings.gmail_search_query = payload.search_query
    if payload.sync_interval_minutes is not None:
        settings.gmail_sync_interval_minutes = payload.sync_interval_minutes
        scheduler.update_gmail_job(payload.sync_interval_minutes)
    return get_config()


# ---------- OAuth flow ----------


@router.get("/auth/start", response_model=GmailAuthStart)
def auth_start(request: Request) -> GmailAuthStart:
    redirect_uri = str(request.url_for("gmail_auth_callback"))
    try:
        flow = gmail_fetcher.build_auth_flow(redirect_uri)
    except FileNotFoundError as e:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, str(e)) from e
    auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
    return GmailAuthStart(auth_url=auth_url)


@router.get("/auth/callback", name="gmail_auth_callback")
def auth_callback(code: str, request: Request) -> RedirectResponse:
    redirect_uri = str(request.url_for("gmail_auth_callback"))
    flow = gmail_fetcher.build_auth_flow(redirect_uri)
    flow.fetch_token(code=code)
    gmail_fetcher.save_credentials(flow.credentials)
    # Send the user back to the frontend
    front = settings.cors_origins[0] if settings.cors_origins else "/"
    return RedirectResponse(url=f"{front}/?gmail=connected")


# ---------- Manual list / sync ----------


@router.get("/messages", response_model=List[GmailMessageSummary])
def list_messages(
    query: Optional[str] = Query(None),
    limit: int = Query(10, ge=1, le=50),
) -> List[GmailMessageSummary]:
    """Preview which Gmail messages would be picked up by the current search query."""
    q = query or settings.gmail_search_query
    try:
        ids = gmail_fetcher.search_messages(q, max_results=limit)
    except PermissionError as e:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, str(e)) from e
    except RuntimeError as e:
        raise HTTPException(status.HTTP_502_BAD_GATEWAY, str(e)) from e
    out: list[GmailMessageSummary] = []
    for mid in ids:
        meta = gmail_fetcher.fetch_message_meta(mid)
        out.append(
            GmailMessageSummary(
                id=meta.id,
                subject=meta.subject,
                sender=meta.sender,
                date=meta.date,
                snippet=meta.snippet,
                attachment_names=meta.attachment_names,
            )
        )
    return out


@router.post("/sync", response_model=GmailSyncResult)
def sync_now(db: Session = Depends(get_db)) -> GmailSyncResult:
    """Trigger a Gmail sync immediately. Records a GmailRun row."""
    return run_gmail_sync(db=db, triggered_by="manual")


@router.get("/runs", response_model=List[GmailRunOut])
def list_runs(limit: int = Query(20, ge=1, le=100), db: Session = Depends(get_db)) -> List[GmailRunOut]:
    rows = db.query(GmailRun).order_by(GmailRun.id.desc()).limit(limit).all()
    return [GmailRunOut.model_validate(r) for r in rows]


# ---------- Core sync routine ----------


def run_gmail_sync(db: Session, triggered_by: str) -> GmailSyncResult:
    """Search Gmail with the configured query, download Excel attachments, store as uploads.

    Heuristic: filenames are inspected — those that look like bank statements
    (contain 'bank', 'settle', 'mid') go to BankUpload; others go to SourceUpload.
    """
    started = datetime.utcnow()
    run = GmailRun(
        triggered_by=triggered_by,
        query=settings.gmail_search_query,
        status="ok",
        started_at=started,
        messages_seen=0,
        attachments_saved=0,
        uploads_created=0,
        payload={},
    )
    db.add(run)
    db.flush()

    saved_files: list[str] = []
    try:
        ids = gmail_fetcher.search_messages(settings.gmail_search_query, max_results=25)
        run.messages_seen = len(ids)
        for mid in ids:
            attachments = gmail_fetcher.download_attachments(
                mid, allowed_extensions={".xlsx", ".xls", ".csv"}
            )
            run.attachments_saved += len(attachments)
            for att in attachments:
                saved_files.append(att.filename)
                _store_attachment(db, att.filename, att.data, source="gmail")
                run.uploads_created += 1
    except PermissionError as e:
        run.status = "error"
        run.error_message = str(e)
    except Exception as e:  # noqa: BLE001
        run.status = "error"
        run.error_message = f"{type(e).__name__}: {e}"

    run.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(run)
    return GmailSyncResult(run=GmailRunOut.model_validate(run), saved_files=saved_files)


# ---------- Internal ----------


def _is_bank_filename(name: str) -> bool:
    name = name.lower()
    return any(k in name for k in ("bank", "settle", "mid"))


def _store_attachment(db: Session, filename: str, data: bytes, source: str) -> None:
    """Parse and store an Excel attachment as either a BankUpload or SourceUpload."""
    try:
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
    except ValueError as e:
        # Parse failure — skip this file but don't fail the whole sync
        # (caller's GmailRun will still record uploads_created for what worked)
        raise RuntimeError(f"Couldn't parse {filename}: {e}") from e
