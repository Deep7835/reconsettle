"""Gmail integration: OAuth flow + attachment fetcher.

This module is the only part of the backend that talks to Google. It exposes:

  - build_auth_flow()          → create an OAuth flow for the FastAPI callback
  - get_service()              → return an authenticated Gmail API client
  - search_messages(query)     → list message IDs matching a Gmail search
  - fetch_message_meta(id)     → subject/from/date/attachments meta
  - download_attachments(id)   → bytes + filename for each attachment

The fetcher is intentionally storage-agnostic — the recon router decides what
to do with each attachment (typically: parse via excel_parser, save as
BankUpload/SourceUpload).
"""

from __future__ import annotations

import base64
import json
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from app.config import settings


GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


# ---------- OAuth helpers ----------


def has_credentials_file() -> bool:
    return Path(settings.gmail_client_secrets_file).exists()


def has_token_file() -> bool:
    return Path(settings.gmail_token_file).exists()


def build_auth_flow(redirect_uri: str) -> Flow:
    """Create an OAuth Flow object configured against client_secrets.json."""
    if not has_credentials_file():
        raise FileNotFoundError(
            f"Missing Google OAuth client secrets at {settings.gmail_client_secrets_file}. "
            "Download from Google Cloud Console → APIs & Services → Credentials → OAuth client."
        )
    return Flow.from_client_secrets_file(
        str(settings.gmail_client_secrets_file),
        scopes=GMAIL_SCOPES,
        redirect_uri=redirect_uri,
    )


def save_credentials(creds: Credentials) -> None:
    Path(settings.gmail_token_file).parent.mkdir(parents=True, exist_ok=True)
    Path(settings.gmail_token_file).write_text(creds.to_json())


def load_credentials() -> Optional[Credentials]:
    if not has_token_file():
        return None
    data = json.loads(Path(settings.gmail_token_file).read_text())
    creds = Credentials.from_authorized_user_info(data, GMAIL_SCOPES)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        save_credentials(creds)
    return creds


def get_service():
    """Return an authenticated Gmail API client, or raise if not authorized."""
    creds = load_credentials()
    if not creds:
        raise PermissionError(
            "Gmail not authorized yet. Hit POST /api/gmail/auth/start to begin OAuth."
        )
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


# ---------- Message + attachment helpers ----------


@dataclass
class GmailAttachment:
    message_id: str
    filename: str
    mime_type: str
    data: bytes


@dataclass
class GmailMessageMeta:
    id: str
    subject: str
    sender: str
    date: Optional[str]
    snippet: Optional[str]
    attachment_names: List[str]


def search_messages(query: str, max_results: int = 25) -> List[str]:
    """Return Gmail message IDs matching a Gmail-style search query."""
    service = get_service()
    try:
        resp = (
            service.users()
            .messages()
            .list(userId="me", q=query, maxResults=max_results)
            .execute()
        )
    except HttpError as e:
        raise RuntimeError(f"Gmail search failed: {e}") from e
    return [m["id"] for m in resp.get("messages", [])]


def fetch_message_meta(message_id: str) -> GmailMessageMeta:
    """Return subject / from / date / attachment filenames for one message."""
    service = get_service()
    msg = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="metadata", metadataHeaders=["Subject", "From", "Date"])
        .execute()
    )
    headers = {h["name"].lower(): h["value"] for h in msg.get("payload", {}).get("headers", [])}
    snippet = msg.get("snippet")
    # We need a second call (or full format) to see attachment filenames
    full = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )
    attachment_names = [
        part["filename"]
        for part in _walk_parts(full.get("payload", {}))
        if part.get("filename")
    ]
    return GmailMessageMeta(
        id=message_id,
        subject=headers.get("subject", ""),
        sender=headers.get("from", ""),
        date=headers.get("date"),
        snippet=snippet,
        attachment_names=attachment_names,
    )


def download_attachments(message_id: str, allowed_extensions: set[str] | None = None) -> List[GmailAttachment]:
    """Download every attachment of a message. Optionally filter by extension."""
    service = get_service()
    full = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )
    out: List[GmailAttachment] = []
    for part in _walk_parts(full.get("payload", {})):
        filename = part.get("filename")
        body = part.get("body", {})
        attachment_id = body.get("attachmentId")
        if not filename or not attachment_id:
            continue
        if allowed_extensions and not any(filename.lower().endswith(ext) for ext in allowed_extensions):
            continue
        att = (
            service.users()
            .messages()
            .attachments()
            .get(userId="me", messageId=message_id, id=attachment_id)
            .execute()
        )
        data = base64.urlsafe_b64decode(att["data"])
        out.append(
            GmailAttachment(
                message_id=message_id,
                filename=filename,
                mime_type=part.get("mimeType", "application/octet-stream"),
                data=data,
            )
        )
    return out


def _walk_parts(payload: dict):
    """Yield every leaf part (DFS) of a Gmail message payload."""
    yield payload
    for sub in payload.get("parts", []) or []:
        yield from _walk_parts(sub)
