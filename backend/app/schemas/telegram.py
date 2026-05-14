"""Telegram API schemas."""

from __future__ import annotations

from datetime import datetime
from typing import List, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field


class TelegramConfig(BaseModel):
    has_bot_token: bool
    report_chat_id: str
    inbound_chat_ids: List[str]
    sync_interval_minutes: int
    digest_cron: str


class TelegramConfigUpdate(BaseModel):
    bot_token: Optional[str] = None
    report_chat_id: Optional[str] = None
    inbound_chat_ids: Optional[List[str]] = None
    sync_interval_minutes: Optional[int] = None
    digest_cron: Optional[str] = None


class TelegramMe(BaseModel):
    id: int
    is_bot: bool
    first_name: str
    username: Optional[str] = None


class TelegramRunOut(BaseModel):
    id: int
    kind: str
    status: Literal["ok", "error"]
    started_at: datetime
    finished_at: Optional[datetime]
    chat_id: Optional[str]
    messages_seen: int
    attachments_saved: int
    uploads_created: int
    error_message: Optional[str]

    model_config = ConfigDict(from_attributes=True)


class TelegramSyncResult(BaseModel):
    run: TelegramRunOut
    saved_files: List[str]


class TelegramSendMessage(BaseModel):
    chat_id: Optional[str] = Field(
        None,
        description="Override target chat. Defaults to TELEGRAM_REPORT_CHAT_ID.",
    )
    text: str = Field(..., min_length=1, max_length=4000)
    parse_mode: Literal["Markdown", "MarkdownV2", "HTML", "none"] = "Markdown"


class TelegramSendReport(BaseModel):
    chat_id: Optional[str] = None
    include_excel: bool = True
    caption: Optional[str] = None


# ----- Recipient contact book -----


class RecipientCreate(BaseModel):
    name: str = Field(..., min_length=1, max_length=64)
    chat_id: str = Field(..., min_length=1, max_length=64)
    description: Optional[str] = None
    is_active: bool = True
    digest_cron: str = ""
    include_excel: bool = True


class RecipientUpdate(BaseModel):
    name: Optional[str] = None
    chat_id: Optional[str] = None
    description: Optional[str] = None
    is_active: Optional[bool] = None
    digest_cron: Optional[str] = None
    include_excel: Optional[bool] = None


class RecipientOut(BaseModel):
    id: int
    name: str
    chat_id: str
    description: Optional[str]
    is_active: bool
    digest_cron: str
    include_excel: bool

    model_config = ConfigDict(from_attributes=True)
