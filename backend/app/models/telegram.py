"""Telegram run audit — one row per inbound poll or outbound send."""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from sqlalchemy import DateTime, Integer, JSON, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.models.base import Base, TimestampMixin


class TelegramRun(Base, TimestampMixin):
    __tablename__ = "telegram_runs"

    id: Mapped[int] = mapped_column(primary_key=True)
    # 'inbound_manual' | 'inbound_scheduler' | 'outbound_message' | 'outbound_report' | 'outbound_digest'
    kind: Mapped[str] = mapped_column(String(32), nullable=False)
    status: Mapped[str] = mapped_column(String(16), nullable=False)  # 'ok' | 'error'

    started_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    finished_at: Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True), nullable=True)

    chat_id: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    messages_seen: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    attachments_saved: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    uploads_created: Mapped[int] = mapped_column(Integer, nullable=False, default=0)

    error_message: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    payload: Mapped[dict] = mapped_column(JSON, nullable=False, default=dict)
