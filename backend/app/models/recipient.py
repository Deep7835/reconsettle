"""Named Telegram recipients — contact book + per-user digest schedule."""

from __future__ import annotations

from typing import Optional

from sqlalchemy import Boolean, Index, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.models.base import Base, TimestampMixin


class TelegramRecipient(Base, TimestampMixin):
    __tablename__ = "telegram_recipients"
    __table_args__ = (
        Index("ix_recipients_chat_id", "chat_id", unique=True),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False)
    chat_id: Mapped[str] = mapped_column(String(64), nullable=False)

    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)

    # Per-recipient daily digest schedule (5-field crontab). Empty = no schedule.
    digest_cron: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    include_excel: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
