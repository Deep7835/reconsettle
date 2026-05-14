"""Settlement entries — one row per recorded settlement."""

from __future__ import annotations

from datetime import date as Date
from typing import Optional

from sqlalchemy import Date as SADate, Float, Index, Integer, String
from sqlalchemy.orm import Mapped, mapped_column

from app.models.base import Base, TimestampMixin


class Entry(Base, TimestampMixin):
    __tablename__ = "entries"
    __table_args__ = (
        Index("ix_entries_merchant", "merchant"),
        Index("ix_entries_company", "company"),
        Index("ix_entries_date", "entry_date"),
        Index("ix_entries_cycle", "cycle"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)

    merchant: Mapped[str] = mapped_column(String(64), nullable=False)
    company: Mapped[str] = mapped_column(String(255), nullable=False)

    payin: Mapped[float] = mapped_column(Float, nullable=False)
    chargeback: Mapped[float] = mapped_column(Float, nullable=False, default=0.0)
    charge: Mapped[float] = mapped_column(Float, nullable=False)
    gst: Mapped[float] = mapped_column(Float, nullable=False)
    gst_rate: Mapped[float] = mapped_column(Float, nullable=False, default=0.0)
    deduction: Mapped[float] = mapped_column(Float, nullable=False)
    settlement: Mapped[float] = mapped_column(Float, nullable=False)

    cycle: Mapped[int] = mapped_column(Integer, nullable=False)  # 1 or 2
    entry_date: Mapped[Date] = mapped_column(SADate, nullable=False)
    note: Mapped[Optional[str]] = mapped_column(String(500), nullable=True)
