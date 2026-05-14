"""Reconciliation records + parsed bank/source upload rows."""

from __future__ import annotations

from datetime import date as Date
from typing import Optional

from sqlalchemy import Boolean, Date as SADate, Float, Index, Integer, JSON, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.models.base import Base, TimestampMixin


class ReconEntry(Base, TimestampMixin):
    """Single-transaction reconciliation: our settlement vs bank claim."""

    __tablename__ = "recon_entries"
    __table_args__ = (
        Index("ix_recon_merchant", "merchant"),
        Index("ix_recon_status", "status"),
        Index("ix_recon_txn_id", "txn_id"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)

    merchant: Mapped[str] = mapped_column(String(64), nullable=False)
    company: Mapped[str] = mapped_column(String(255), nullable=False)
    txn_id: Mapped[str] = mapped_column(String(64), nullable=False)

    payin: Mapped[float] = mapped_column(Float, nullable=False)
    our_settle: Mapped[float] = mapped_column(Float, nullable=False)
    claimed: Mapped[float] = mapped_column(Float, nullable=False, default=0.0)
    diff: Mapped[float] = mapped_column(Float, nullable=False, default=0.0)

    utr: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    cycle: Mapped[int] = mapped_column(Integer, nullable=False, default=1)
    date: Mapped[Optional[Date]] = mapped_column(SADate, nullable=True)

    # 'Matched' | 'Mismatch' | 'Pending'
    status: Mapped[str] = mapped_column(String(16), nullable=False, default="Pending")
    reason: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    resolved: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    notes: Mapped[Optional[str]] = mapped_column(Text, nullable=True)


class BankUpload(Base, TimestampMixin):
    """A parsed bank settlement file (one row per uploaded sheet).

    `rows_json` holds the normalized parsed rows so we can re-aggregate later
    without re-parsing the original Excel.
    """

    __tablename__ = "bank_uploads"

    id: Mapped[int] = mapped_column(primary_key=True)
    filename: Mapped[str] = mapped_column(String(255), nullable=False)
    sha256: Mapped[Optional[str]] = mapped_column(String(64), nullable=True, index=True)
    source: Mapped[str] = mapped_column(String(32), nullable=False, default="manual")  # manual | gmail
    row_count: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    total_amount: Mapped[float] = mapped_column(Float, nullable=False, default=0.0)
    rows_json: Mapped[list] = mapped_column(JSON, nullable=False, default=list)
    notes: Mapped[Optional[str]] = mapped_column(Text, nullable=True)


class SourceUpload(Base, TimestampMixin):
    """A parsed internal payin file (one row per uploaded sheet)."""

    __tablename__ = "source_uploads"

    id: Mapped[int] = mapped_column(primary_key=True)
    filename: Mapped[str] = mapped_column(String(255), nullable=False)
    sha256: Mapped[Optional[str]] = mapped_column(String(64), nullable=True, index=True)
    source: Mapped[str] = mapped_column(String(32), nullable=False, default="manual")
    row_count: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    total_amount: Mapped[float] = mapped_column(Float, nullable=False, default=0.0)
    rows_json: Mapped[list] = mapped_column(JSON, nullable=False, default=list)
    notes: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
