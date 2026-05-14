"""Merchant + Company catalog (replaces the hard-coded MASTER_DATA in App.jsx)."""

from __future__ import annotations

from typing import List, Optional

from sqlalchemy import ForeignKey, Index, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.models.base import Base, TimestampMixin


class Merchant(Base, TimestampMixin):
    __tablename__ = "merchants"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(64), unique=True, nullable=False, index=True)

    companies: Mapped[List["Company"]] = relationship(
        back_populates="merchant",
        cascade="all, delete-orphan",
    )


class Company(Base, TimestampMixin):
    __tablename__ = "companies"
    __table_args__ = (
        Index("ix_companies_mid", "mid"),
        Index("ix_companies_gst", "gst_number"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    merchant_id: Mapped[int] = mapped_column(
        ForeignKey("merchants.id", ondelete="CASCADE"), nullable=False
    )
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    mid: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    gst_number: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    moa: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    incorporation: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)

    merchant: Mapped[Merchant] = relationship(back_populates="companies")
