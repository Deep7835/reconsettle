"""Settlement Entry API schemas."""

from __future__ import annotations

from datetime import date as Date
from typing import Optional

from pydantic import BaseModel, ConfigDict, Field


class EntryCreate(BaseModel):
    merchant: str
    company: str
    payin: float = Field(..., gt=0)
    chargeback: float = Field(0.0, ge=0)
    gst_rate: float = Field(0.0, ge=0, le=100)
    cycle: Optional[int] = Field(None, ge=1, le=2, description="Auto-detected if None")
    entry_date: Optional[Date] = None
    note: Optional[str] = None


class EntryOut(BaseModel):
    id: int
    merchant: str
    company: str
    payin: float
    chargeback: float
    charge: float
    gst: float
    gst_rate: float
    deduction: float
    settlement: float
    cycle: int
    entry_date: Date
    note: Optional[str] = None

    model_config = ConfigDict(from_attributes=True)


class EntryTotals(BaseModel):
    count: int
    payin: float
    charge: float
    gst: float
    chargeback: float
    deduction: float
    settlement: float


class GroupedTotals(BaseModel):
    name: str
    count: int
    payin: float
    charge: float
    gst: float
    chargeback: float
    deduction: float
    settlement: float
