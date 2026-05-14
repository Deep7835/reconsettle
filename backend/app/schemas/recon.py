"""Reconciliation API schemas."""

from __future__ import annotations

from datetime import date as Date
from typing import List, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field


# ----- Single-transaction recon -----


class ReconCreate(BaseModel):
    merchant: str
    company: str
    txn_id: str
    payin: float = Field(..., gt=0)
    claimed: float = Field(0.0, ge=0)
    gst_rate: float = Field(0.0, ge=0, le=100, description="GST % used to compute our settlement")
    chargeback: float = Field(0.0, ge=0)
    utr: Optional[str] = None
    cycle: int = Field(1, ge=1, le=2)
    date: Optional[Date] = None


class ReconUpdate(BaseModel):
    reason: Optional[str] = None
    resolved: Optional[bool] = None
    notes: Optional[str] = None


class ReconOut(BaseModel):
    id: int
    merchant: str
    company: str
    txn_id: str
    payin: float
    our_settle: float
    claimed: float
    diff: float
    utr: Optional[str]
    cycle: int
    date: Optional[Date]
    status: Literal["Matched", "Mismatch", "Pending"]
    reason: Optional[str]
    resolved: bool
    notes: Optional[str]

    model_config = ConfigDict(from_attributes=True)


class ReconStats(BaseModel):
    total: int
    matched: int
    mismatch: int
    pending: int
    resolved: int
    total_discrepancy: float


# ----- File uploads / aggregated views -----


class UploadInfo(BaseModel):
    id: int
    filename: str
    source: str
    row_count: int
    total_amount: float


class BankRow(BaseModel):
    mid: Optional[str] = None
    company: str
    amount: float = 0.0
    fee: float = 0.0
    gst: float = 0.0
    settle: float = 0.0
    chargeback: float = 0.0
    matched_merchant: Optional[str] = None


class SourceRow(BaseModel):
    company: str
    amount: float = 0.0
    matched_merchant: Optional[str] = None


class AggregatedRow(BaseModel):
    merchant: str
    source_payin: float = 0.0
    amount: float = 0.0
    fee: float = 0.0
    gst: float = 0.0
    settle: float = 0.0
    chargeback: float = 0.0
    net_settle: float = 0.0
    our_net: float = 0.0
    diff: float = 0.0
    payin_diff: float = 0.0
    unmatched: int = 0


class AggregatedReport(BaseModel):
    rows: List[AggregatedRow]
    totals: AggregatedRow
    source_upload: Optional[UploadInfo] = None
    bank_upload: Optional[UploadInfo] = None
