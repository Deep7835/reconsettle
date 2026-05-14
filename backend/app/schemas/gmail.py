"""Gmail-related API schemas."""

from __future__ import annotations

from datetime import datetime
from typing import List, Literal, Optional

from pydantic import BaseModel, ConfigDict


class GmailConfig(BaseModel):
    search_query: str
    sync_interval_minutes: int
    has_credentials: bool
    has_token: bool


class GmailConfigUpdate(BaseModel):
    search_query: Optional[str] = None
    sync_interval_minutes: Optional[int] = None


class GmailAuthStart(BaseModel):
    auth_url: str


class GmailRunOut(BaseModel):
    id: int
    triggered_by: Literal["manual", "scheduler"]
    query: str
    status: Literal["ok", "error", "partial"]
    started_at: datetime
    finished_at: Optional[datetime]
    messages_seen: int
    attachments_saved: int
    uploads_created: int
    error_message: Optional[str]

    model_config = ConfigDict(from_attributes=True)


class GmailSyncResult(BaseModel):
    run: GmailRunOut
    saved_files: List[str]


class GmailMessageSummary(BaseModel):
    id: str
    subject: str
    sender: str
    date: Optional[str]
    snippet: Optional[str]
    attachment_names: List[str]


# ----- Dashboard combined response -----


class DashboardKpi(BaseModel):
    total_payin: float
    total_charge: float
    total_gst: float
    total_settlement: float
    entry_count: int
    recon_mismatch: int
    recon_total: int


class DashboardCycleSplit(BaseModel):
    c1_count: int
    c1_payin: float
    c2_count: int
    c2_payin: float


class DashboardReconStatus(BaseModel):
    matched: int
    mismatch: int
    pending: int
    resolved: int
    total_discrepancy: float


class DashboardCoverage(BaseModel):
    merchant_count: int
    company_count: int


class DashboardTopMerchant(BaseModel):
    merchant: str
    count: int
    payin: float
    settlement: float


class DashboardOpenIssue(BaseModel):
    id: int
    txn_id: str
    merchant: str
    company: str
    diff: float


class DashboardSnapshot(BaseModel):
    kpi: DashboardKpi
    today_entries: int
    today_payin: float
    today_date: str
    cycle_split: DashboardCycleSplit
    recon_status: DashboardReconStatus
    coverage: DashboardCoverage
    top_merchants: List[DashboardTopMerchant]
    open_recon_issues: List[DashboardOpenIssue]
    current_cycle: int
