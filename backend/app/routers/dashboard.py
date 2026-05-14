"""/api/dashboard — single endpoint returning everything the dashboard renders."""

from __future__ import annotations

from datetime import date as Date

from fastapi import APIRouter, Depends
from sqlalchemy import func
from sqlalchemy.orm import Session

from app.core.cycle import detect_cycle
from app.db import get_db
from app.models import Company, Entry, Merchant, ReconEntry
from app.schemas.gmail import (
    DashboardCoverage,
    DashboardCycleSplit,
    DashboardKpi,
    DashboardOpenIssue,
    DashboardReconStatus,
    DashboardSnapshot,
    DashboardTopMerchant,
)


router = APIRouter(prefix="/api/dashboard", tags=["dashboard"])


@router.get("/stats", response_model=DashboardSnapshot)
def dashboard_stats(db: Session = Depends(get_db)) -> DashboardSnapshot:
    today = Date.today()
    today_iso = today.isoformat()

    # KPI totals
    rows = db.query(Entry).all()
    kpi = DashboardKpi(
        total_payin=round(sum(r.payin for r in rows), 2),
        total_charge=round(sum(r.charge for r in rows), 2),
        total_gst=round(sum(r.gst for r in rows), 2),
        total_settlement=round(sum(r.settlement for r in rows), 2),
        entry_count=len(rows),
        recon_mismatch=db.query(ReconEntry)
        .filter(ReconEntry.status == "Mismatch", ReconEntry.resolved.is_(False))
        .count(),
        recon_total=db.query(ReconEntry).count(),
    )

    # Today
    today_rows = [r for r in rows if r.entry_date == today]
    today_entries = len(today_rows)
    today_payin = round(sum(r.payin for r in today_rows), 2)

    # Cycle split
    c1 = [r for r in rows if r.cycle == 1]
    c2 = [r for r in rows if r.cycle == 2]
    cycle_split = DashboardCycleSplit(
        c1_count=len(c1),
        c1_payin=round(sum(r.payin for r in c1), 2),
        c2_count=len(c2),
        c2_payin=round(sum(r.payin for r in c2), 2),
    )

    # Recon status
    recon_rows = db.query(ReconEntry).all()
    recon_status = DashboardReconStatus(
        matched=sum(1 for r in recon_rows if r.status == "Matched"),
        mismatch=sum(1 for r in recon_rows if r.status == "Mismatch" and not r.resolved),
        pending=sum(1 for r in recon_rows if r.status == "Pending"),
        resolved=sum(1 for r in recon_rows if r.resolved),
        total_discrepancy=round(
            sum(abs(r.diff) for r in recon_rows if r.status == "Mismatch" and not r.resolved), 2
        ),
    )

    # Coverage
    merchant_count = db.query(func.count(Merchant.id)).scalar() or 0
    company_count = db.query(func.count(Company.id)).scalar() or 0
    coverage = DashboardCoverage(
        merchant_count=int(merchant_count), company_count=int(company_count)
    )

    # Top merchants by settlement
    top: dict[str, dict] = {}
    for r in rows:
        b = top.setdefault(
            r.merchant, {"merchant": r.merchant, "count": 0, "payin": 0.0, "settlement": 0.0}
        )
        b["count"] += 1
        b["payin"] += r.payin
        b["settlement"] += r.settlement
    top_merchants = [
        DashboardTopMerchant(
            merchant=t["merchant"],
            count=t["count"],
            payin=round(t["payin"], 2),
            settlement=round(t["settlement"], 2),
        )
        for t in sorted(top.values(), key=lambda x: x["settlement"], reverse=True)[:5]
    ]

    # Open recon issues
    open_issues = (
        db.query(ReconEntry)
        .filter(ReconEntry.status == "Mismatch", ReconEntry.resolved.is_(False))
        .order_by(ReconEntry.id.desc())
        .limit(5)
        .all()
    )
    open_issues_out = [
        DashboardOpenIssue(
            id=r.id, txn_id=r.txn_id, merchant=r.merchant, company=r.company, diff=r.diff
        )
        for r in open_issues
    ]

    return DashboardSnapshot(
        kpi=kpi,
        today_entries=today_entries,
        today_payin=today_payin,
        today_date=today_iso,
        cycle_split=cycle_split,
        recon_status=recon_status,
        coverage=coverage,
        top_merchants=top_merchants,
        open_recon_issues=open_issues_out,
        current_cycle=detect_cycle().cycle,
    )
