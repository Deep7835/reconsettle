"""/api/entries — settlement entries CRUD + filters + totals + Excel export."""

from __future__ import annotations

import io
from datetime import date as Date
from typing import List, Optional

import pandas as pd
from fastapi import APIRouter, Depends, HTTPException, Query, Response, status
from sqlalchemy import and_
from sqlalchemy.orm import Session

from app.core.cycle import detect_cycle
from app.core.settle_math import compute_settlement
from app.db import get_db
from app.models import Entry
from app.schemas.entry import EntryCreate, EntryOut, EntryTotals, GroupedTotals


router = APIRouter(prefix="/api/entries", tags=["entries"])


# ---------- Helpers ----------


def _filter_query(
    db: Session,
    merchant: Optional[str],
    company: Optional[str],
    cycle: Optional[int],
    date_from: Optional[Date],
    date_to: Optional[Date],
):
    q = db.query(Entry)
    conditions = []
    if merchant and merchant != "__all__":
        conditions.append(Entry.merchant == merchant)
    if company and company != "__all__":
        conditions.append(Entry.company == company)
    if cycle:
        conditions.append(Entry.cycle == cycle)
    if date_from:
        conditions.append(Entry.entry_date >= date_from)
    if date_to:
        conditions.append(Entry.entry_date <= date_to)
    if conditions:
        q = q.filter(and_(*conditions))
    return q


# ---------- Endpoints ----------


@router.get("", response_model=List[EntryOut])
def list_entries(
    merchant: Optional[str] = None,
    company: Optional[str] = None,
    cycle: Optional[int] = Query(None, ge=1, le=2),
    date_from: Optional[Date] = None,
    date_to: Optional[Date] = None,
    limit: int = Query(500, ge=1, le=5000),
    db: Session = Depends(get_db),
) -> List[EntryOut]:
    q = _filter_query(db, merchant, company, cycle, date_from, date_to)
    rows = q.order_by(Entry.entry_date.desc(), Entry.id.desc()).limit(limit).all()
    return [EntryOut.model_validate(r) for r in rows]


@router.post("", response_model=EntryOut, status_code=status.HTTP_201_CREATED)
def create_entry(payload: EntryCreate, db: Session = Depends(get_db)) -> EntryOut:
    breakdown = compute_settlement(
        payin=payload.payin,
        chargeback=payload.chargeback,
        gst_rate=payload.gst_rate,
    )
    cycle = payload.cycle if payload.cycle is not None else detect_cycle().cycle
    entry_date = payload.entry_date or Date.today()

    e = Entry(
        merchant=payload.merchant,
        company=payload.company,
        payin=breakdown.payin,
        chargeback=breakdown.chargeback,
        charge=breakdown.charge,
        gst=breakdown.gst,
        gst_rate=breakdown.gst_rate,
        deduction=breakdown.deduction,
        settlement=breakdown.settlement,
        cycle=cycle,
        entry_date=entry_date,
        note=payload.note,
    )
    db.add(e)
    db.commit()
    db.refresh(e)
    return EntryOut.model_validate(e)


@router.delete("/{entry_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_entry(entry_id: int, db: Session = Depends(get_db)) -> Response:
    e = db.get(Entry, entry_id)
    if not e:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Entry not found")
    db.delete(e)
    db.commit()
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.delete("", status_code=status.HTTP_204_NO_CONTENT)
def clear_all_entries(db: Session = Depends(get_db)) -> Response:
    db.query(Entry).delete()
    db.commit()
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.get("/totals", response_model=EntryTotals)
def get_totals(
    merchant: Optional[str] = None,
    company: Optional[str] = None,
    cycle: Optional[int] = None,
    date_from: Optional[Date] = None,
    date_to: Optional[Date] = None,
    db: Session = Depends(get_db),
) -> EntryTotals:
    rows = _filter_query(db, merchant, company, cycle, date_from, date_to).all()
    return _sum_rows(rows)


@router.get("/grouped", response_model=List[GroupedTotals])
def get_grouped(
    by: str = Query("merchant", pattern="^(merchant|company)$"),
    merchant: Optional[str] = None,
    company: Optional[str] = None,
    date_from: Optional[Date] = None,
    date_to: Optional[Date] = None,
    db: Session = Depends(get_db),
) -> List[GroupedTotals]:
    rows = _filter_query(db, merchant, company, None, date_from, date_to).all()
    buckets: dict[str, list[Entry]] = {}
    for r in rows:
        key = r.merchant if by == "merchant" else r.company
        buckets.setdefault(key, []).append(r)
    out = [
        GroupedTotals(name=key, **_sum_rows(items).model_dump())
        for key, items in buckets.items()
    ]
    out.sort(key=lambda g: g.payin, reverse=True)
    return out


@router.get("/export.xlsx")
def export_xlsx(
    merchant: Optional[str] = None,
    company: Optional[str] = None,
    cycle: Optional[int] = None,
    date_from: Optional[Date] = None,
    date_to: Optional[Date] = None,
    db: Session = Depends(get_db),
) -> Response:
    rows = _filter_query(db, merchant, company, cycle, date_from, date_to).order_by(
        Entry.entry_date.desc(), Entry.id.desc()
    ).all()

    df = pd.DataFrame(
        [
            {
                "Date": r.entry_date,
                "Merchant": r.merchant,
                "Company": r.company,
                "Cycle": f"C{r.cycle}",
                "Payin": r.payin,
                "Charge": r.charge,
                "GST Rate": r.gst_rate,
                "GST": r.gst,
                "Chargeback": r.chargeback,
                "Deduction": r.deduction,
                "Settlement": r.settlement,
                "Note": r.note,
            }
            for r in rows
        ]
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Settlements")
    buf.seek(0)
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="settlements.xlsx"'},
    )


# ---------- Internal ----------


def _sum_rows(rows: list[Entry]) -> EntryTotals:
    return EntryTotals(
        count=len(rows),
        payin=round(sum(r.payin for r in rows), 2),
        charge=round(sum(r.charge for r in rows), 2),
        gst=round(sum(r.gst for r in rows), 2),
        chargeback=round(sum(r.chargeback for r in rows), 2),
        deduction=round(sum(r.deduction for r in rows), 2),
        settlement=round(sum(r.settlement for r in rows), 2),
    )
