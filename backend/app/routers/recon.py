"""/api/recon — single-transaction reconciliation + bulk Excel upload + aggregated view."""

from __future__ import annotations

from typing import List, Optional

from fastapi import APIRouter, Depends, File, HTTPException, Query, Response, UploadFile, status
from sqlalchemy.orm import Session

from app.db import get_db
from app.models import BankUpload, ReconEntry, SourceUpload
from app.schemas.recon import (
    AggregatedReport,
    AggregatedRow,
    ReconCreate,
    ReconOut,
    ReconStats,
    ReconUpdate,
    UploadInfo,
)
from app.services import excel_parser, recon_engine


router = APIRouter(prefix="/api/recon", tags=["recon"])


# ---------- Single-transaction recon ----------


@router.get("", response_model=List[ReconOut])
def list_recon(
    status_filter: Optional[str] = Query(None, alias="status"),
    merchant: Optional[str] = None,
    db: Session = Depends(get_db),
) -> List[ReconOut]:
    q = db.query(ReconEntry)
    if status_filter and status_filter != "all":
        q = q.filter(ReconEntry.status == status_filter)
    if merchant and merchant != "__all__":
        q = q.filter(ReconEntry.merchant == merchant)
    rows = q.order_by(ReconEntry.id.desc()).limit(1000).all()
    return [ReconOut.model_validate(r) for r in rows]


@router.post("", response_model=ReconOut, status_code=status.HTTP_201_CREATED)
def create_recon(payload: ReconCreate, db: Session = Depends(get_db)) -> ReconOut:
    our_settle = recon_engine.settlement_for(
        payin=payload.payin, gst_rate=payload.gst_rate, chargeback=payload.chargeback
    )
    rstatus, reason, diff = recon_engine.classify(
        payin=payload.payin, our_settle=our_settle, claimed=payload.claimed
    )
    r = ReconEntry(
        merchant=payload.merchant,
        company=payload.company,
        txn_id=payload.txn_id,
        payin=payload.payin,
        our_settle=our_settle,
        claimed=payload.claimed,
        diff=diff,
        utr=payload.utr,
        cycle=payload.cycle,
        date=payload.date,
        status=rstatus,
        reason=reason,
        resolved=False,
    )
    db.add(r)
    db.commit()
    db.refresh(r)
    return ReconOut.model_validate(r)


@router.patch("/{recon_id}", response_model=ReconOut)
def update_recon(recon_id: int, payload: ReconUpdate, db: Session = Depends(get_db)) -> ReconOut:
    r = db.get(ReconEntry, recon_id)
    if not r:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Recon entry not found")
    if payload.reason is not None:
        r.reason = payload.reason
    if payload.notes is not None:
        r.notes = payload.notes
    if payload.resolved is not None:
        r.resolved = payload.resolved
    db.commit()
    db.refresh(r)
    return ReconOut.model_validate(r)


@router.delete("/{recon_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_recon(recon_id: int, db: Session = Depends(get_db)) -> Response:
    r = db.get(ReconEntry, recon_id)
    if not r:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Recon entry not found")
    db.delete(r)
    db.commit()
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.delete("", status_code=status.HTTP_204_NO_CONTENT)
def clear_all_recon(db: Session = Depends(get_db)) -> Response:
    db.query(ReconEntry).delete()
    db.commit()
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.get("/stats", response_model=ReconStats)
def recon_stats(db: Session = Depends(get_db)) -> ReconStats:
    rows = db.query(ReconEntry).all()
    matched = sum(1 for r in rows if r.status == "Matched")
    mismatch = sum(1 for r in rows if r.status == "Mismatch" and not r.resolved)
    pending = sum(1 for r in rows if r.status == "Pending")
    resolved = sum(1 for r in rows if r.resolved)
    total_disc = sum(abs(r.diff) for r in rows if r.status == "Mismatch" and not r.resolved)
    return ReconStats(
        total=len(rows),
        matched=matched,
        mismatch=mismatch,
        pending=pending,
        resolved=resolved,
        total_discrepancy=round(total_disc, 2),
    )


# ---------- Excel uploads ----------


@router.post("/upload/bank", response_model=UploadInfo, status_code=status.HTTP_201_CREATED)
async def upload_bank(file: UploadFile = File(...), db: Session = Depends(get_db)) -> UploadInfo:
    raw = await file.read()
    try:
        parsed = excel_parser.parse_bank_excel(raw)
    except ValueError as e:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, str(e)) from e
    upload = BankUpload(
        filename=file.filename or "bank.xlsx",
        sha256=excel_parser.file_sha256(raw),
        source="manual",
        row_count=len(parsed.rows),
        total_amount=round(parsed.total_amount, 2),
        rows_json=parsed.rows,
    )
    db.add(upload)
    db.commit()
    db.refresh(upload)
    return UploadInfo(
        id=upload.id,
        filename=upload.filename,
        source=upload.source,
        row_count=upload.row_count,
        total_amount=upload.total_amount,
    )


@router.post("/upload/source", response_model=UploadInfo, status_code=status.HTTP_201_CREATED)
async def upload_source(file: UploadFile = File(...), db: Session = Depends(get_db)) -> UploadInfo:
    raw = await file.read()
    try:
        parsed = excel_parser.parse_source_excel(raw)
    except ValueError as e:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, str(e)) from e
    upload = SourceUpload(
        filename=file.filename or "source.xlsx",
        sha256=excel_parser.file_sha256(raw),
        source="manual",
        row_count=len(parsed.rows),
        total_amount=round(parsed.total_amount, 2),
        rows_json=parsed.rows,
    )
    db.add(upload)
    db.commit()
    db.refresh(upload)
    return UploadInfo(
        id=upload.id,
        filename=upload.filename,
        source=upload.source,
        row_count=upload.row_count,
        total_amount=upload.total_amount,
    )


@router.get("/uploads/bank", response_model=List[UploadInfo])
def list_bank_uploads(db: Session = Depends(get_db)) -> List[UploadInfo]:
    rows = db.query(BankUpload).order_by(BankUpload.id.desc()).limit(50).all()
    return [
        UploadInfo(
            id=u.id, filename=u.filename, source=u.source, row_count=u.row_count, total_amount=u.total_amount
        )
        for u in rows
    ]


@router.get("/uploads/source", response_model=List[UploadInfo])
def list_source_uploads(db: Session = Depends(get_db)) -> List[UploadInfo]:
    rows = db.query(SourceUpload).order_by(SourceUpload.id.desc()).limit(50).all()
    return [
        UploadInfo(
            id=u.id, filename=u.filename, source=u.source, row_count=u.row_count, total_amount=u.total_amount
        )
        for u in rows
    ]


@router.get("/aggregated", response_model=AggregatedReport)
def get_aggregated(
    bank_upload_id: Optional[int] = None,
    source_upload_id: Optional[int] = None,
    db: Session = Depends(get_db),
) -> AggregatedReport:
    """Aggregate latest (or specific) uploads by merchant and compute diffs."""
    bank = (
        db.get(BankUpload, bank_upload_id)
        if bank_upload_id
        else db.query(BankUpload).order_by(BankUpload.id.desc()).first()
    )
    source = (
        db.get(SourceUpload, source_upload_id)
        if source_upload_id
        else db.query(SourceUpload).order_by(SourceUpload.id.desc()).first()
    )

    result = recon_engine.aggregate(
        db=db,
        source_rows=source.rows_json if source else [],
        bank_rows=bank.rows_json if bank else [],
    )

    rows = [AggregatedRow(**r) for r in result["rows"]]
    totals = AggregatedRow(**result["totals"])
    return AggregatedReport(
        rows=rows,
        totals=totals,
        source_upload=UploadInfo(
            id=source.id,
            filename=source.filename,
            source=source.source,
            row_count=source.row_count,
            total_amount=source.total_amount,
        )
        if source
        else None,
        bank_upload=UploadInfo(
            id=bank.id,
            filename=bank.filename,
            source=bank.source,
            row_count=bank.row_count,
            total_amount=bank.total_amount,
        )
        if bank
        else None,
    )
