"""/api/merchants — list / search / add merchants + companies."""

from __future__ import annotations

from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy import func, or_
from sqlalchemy.orm import Session, selectinload

from app.db import get_db
from app.models import Company, Merchant
from app.schemas.merchant import (
    CompanyCreate,
    CompanyOut,
    MerchantOut,
    MerchantSeed,
    MerchantSummary,
)


router = APIRouter(prefix="/api/merchants", tags=["merchants"])


@router.get("", response_model=List[MerchantOut])
def list_merchants(
    search: Optional[str] = Query(None, description="Match against company name / MID / GST"),
    merchant: Optional[str] = Query(None, description="Filter to a specific merchant"),
    include_companies: bool = Query(True),
    db: Session = Depends(get_db),
) -> List[MerchantOut]:
    q = db.query(Merchant).options(selectinload(Merchant.companies))
    if merchant:
        q = q.filter(Merchant.name == merchant)

    merchants = q.order_by(Merchant.name).all()

    if search:
        s = search.lower()
        filtered: list[Merchant] = []
        for m in merchants:
            matched = [
                c
                for c in m.companies
                if s in (c.name or "").lower()
                or s in (c.mid or "").lower()
                or s in (c.gst_number or "").lower()
            ]
            if matched:
                m.companies = matched  # type: ignore[assignment]
                filtered.append(m)
        merchants = filtered

    out: list[MerchantOut] = []
    for m in merchants:
        out.append(
            MerchantOut(
                id=m.id,
                name=m.name,
                company_count=len(m.companies),
                companies=[CompanyOut.model_validate(c) for c in m.companies] if include_companies else [],
            )
        )
    return out


@router.get("/summary", response_model=List[MerchantSummary])
def merchant_summary(db: Session = Depends(get_db)) -> List[MerchantSummary]:
    rows = (
        db.query(Merchant.name, func.count(Company.id))
        .outerjoin(Company, Company.merchant_id == Merchant.id)
        .group_by(Merchant.id)
        .order_by(Merchant.name)
        .all()
    )
    return [MerchantSummary(name=name, company_count=count) for name, count in rows]


@router.post("/companies", response_model=CompanyOut, status_code=status.HTTP_201_CREATED)
def add_company(payload: CompanyCreate, db: Session = Depends(get_db)) -> CompanyOut:
    """Add a company under a merchant. Auto-creates the merchant if new."""
    merchant = db.query(Merchant).filter(Merchant.name == payload.merchant_name).first()
    if not merchant:
        merchant = Merchant(name=payload.merchant_name)
        db.add(merchant)
        db.flush()

    company = Company(
        merchant_id=merchant.id,
        name=payload.name,
        mid=payload.mid,
        gst_number=payload.gst_number,
        moa=payload.moa,
        incorporation=payload.incorporation,
    )
    db.add(company)
    db.commit()
    db.refresh(company)
    return CompanyOut.model_validate(company)


@router.post("/seed", status_code=status.HTTP_201_CREATED)
def seed_merchants(rows: List[MerchantSeed], db: Session = Depends(get_db)) -> dict:
    """Bulk-seed merchants/companies from a JSON payload (used by the extractor script)."""
    seeded_merchants = 0
    seeded_companies = 0
    merchants_cache: dict[str, Merchant] = {m.name: m for m in db.query(Merchant).all()}

    for r in rows:
        m_name = r.merchant.strip()
        if not m_name:
            continue
        merchant = merchants_cache.get(m_name)
        if not merchant:
            merchant = Merchant(name=m_name)
            db.add(merchant)
            db.flush()
            merchants_cache[m_name] = merchant
            seeded_merchants += 1
        # Avoid dup companies for the same merchant
        existing = (
            db.query(Company)
            .filter(Company.merchant_id == merchant.id, Company.name == r.company)
            .first()
        )
        if existing:
            continue
        db.add(
            Company(
                merchant_id=merchant.id,
                name=r.company,
                mid=r.mid or None,
                gst_number=r.gst_number or None,
                moa=r.moa or None,
                incorporation=r.incorporation or None,
            )
        )
        seeded_companies += 1

    db.commit()
    return {"merchants_added": seeded_merchants, "companies_added": seeded_companies}
