"""Aggregation + auto-classification for reconciliation.

Given:
  - parsed source rows  (internal payin per company)
  - parsed bank rows    (bank settlement per company)
  - settlement entries  (our internal recorded settlements)

…produce the merchant-by-merchant aggregated view + suggest a status/reason
for each row using the same RULES the frontend was showing manually.
"""

from __future__ import annotations

from typing import Iterable, List

from sqlalchemy.orm import Session

from app.core.settle_math import compute_settlement
from app.models import Company, Entry, Merchant


# ---- Auto-classification rules for single-transaction recon ----

RECON_TOLERANCE = 1.0  # ₹1 rounding tolerance


def classify(payin: float, our_settle: float, claimed: float) -> tuple[str, str | None, float]:
    """Return (status, suggested_reason, diff)."""
    diff = round(our_settle - claimed, 2)
    if claimed <= 0:
        return "Pending", None, diff
    if abs(diff) <= RECON_TOLERANCE:
        return "Matched", "Rounding" if abs(diff) > 0 else None, diff
    # Heuristic reason guesses
    reason: str | None
    if diff < 0 and abs(diff) >= 0.10 * payin:
        reason = "TDS Deducted"
    elif diff < 0:
        reason = "Bank Charges"
    else:
        reason = "Missing Txn"
    return "Mismatch", reason, diff


def settlement_for(payin: float, gst_rate: float = 0.0, chargeback: float = 0.0) -> float:
    """Convenience wrapper used by routers."""
    return compute_settlement(payin=payin, gst_rate=gst_rate, chargeback=chargeback).settlement


# ---- Aggregated bank vs source vs internal view ----


def _build_company_to_merchant(db: Session) -> dict[str, str]:
    """Map lowercase company name → merchant name from the master data."""
    rows = (
        db.query(Company.name, Merchant.name)
        .join(Merchant, Merchant.id == Company.merchant_id)
        .all()
    )
    return {company.lower(): merchant for company, merchant in rows}


def aggregate(
    db: Session,
    source_rows: list[dict] | None,
    bank_rows: list[dict] | None,
) -> dict:
    """Aggregate parsed source + bank rows by merchant and compute diffs."""
    source_rows = source_rows or []
    bank_rows = bank_rows or []

    company_to_merchant = _build_company_to_merchant(db)

    agg: dict[str, dict] = {}
    unmatched_bank = 0

    # Source rows
    for r in source_rows:
        company = (r.get("company") or "").strip()
        if not company:
            continue
        merchant = company_to_merchant.get(company.lower(), "Unmatched")
        bucket = agg.setdefault(
            merchant,
            _new_bucket(merchant),
        )
        bucket["source_payin"] += float(r.get("amount") or 0.0)

    # Bank rows
    for r in bank_rows:
        company = (r.get("company") or "").strip()
        if not company:
            continue
        merchant = company_to_merchant.get(company.lower())
        if not merchant:
            merchant = "Unmatched"
            unmatched_bank += 1
        bucket = agg.setdefault(merchant, _new_bucket(merchant))
        bucket["amount"] += float(r.get("amount") or 0.0)
        bucket["fee"] += float(r.get("fee") or 0.0)
        bucket["gst"] += float(r.get("gst") or 0.0)
        bucket["settle"] += float(r.get("settle") or 0.0)
        bucket["chargeback"] += float(r.get("chargeback") or 0.0)
        if merchant == "Unmatched":
            bucket["unmatched"] += 1

    # Internal "our settlement" per merchant from Entry table (for direct comparison)
    internal_by_merchant = _internal_by_merchant(db)

    rows: List[dict] = []
    totals = _new_bucket("TOTAL")

    for merchant, b in sorted(agg.items()):
        our_net = internal_by_merchant.get(merchant, 0.0)
        net_settle = b["settle"] - b["chargeback"]
        b["net_settle"] = round(net_settle, 2)
        b["our_net"] = round(our_net, 2)
        b["diff"] = round(net_settle - our_net, 2)
        b["payin_diff"] = round(b["source_payin"] - b["amount"], 2)
        for key in (
            "source_payin",
            "amount",
            "fee",
            "gst",
            "settle",
            "chargeback",
            "net_settle",
            "our_net",
            "diff",
            "payin_diff",
        ):
            b[key] = round(b[key], 2)
            totals[key] += b[key]
        rows.append(b)

    totals["unmatched"] = unmatched_bank
    for k, v in totals.items():
        if isinstance(v, float):
            totals[k] = round(v, 2)

    return {"rows": rows, "totals": totals}


def _new_bucket(merchant: str) -> dict:
    return {
        "merchant": merchant,
        "source_payin": 0.0,
        "amount": 0.0,
        "fee": 0.0,
        "gst": 0.0,
        "settle": 0.0,
        "chargeback": 0.0,
        "net_settle": 0.0,
        "our_net": 0.0,
        "diff": 0.0,
        "payin_diff": 0.0,
        "unmatched": 0,
    }


def _internal_by_merchant(db: Session) -> dict[str, float]:
    """Sum of internal settlement entries grouped by merchant."""
    rows = db.query(Entry.merchant, Entry.settlement).all()
    out: dict[str, float] = {}
    for merchant, settlement in rows:
        out[merchant] = out.get(merchant, 0.0) + float(settlement or 0.0)
    return out
