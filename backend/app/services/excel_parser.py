"""Header-driven Excel parsers for bank settlement + internal payin files.

Mirrors the heuristics the React UI was using:

  Bank file expected columns:
      MID, Name (or Beneficiary), Amount (or Credit), Fee,
      GST (CGST+SGST or IGST), Settle (Net), Chargeback (negative entries)

  Source / internal payin file expected columns:
      Row Labels (or Name / Beneficiary), Sum of Amount (or Amount)

Headers are matched case-insensitively against a list of synonyms — that way
slightly different bank templates still parse.
"""

from __future__ import annotations

import hashlib
import io
from dataclasses import dataclass
from typing import Iterable, List

import pandas as pd


# ---------- Column-name resolution ----------

BANK_HEADERS = {
    "mid": ["mid", "merchant id", "merchantid"],
    "company": ["name", "beneficiary", "company", "merchant name", "party name"],
    "amount": ["amount", "credit", "gross", "txn amount", "transaction amount"],
    "fee": ["fee", "commission", "charge", "mdr"],
    "gst": ["gst", "tax", "igst", "cgst+sgst", "cgst", "sgst"],
    "settle": ["settle", "net", "net settle", "settlement", "net amount"],
    "chargeback": ["chargeback", "refund", "reversal", "debit"],
}

SOURCE_HEADERS = {
    "company": ["row labels", "name", "beneficiary", "company", "merchant name"],
    "amount": ["sum of amount", "amount", "total amount", "total"],
}


def _resolve_columns(df: pd.DataFrame, mapping: dict[str, list[str]]) -> dict[str, str]:
    """Return {logical_name: actual_column} by matching synonyms (case-insensitive)."""
    lookup = {str(col).strip().lower(): col for col in df.columns}
    resolved: dict[str, str] = {}
    for logical, synonyms in mapping.items():
        for syn in synonyms:
            if syn in lookup:
                resolved[logical] = lookup[syn]
                break
    return resolved


# ---------- Public dataclasses ----------


@dataclass
class ParsedBank:
    rows: list[dict]
    total_amount: float


@dataclass
class ParsedSource:
    rows: list[dict]
    total_amount: float


# ---------- Bank parser ----------


def parse_bank_excel(raw: bytes) -> ParsedBank:
    """Parse a bank settlement Excel/CSV into normalized rows grouped by company."""
    df = _read_any(raw)
    cols = _resolve_columns(df, BANK_HEADERS)

    if "company" not in cols:
        raise ValueError(
            "Bank file missing a 'Name' / 'Beneficiary' / 'Company' column. "
            "Required to identify which company each row belongs to."
        )

    grouped: dict[str, dict] = {}
    for _, r in df.iterrows():
        company = _clean_str(r.get(cols["company"]))
        if not company:
            continue
        bucket = grouped.setdefault(
            company,
            {
                "mid": _clean_str(r.get(cols.get("mid"))) if "mid" in cols else None,
                "company": company,
                "amount": 0.0,
                "fee": 0.0,
                "gst": 0.0,
                "settle": 0.0,
                "chargeback": 0.0,
            },
        )
        bucket["amount"] += _to_float(r.get(cols.get("amount")))
        bucket["fee"] += _to_float(r.get(cols.get("fee")))
        bucket["gst"] += _to_float(r.get(cols.get("gst")))
        bucket["settle"] += _to_float(r.get(cols.get("settle")))
        cb = _to_float(r.get(cols.get("chargeback")))
        # Bank files usually show chargebacks as negative — make them positive here
        if cb < 0:
            bucket["chargeback"] += abs(cb)
        else:
            bucket["chargeback"] += cb

    rows = list(grouped.values())
    return ParsedBank(rows=rows, total_amount=sum(r["amount"] for r in rows))


# ---------- Source / internal payin parser ----------


def parse_source_excel(raw: bytes) -> ParsedSource:
    """Parse an internal payin Excel/CSV into normalized rows grouped by company."""
    df = _read_any(raw)
    cols = _resolve_columns(df, SOURCE_HEADERS)

    if "company" not in cols or "amount" not in cols:
        raise ValueError(
            "Source file missing 'Row Labels' / 'Name' or 'Sum of Amount' / 'Amount' column."
        )

    grouped: dict[str, float] = {}
    for _, r in df.iterrows():
        company = _clean_str(r.get(cols["company"]))
        amount = _to_float(r.get(cols["amount"]))
        if not company or company.lower() in {"grand total", "total"}:
            continue
        grouped[company] = grouped.get(company, 0.0) + amount

    rows = [{"company": name, "amount": amt} for name, amt in grouped.items()]
    return ParsedSource(rows=rows, total_amount=sum(r["amount"] for r in rows))


# ---------- Helpers ----------


def _read_any(raw: bytes) -> pd.DataFrame:
    """Read either Excel or CSV based on content sniffing."""
    bio = io.BytesIO(raw)
    # Try Excel first
    try:
        return pd.read_excel(bio, dtype=str)
    except Exception:
        bio.seek(0)
        return pd.read_csv(bio, dtype=str)


def _to_float(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if not pd.isna(v) else 0.0
    s = str(v).replace(",", "").replace("₹", "").strip()
    if not s or s.lower() in {"nan", "none"}:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _clean_str(v) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    return s


def file_sha256(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()
