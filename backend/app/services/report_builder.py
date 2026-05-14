"""Build human-readable daily reports + Excel exports for Telegram delivery.

Two shapes:
  - `build_digest_text()`  → short Markdown summary suitable for a Telegram message
  - `build_entries_xlsx()` → bytes of a sortable Excel workbook of today's entries
"""

from __future__ import annotations

import io
from datetime import date as Date
from typing import Optional

import pandas as pd
from sqlalchemy.orm import Session

from app.core.cycle import detect_cycle
from app.models import Entry, ReconEntry


def _fmt_inr(n: float) -> str:
    return "INR " + f"{n:,.2f}"


def _fmt_short(n: float) -> str:
    n = float(n or 0.0)
    if n >= 1e7:
        return f"INR {n/1e7:.2f} Cr"
    if n >= 1e5:
        return f"INR {n/1e5:.2f} L"
    if n >= 1e3:
        return f"INR {n/1e3:.1f}K"
    return _fmt_inr(n)


def build_digest_text(db: Session, for_date: Optional[Date] = None) -> str:
    """Return a Markdown summary covering totals + today + open recon issues."""
    for_date = for_date or Date.today()
    cycle = detect_cycle()

    rows = db.query(Entry).all()
    today_rows = [r for r in rows if r.entry_date == for_date]
    total_payin = sum(r.payin for r in rows)
    total_settle = sum(r.settlement for r in rows)
    today_payin = sum(r.payin for r in today_rows)
    today_settle = sum(r.settlement for r in today_rows)

    c1 = sum(1 for r in today_rows if r.cycle == 1)
    c2 = sum(1 for r in today_rows if r.cycle == 2)

    open_issues = (
        db.query(ReconEntry)
        .filter(ReconEntry.status == "Mismatch", ReconEntry.resolved.is_(False))
        .order_by(ReconEntry.id.desc())
        .limit(5)
        .all()
    )
    total_disc = sum(abs(r.diff) for r in open_issues)

    # Telegram Markdown is a bit fussy — keep it simple, no nested formatting.
    lines = [
        f"*SettleOps Digest — {for_date.isoformat()}*",
        f"_Active cycle: C{cycle.cycle} ({cycle.label})_",
        "",
        "*Today*",
        f"  Entries: *{len(today_rows)}*    C1: {c1}   C2: {c2}",
        f"  Payin: *{_fmt_short(today_payin)}*    Settlement: *{_fmt_short(today_settle)}*",
        "",
        "*All-time*",
        f"  Total entries: *{len(rows)}*",
        f"  Total payin: *{_fmt_short(total_payin)}*",
        f"  Total settlement: *{_fmt_short(total_settle)}*",
        "",
        f"*Open recon issues:* {len(open_issues)}   Discrepancy: {_fmt_short(total_disc)}",
    ]
    if open_issues:
        lines.append("")
        for i in open_issues[:5]:
            sign = "+" if i.diff > 0 else ""
            lines.append(f"  `{i.txn_id}` — {i.merchant} — {sign}{_fmt_inr(i.diff)}")
    else:
        lines.append("All clear — no open mismatches.")
    return "\n".join(lines)


def build_entries_xlsx(db: Session, for_date: Optional[Date] = None) -> tuple[str, bytes]:
    """Return (filename, bytes) of an Excel workbook of entries on a given date."""
    for_date = for_date or Date.today()
    rows = (
        db.query(Entry)
        .filter(Entry.entry_date == for_date)
        .order_by(Entry.cycle, Entry.id)
        .all()
    )
    df = pd.DataFrame(
        [
            {
                "Date": r.entry_date,
                "Cycle": f"C{r.cycle}",
                "Merchant": r.merchant,
                "Company": r.company,
                "Payin": r.payin,
                "Charge": r.charge,
                "GST Rate %": r.gst_rate,
                "GST": r.gst,
                "Chargeback": r.chargeback,
                "Deduction": r.deduction,
                "Settlement": r.settlement,
            }
            for r in rows
        ]
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Entries")
        if not df.empty:
            ws = writer.sheets["Entries"]
            ws.set_column("A:A", 11)
            ws.set_column("B:B", 7)
            ws.set_column("C:C", 14)
            ws.set_column("D:D", 36)
            ws.set_column("E:K", 13)
    buf.seek(0)
    filename = f"settleops_{for_date.isoformat()}.xlsx"
    return filename, buf.getvalue()
