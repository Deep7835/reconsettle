"""Smoke tests: verify the app imports, computes math correctly, and parses Excel."""

from __future__ import annotations

import io

import pandas as pd

from app.core.cycle import detect_cycle
from app.core.settle_math import compute_settlement
from app.services.excel_parser import parse_bank_excel, parse_source_excel


def test_settle_math() -> None:
    b = compute_settlement(payin=10000, gst_rate=18, chargeback=0)
    assert b.charge == 41.0
    assert b.gst == 7.38
    assert b.deduction == 48.38
    assert b.settlement == 9951.62


def test_settle_math_with_chargeback() -> None:
    b = compute_settlement(payin=10000, gst_rate=0, chargeback=500)
    assert b.charge == 41.0
    assert b.gst == 0.0
    assert b.deduction == 541.0
    assert b.settlement == 9459.0


def test_cycle_detect() -> None:
    info = detect_cycle()
    assert info.cycle in (1, 2)


def test_parse_source_excel() -> None:
    df = pd.DataFrame(
        {"Row Labels": ["Acme Pvt Ltd", "Beta LLP", "Grand Total"], "Sum of Amount": [1000.5, 2500, 3500.5]}
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    parsed = parse_source_excel(buf.getvalue())
    names = {r["company"] for r in parsed.rows}
    assert "Acme Pvt Ltd" in names
    assert "Beta LLP" in names
    assert "Grand Total" not in names  # should be filtered out
    assert parsed.total_amount == 3500.5


def test_parse_bank_excel() -> None:
    df = pd.DataFrame(
        {
            "MID": ["MID001", "MID002"],
            "Name": ["Acme Pvt Ltd", "Beta LLP"],
            "Amount": [10000, 20000],
            "Fee": [41, 82],
            "GST": [7.38, 14.76],
            "Settle": [9951.62, 19903.24],
            "Chargeback": [0, -500],
        }
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    parsed = parse_bank_excel(buf.getvalue())
    assert len(parsed.rows) == 2
    assert parsed.total_amount == 30000
    beta = next(r for r in parsed.rows if r["company"] == "Beta LLP")
    assert beta["chargeback"] == 500  # negative flipped to positive


def test_app_imports() -> None:
    """Just import the FastAPI app — catches wiring problems."""
    from app.main import app

    assert app.title == "SettleOps API"


def test_report_builder_digest_text() -> None:
    """Digest builder works on an empty DB and produces valid Markdown."""
    from app.db import SessionLocal
    from app.models import Base
    from app.db import engine
    from app.services import report_builder

    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        text = report_builder.build_digest_text(db)
        assert "SettleOps Digest" in text
        assert "All-time" in text
        # Should never reference Unicode arrow that crashes on Windows console
        assert "\\u" not in text
    finally:
        db.close()


def test_report_builder_entries_xlsx() -> None:
    """Excel builder returns bytes that look like an Excel file."""
    from app.db import SessionLocal
    from app.models import Base
    from app.db import engine
    from app.services import report_builder

    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        filename, data = report_builder.build_entries_xlsx(db)
        assert filename.endswith(".xlsx")
        assert data.startswith(b"PK")  # xlsx files are zip-based and start with PK
    finally:
        db.close()


def test_telegram_client_not_configured() -> None:
    """When no token is set, the client raises a clear, typed error."""
    from app.config import settings
    from app.services import telegram_client

    original = settings.telegram_bot_token
    settings.telegram_bot_token = ""
    try:
        try:
            telegram_client.get_me()
        except telegram_client.TelegramNotConfigured as e:
            assert "TELEGRAM_BOT_TOKEN" in str(e)
        else:
            raise AssertionError("expected TelegramNotConfigured")
    finally:
        settings.telegram_bot_token = original
