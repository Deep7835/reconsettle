"""One-time extractor: read MASTER_DATA from ../../src/App.jsx → data/seed_merchants.json.

Usage:
    python -m scripts.extract_merchants

Then to load into the DB:
    python -m scripts.seed_merchants
"""

from __future__ import annotations

import json
import re
from pathlib import Path


BACKEND_DIR = Path(__file__).resolve().parent.parent
APP_JSX = BACKEND_DIR.parent / "src" / "App.jsx"
OUT = BACKEND_DIR / "data" / "seed_merchants.json"


def extract() -> list[dict]:
    text = APP_JSX.read_text(encoding="utf-8")

    # Find MASTER_DATA = [ ... ];
    match = re.search(r"const\s+MASTER_DATA\s*=\s*\[(.*?)\];", text, re.DOTALL)
    if not match:
        raise SystemExit("Could not locate MASTER_DATA in App.jsx")

    body = match.group(1)

    # Match each { mid: "...", merchant: "...", company: "...", moa: "...", incorporation: "...", gstNumber: "..." }
    entry_re = re.compile(
        r'\{\s*'
        r'mid:\s*"(?P<mid>[^"]*)"\s*,\s*'
        r'merchant:\s*"(?P<merchant>[^"]*)"\s*,\s*'
        r'company:\s*"(?P<company>[^"]*)"\s*,\s*'
        r'moa:\s*"(?P<moa>[^"]*)"\s*,\s*'
        r'incorporation:\s*"(?P<incorporation>[^"]*)"\s*,\s*'
        r'gstNumber:\s*"(?P<gstNumber>[^"]*)"\s*'
        r'\}',
    )

    rows: list[dict] = []
    for m in entry_re.finditer(body):
        d = m.groupdict()
        rows.append(
            {
                "mid": d["mid"] or None,
                "merchant": d["merchant"],
                "company": d["company"],
                "moa": d["moa"] or None,
                "incorporation": d["incorporation"] or None,
                "gst_number": d["gstNumber"] or None,
            }
        )

    return rows


def main() -> None:
    rows = extract()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Extracted {len(rows)} merchant rows -> {OUT}")
    merchants = sorted({r['merchant'] for r in rows})
    print(f"Unique merchants ({len(merchants)}): {', '.join(merchants)}")


if __name__ == "__main__":
    main()
