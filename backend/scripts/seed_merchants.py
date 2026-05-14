"""Load data/seed_merchants.json into the database.

Usage:
    python -m scripts.seed_merchants
"""

from __future__ import annotations

import json
from pathlib import Path

from app.db import SessionLocal, engine
from app.models import Base, Company, Merchant


BACKEND_DIR = Path(__file__).resolve().parent.parent
SEED = BACKEND_DIR / "data" / "seed_merchants.json"


def main() -> None:
    if not SEED.exists():
        raise SystemExit(
            f"Seed file missing: {SEED}\n"
            "Run `python -m scripts.extract_merchants` first."
        )

    # Ensure tables exist
    Base.metadata.create_all(bind=engine)

    rows = json.loads(SEED.read_text(encoding="utf-8"))
    db = SessionLocal()
    try:
        cache = {m.name: m for m in db.query(Merchant).all()}
        added_m, added_c = 0, 0
        for r in rows:
            mname = (r.get("merchant") or "").strip()
            cname = (r.get("company") or "").strip()
            if not mname or not cname:
                continue
            merchant = cache.get(mname)
            if not merchant:
                merchant = Merchant(name=mname)
                db.add(merchant)
                db.flush()
                cache[mname] = merchant
                added_m += 1
            exists = (
                db.query(Company)
                .filter(Company.merchant_id == merchant.id, Company.name == cname)
                .first()
            )
            if exists:
                continue
            db.add(
                Company(
                    merchant_id=merchant.id,
                    name=cname,
                    mid=r.get("mid") or None,
                    gst_number=r.get("gst_number") or None,
                    moa=r.get("moa") or None,
                    incorporation=r.get("incorporation") or None,
                )
            )
            added_c += 1
        db.commit()
        print(f"Seeded {added_m} new merchants, {added_c} new companies.")
    finally:
        db.close()


if __name__ == "__main__":
    main()
