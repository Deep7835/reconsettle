"""Cycle detection — C1 = 6 AM → 4 PM, C2 = 4 PM → 6 AM next day (local time)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class CycleInfo:
    cycle: int
    label: str
    color: str


CYCLE_META: dict[int, CycleInfo] = {
    1: CycleInfo(cycle=1, label="6 AM → 4 PM", color="#10b981"),
    2: CycleInfo(cycle=2, label="4 PM → 6 AM", color="#f59e0b"),
}


def detect_cycle(at: datetime | None = None) -> CycleInfo:
    """Return the cycle for a given moment (defaults to now)."""
    at = at or datetime.now()
    in_day = 6 <= at.hour < 16
    return CYCLE_META[1] if in_day else CYCLE_META[2]
