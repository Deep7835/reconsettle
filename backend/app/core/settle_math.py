"""Settlement math — mirrors the formula used in the React frontend.

charge      = payin * charge_rate / 100        (your commission)
gst         = charge * gst_rate / 100          (GST on commission)
deduction   = charge + gst + chargeback
settlement  = payin - deduction                (paid to merchant)
"""

from __future__ import annotations

from dataclasses import dataclass

from app.config import settings


@dataclass(frozen=True)
class SettlementBreakdown:
    payin: float
    chargeback: float
    charge: float
    gst: float
    gst_rate: float
    deduction: float
    settlement: float


def compute_settlement(
    payin: float,
    chargeback: float = 0.0,
    gst_rate: float | None = None,
    charge_rate: float | None = None,
) -> SettlementBreakdown:
    """Compute charge / gst / deduction / final settlement for one transaction."""
    payin = float(payin or 0.0)
    chargeback = float(chargeback or 0.0)
    gst_rate = float(gst_rate if gst_rate is not None else settings.default_gst_rate)
    charge_rate = float(charge_rate if charge_rate is not None else settings.charge_rate)

    charge = (payin * charge_rate) / 100.0
    gst = (charge * gst_rate) / 100.0
    deduction = charge + gst + chargeback
    settlement = payin - deduction

    return SettlementBreakdown(
        payin=round(payin, 2),
        chargeback=round(chargeback, 2),
        charge=round(charge, 2),
        gst=round(gst, 2),
        gst_rate=gst_rate,
        deduction=round(deduction, 2),
        settlement=round(settlement, 2),
    )
