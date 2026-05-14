"""Merchant + Company API schemas."""

from __future__ import annotations

from typing import List, Optional

from pydantic import BaseModel, ConfigDict, Field


class CompanyBase(BaseModel):
    name: str
    mid: Optional[str] = None
    gst_number: Optional[str] = None
    moa: Optional[str] = None
    incorporation: Optional[str] = None


class CompanyCreate(CompanyBase):
    merchant_name: str = Field(..., description="Merchant name, auto-created if new")


class CompanyOut(CompanyBase):
    id: int
    merchant_id: int

    model_config = ConfigDict(from_attributes=True)


class MerchantOut(BaseModel):
    id: int
    name: str
    company_count: int = 0
    companies: List[CompanyOut] = []

    model_config = ConfigDict(from_attributes=True)


class MerchantSummary(BaseModel):
    name: str
    company_count: int


class MerchantSeed(BaseModel):
    """Used by the seed script and by POST /api/merchants for bulk add."""

    merchant: str
    company: str
    mid: Optional[str] = None
    gst_number: Optional[str] = ""
    moa: Optional[str] = ""
    incorporation: Optional[str] = ""
