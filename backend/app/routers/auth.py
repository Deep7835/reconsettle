"""/api/auth — login + current-user endpoints. PUBLIC (no Bearer required)."""

from __future__ import annotations

import logging
import time

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from app.config import settings
from app.core import security
from app.core.auth_deps import current_user
from app.db import get_db
from app.models import User
from app.schemas.auth import LoginRequest, TokenResponse, UserOut


logger = logging.getLogger("settleops.auth")
router = APIRouter(prefix="/api/auth", tags=["auth"])

# In-memory rate limiter: <=5 failed attempts / 60s per username.
_failed_attempts: dict[str, list[float]] = {}
_MAX_ATTEMPTS = 5
_WINDOW_SECONDS = 60


def _check_rate_limit(username: str) -> None:
    now = time.time()
    attempts = _failed_attempts.setdefault(username, [])
    # Drop expired entries
    cutoff = now - _WINDOW_SECONDS
    while attempts and attempts[0] < cutoff:
        attempts.pop(0)
    if len(attempts) >= _MAX_ATTEMPTS:
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail=f"Too many failed attempts. Try again in {_WINDOW_SECONDS} seconds.",
        )


def _record_failed_attempt(username: str) -> None:
    _failed_attempts.setdefault(username, []).append(time.time())


def _clear_failed_attempts(username: str) -> None:
    _failed_attempts.pop(username, None)


@router.post("/login", response_model=TokenResponse)
def login(payload: LoginRequest, db: Session = Depends(get_db)) -> TokenResponse:
    _check_rate_limit(payload.username)
    user = db.query(User).filter(User.username == payload.username).first()
    if not user or not user.is_active or not security.verify_password(payload.password, user.password_hash):
        _record_failed_attempt(payload.username)
        logger.info("Failed login attempt for %r", payload.username)
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid username or password",
        )
    _clear_failed_attempts(payload.username)
    token = security.create_token(user.id, extra={"u": user.username})
    return TokenResponse(
        access_token=token,
        token_type="bearer",
        expires_in=settings.jwt_expires_hours * 3600,
        user=UserOut.model_validate(user),
    )


@router.get("/me", response_model=UserOut)
def me(user: User = Depends(current_user)) -> UserOut:
    """Verify the current token is valid and return the user info."""
    return UserOut.model_validate(user)
