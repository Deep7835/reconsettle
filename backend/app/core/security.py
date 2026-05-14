"""Password hashing (PBKDF2-SHA256) + JWT helpers.

We use stdlib `hashlib.pbkdf2_hmac` rather than bcrypt to avoid pulling in a
native dep on Windows. PBKDF2 with 200k iterations is well within OWASP
guidance for password hashing.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import os
import secrets
from datetime import datetime, timedelta, timezone
from typing import Optional

import jwt

from app.config import settings


PBKDF2_ITERATIONS = 200_000
SALT_BYTES = 16


# ---------- Password hashing ----------


def hash_password(password: str) -> str:
    """Return a self-describing hash string: `pbkdf2_sha256$<iters>$<b64 salt>$<b64 hash>`."""
    salt = secrets.token_bytes(SALT_BYTES)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return "pbkdf2_sha256${}${}${}".format(
        PBKDF2_ITERATIONS,
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(digest).decode("ascii"),
    )


def verify_password(password: str, stored: str) -> bool:
    """Constant-time verification against a hash produced by `hash_password`."""
    try:
        scheme, iters_s, salt_b64, digest_b64 = stored.split("$")
        if scheme != "pbkdf2_sha256":
            return False
        iters = int(iters_s)
        salt = base64.b64decode(salt_b64)
        expected = base64.b64decode(digest_b64)
    except (ValueError, TypeError):
        return False
    computed = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iters)
    return hmac.compare_digest(expected, computed)


# ---------- JWT ----------


_jwt_secret_cache: Optional[str] = None


def _get_secret() -> str:
    """Resolve the signing secret. Generate a strong random one if not configured.

    The generated secret only lives for the process lifetime — restart invalidates
    all tokens. For prod set JWT_SECRET in .env to keep tokens stable.
    """
    global _jwt_secret_cache
    if _jwt_secret_cache:
        return _jwt_secret_cache
    if settings.jwt_secret:
        _jwt_secret_cache = settings.jwt_secret
    else:
        _jwt_secret_cache = secrets.token_urlsafe(48)
    return _jwt_secret_cache


def create_token(subject: int | str, extra: dict | None = None) -> str:
    now = datetime.now(tz=timezone.utc)
    payload = {
        "sub": str(subject),
        "iat": int(now.timestamp()),
        "exp": int((now + timedelta(hours=settings.jwt_expires_hours)).timestamp()),
    }
    if extra:
        payload.update(extra)
    return jwt.encode(payload, _get_secret(), algorithm=settings.jwt_alg)


def decode_token(token: str) -> dict:
    return jwt.decode(token, _get_secret(), algorithms=[settings.jwt_alg])
