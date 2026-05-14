"""FastAPI app entrypoint."""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastapi import Depends, FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import settings
from app.core import security
from app.core.auth_deps import current_user
from app.db import SessionLocal, engine
from app.models import Base, User
from app.routers import auth, dashboard, entries, gmail, merchants, recon, telegram
from app.services import scheduler


logging.basicConfig(
    level=settings.app_log_level.upper(),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("settleops")


def seed_default_user() -> None:
    """Create the seeded admin user if it doesn't exist yet."""
    db = SessionLocal()
    try:
        existing = db.query(User).filter(User.username == settings.seed_username).first()
        if existing:
            return
        if not settings.seed_password:
            logger.warning("SEED_PASSWORD is empty; skipping seed user creation")
            return
        user = User(
            username=settings.seed_username,
            password_hash=security.hash_password(settings.seed_password),
            is_active=True,
        )
        db.add(user)
        db.commit()
        logger.info("Seeded default user %r", settings.seed_username)
    finally:
        db.close()


@asynccontextmanager
async def lifespan(app: FastAPI):
    Base.metadata.create_all(bind=engine)
    logger.info("Database schema ensured")

    seed_default_user()

    scheduler.start()
    logger.info("Scheduler started")

    yield

    scheduler.shutdown()
    logger.info("Scheduler stopped")


app = FastAPI(
    title="SettleOps API",
    description="Backend for SettleOps: settlements, reconciliation, automation.",
    version="0.1.0",
    lifespan=lifespan,
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Public routers (no Bearer token required)
app.include_router(auth.router)


# Protected routers — every endpoint requires Authorization: Bearer <jwt>
PROTECTED = [Depends(current_user)]
app.include_router(dashboard.router, dependencies=PROTECTED)
app.include_router(entries.router, dependencies=PROTECTED)
app.include_router(recon.router, dependencies=PROTECTED)
app.include_router(merchants.router, dependencies=PROTECTED)
app.include_router(gmail.router, dependencies=PROTECTED)
app.include_router(telegram.router, dependencies=PROTECTED)


@app.get("/api/health", tags=["meta"])
def health() -> dict:
    """PUBLIC health check — used by the frontend to detect if the backend is up."""
    return {
        "status": "ok",
        "service": "settleops-backend",
        "version": "0.1.0",
        "database": settings.database_url.split("://")[0],
    }


@app.get("/", include_in_schema=False)
def root() -> dict:
    return {
        "name": "SettleOps API",
        "docs": "/docs",
        "openapi": "/openapi.json",
        "health": "/api/health",
    }
