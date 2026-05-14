"""Database engine, session factory, and FastAPI dependency."""

from __future__ import annotations

from collections.abc import Generator

from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

from app.config import settings


_engine_kwargs: dict = {}
if settings.database_url.startswith("sqlite"):
    # SQLite needs this for use across threads (FastAPI runs handlers in threadpools)
    _engine_kwargs["connect_args"] = {"check_same_thread": False}


engine: Engine = create_engine(
    settings.database_url,
    future=True,
    pool_pre_ping=True,
    **_engine_kwargs,
)


SessionLocal = sessionmaker(
    bind=engine,
    autoflush=False,
    autocommit=False,
    future=True,
    expire_on_commit=False,
)


def get_db() -> Generator[Session, None, None]:
    """FastAPI dependency that yields a scoped session and closes it after."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
