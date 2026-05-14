"""Application settings loaded from environment / .env file."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, List

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, NoDecode, SettingsConfigDict


BACKEND_DIR = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    """Typed settings backed by .env + environment variables."""

    model_config = SettingsConfigDict(
        env_file=str(BACKEND_DIR / ".env"),
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # Database
    database_url: str = "sqlite:///./settleops.db"

    # CORS
    cors_origins: Annotated[List[str], NoDecode] = Field(
        default_factory=lambda: [
            "http://localhost:5173",
            "http://127.0.0.1:5173",
        ]
    )

    # Business rules
    charge_rate: float = 0.41
    default_gst_rate: float = 0.0

    # Gmail
    gmail_client_secrets_file: Path = BACKEND_DIR / "secrets" / "credentials.json"
    gmail_token_file: Path = BACKEND_DIR / "secrets" / "token.json"
    gmail_search_query: str = "from:bank has:attachment newer_than:7d"
    gmail_sync_interval_minutes: int = 30

    # Telegram
    telegram_bot_token: str = ""
    telegram_report_chat_id: str = ""
    telegram_inbound_chat_ids: Annotated[List[str], NoDecode] = Field(default_factory=list)
    telegram_sync_interval_minutes: int = 10
    telegram_digest_cron: str = ""
    telegram_state_file: Path = BACKEND_DIR / "data" / "telegram_state.json"

    # Auth — single-user. Seeded on startup if missing.
    seed_username: str = "deep7835"
    seed_password: str = "Deep@1234-="
    # JWT signing secret. Auto-generated if not provided (logs out everyone on restart).
    jwt_secret: str = ""
    jwt_alg: str = "HS256"
    jwt_expires_hours: int = 168  # 7 days

    # App
    app_host: str = "0.0.0.0"
    app_port: int = 8000
    app_reload: bool = True
    app_log_level: str = "info"

    @field_validator("cors_origins", "telegram_inbound_chat_ids", mode="before")
    @classmethod
    def _split_csv(cls, v):
        """Accept comma-separated string from env, or pass through a list."""
        if isinstance(v, str):
            return [o.strip() for o in v.split(",") if o.strip()]
        return v


settings = Settings()
