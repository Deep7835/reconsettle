# syntax=docker/dockerfile:1.7

FROM python:3.12-slim AS base

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# curl is used by the HEALTHCHECK; ca-certificates needed for Google/Telegram TLS calls
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl ca-certificates \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# ---- Dependency layer (cached separately from source) ----
COPY pyproject.toml ./
# Copy minimal stubs so `pip install -e .` can resolve the package
RUN mkdir -p app && touch app/__init__.py
RUN pip install --no-cache-dir -e .

# ---- Application code ----
COPY app ./app
COPY scripts ./scripts
COPY alembic.ini ./
COPY alembic ./alembic
COPY data ./data

# Volumes for things that should persist + things that hold secrets
RUN mkdir -p /app/data /app/secrets
VOLUME ["/app/data", "/app/secrets"]

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -fsS http://localhost:8000/api/health || exit 1

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers", "--forwarded-allow-ips=*"]
