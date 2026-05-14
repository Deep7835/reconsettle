# SettleOps — Deployment

Two services, one Compose file. The frontend container terminates HTTP traffic
and reverse-proxies `/api/*` to the backend container — so you only ever expose
**one port** to the internet.

```
              ┌──────────────────────┐
  internet ──►│  frontend (nginx)    │── port 80
              │  - serves React SPA  │
              │  - proxies /api/ →   │
              │    backend:8000      │
              └──────────┬───────────┘
                         │
                         ▼
              ┌──────────────────────┐
              │  backend (uvicorn)   │── port 8000 (internal)
              │  - FastAPI + SQLite  │
              │  - APScheduler jobs  │
              └──────────────────────┘

  volumes:
    backend_data    →  /app/data    (SQLite DB, telegram_state.json)
    backend_secrets →  /app/secrets (Gmail credentials.json + token.json)
```

---

## 1. First-time setup on the server

Requirements: any Linux VPS with **Docker + Docker Compose v2** installed.

```bash
# 1. Clone the repo
git clone https://github.com/<you>/<repo>.git settleops
cd settleops

# 2. Create the production env file
cp .env.production.example .env
nano .env   # fill in JWT_SECRET, TELEGRAM_BOT_TOKEN, CORS_ORIGINS, etc.

# 3. Generate a strong JWT secret and paste it into .env
openssl rand -base64 48

# 4. Build + start
docker compose up -d --build

# 5. Verify
docker compose ps                              # both services should be healthy
curl -fsS http://localhost/api/health           # → {"status":"ok",...}
curl -fsS http://localhost/                     # → React SPA HTML
```

Login on the browser with `deep7835` / `Deep@1234-=` (whatever you set as
`SEED_USERNAME` / `SEED_PASSWORD` in `.env`).

---

## 2. Updating after a code change

```bash
git pull
docker compose up -d --build
```

Compose picks up changed files, rebuilds only what changed (layer cache), and
restarts containers gracefully. Volumes are untouched so the DB + Gmail token
+ Telegram state survive.

---

## 3. Logs

```bash
docker compose logs -f backend                  # follow backend logs
docker compose logs -f frontend                 # follow nginx access/errors
docker compose logs --tail=200                  # both, last 200 lines
```

---

## 4. Persistent data

| Volume                | Path inside container | What's stored                          |
|-----------------------|------------------------|----------------------------------------|
| `settleops_backend_data`    | `/app/data`             | `settleops.db`, `telegram_state.json` |
| `settleops_backend_secrets` | `/app/secrets`          | `credentials.json`, `token.json`      |

**Inspect a volume:**

```bash
docker volume inspect settleops_backend_data
docker run --rm -v settleops_backend_data:/data alpine ls -la /data
```

**Back up the DB:**

```bash
docker run --rm \
  -v settleops_backend_data:/data \
  -v $(pwd)/backups:/out \
  alpine cp /data/settleops.db /out/settleops-$(date +%Y%m%d).db
```

**Drop the Gmail OAuth `credentials.json` into the volume:**

```bash
# After running `docker compose up -d` once, the volume exists
docker cp /path/to/credentials.json settleops-backend:/app/secrets/credentials.json
docker compose restart backend
```

---

## 5. Putting it behind HTTPS (Caddy — easiest)

`Caddyfile` (next to docker-compose.yml):

```
yourdomain.com {
    reverse_proxy localhost:80
}
```

Then run a Caddy container or install Caddy on the host. Caddy handles the TLS
cert via Let's Encrypt automatically.

**Alternative — Traefik or nginx-host** also work; just point them at
`http://localhost:80`. Don't expose the backend port externally; let the
frontend nginx do the proxying.

---

## 6. Switching SQLite → Postgres later

Add a Postgres service to `docker-compose.yml`:

```yaml
  db:
    image: postgres:16
    environment:
      POSTGRES_DB: settleops
      POSTGRES_USER: settleops
      POSTGRES_PASSWORD: ${POSTGRES_PASSWORD}
    volumes:
      - pg_data:/var/lib/postgresql/data
    networks:
      - settleops

volumes:
  pg_data:
```

Update `.env`:

```
DATABASE_URL=postgresql+psycopg://settleops:<password>@db:5432/settleops
```

…and add `psycopg[binary]` to `backend/pyproject.toml`. `Base.metadata.create_all`
in the lifespan handler will create tables on first boot, then run
`docker compose exec backend python -m scripts.seed_merchants`.

---

## 7. Troubleshooting

**Frontend shows "Backend offline" banner**

The frontend probes `/api/health` directly. If nginx can't reach the backend
service:

```bash
docker compose exec frontend wget -qO- http://backend:8000/api/health
```

Should print `{"status":"ok",...}`. If not, backend isn't healthy — check
`docker compose logs backend`.

**"Invalid token" after restart**

Set `JWT_SECRET` in `.env`. With it empty, every container restart generates a
fresh secret and old tokens become invalid.

**Excel uploads fail with 413**

`client_max_body_size 25M` is set in `nginx.conf`. Bump it if you need bigger
files, then rebuild the frontend image.

**Schedule jobs not firing**

The APScheduler runs inside the backend process. Check
`docker compose logs backend | grep -i scheduler` — you should see
"Scheduler started" + "Scheduled Gmail/Telegram sync every X minutes".

---

## 8. One-liners

| Goal | Command |
|---|---|
| Stop everything | `docker compose down` |
| Stop + wipe data (DANGEROUS) | `docker compose down -v` |
| Rebuild from scratch | `docker compose build --no-cache && docker compose up -d` |
| Shell into backend | `docker compose exec backend bash` |
| Open Swagger UI | `http://localhost/docs` |
| Reset Telegram offset | `curl -X POST http://localhost/api/telegram/state/reset -H "Authorization: Bearer <jwt>"` |
| Run pytest | `docker compose exec backend pytest -q` |
