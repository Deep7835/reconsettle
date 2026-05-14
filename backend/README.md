# SettleOps Backend (FastAPI + SQLAlchemy)

Python backend for SettleOps. Owns the database, all business math, Excel
parsing, Gmail-driven bank statement automation, and a clean REST API the
React frontend consumes.

The React app at `/` becomes a thin viewer — it no longer parses Excel or
holds data in memory.

---

## Quick start (Windows / PowerShell)

```powershell
cd backend

# 1. Create a Python virtual env
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Install deps
pip install -e .

# 3. Copy env template
copy .env.example .env

# 4. Seed merchants from the existing App.jsx (one-time)
python -m scripts.extract_merchants
python -m scripts.seed_merchants

# 5. Run the API
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

API runs on `http://localhost:8000`. Interactive docs at **http://localhost:8000/docs**.

CORS is preconfigured for the Vite dev server on `:5173`, so you can call the
API straight from the existing React app once you wire it up.

---

## Project layout

```
backend/
├── pyproject.toml          deps + project metadata
├── alembic.ini             Alembic config
├── alembic/                migrations (env.py + versions/)
├── .env.example            copy to .env to override defaults
├── data/                   seed JSON files
├── secrets/                Gmail credentials.json & token.json
├── scripts/
│   ├── extract_merchants.py    parse src/App.jsx → data/seed_merchants.json
│   └── seed_merchants.py        load seed_merchants.json → DB
└── app/
    ├── main.py             FastAPI app + CORS + lifespan
    ├── config.py           pydantic-settings (reads .env)
    ├── db.py               SQLAlchemy engine + get_db dep
    ├── core/
    │   ├── settle_math.py  Charge/GST/Settlement formula
    │   └── cycle.py        C1 / C2 detection by clock
    ├── models/             SQLAlchemy ORM tables
    ├── schemas/            Pydantic request/response models
    ├── services/
    │   ├── excel_parser.py     Pandas-based bank/source Excel parser
    │   ├── recon_engine.py     Auto-classify + merchant aggregation
    │   ├── gmail_fetcher.py    Google API client + OAuth + downloads
    │   └── scheduler.py        APScheduler interval Gmail sync
    └── routers/
        ├── entries.py      /api/entries CRUD + filters + export
        ├── recon.py        /api/recon CRUD + upload + aggregated
        ├── merchants.py    /api/merchants
        ├── dashboard.py    /api/dashboard/stats
        └── gmail.py        /api/gmail OAuth + sync + runs
```

---

## API surface

Browse **http://localhost:8000/docs** for the full interactive Swagger UI.

| Group       | Method | Path                              | What it does                                 |
| ----------- | ------ | --------------------------------- | -------------------------------------------- |
| meta        | GET    | `/api/health`                     | Liveness probe                                |
| dashboard   | GET    | `/api/dashboard/stats`            | All dashboard numbers in one shot            |
| entries     | GET    | `/api/entries`                    | List with filters (merchant/company/cycle/date) |
| entries     | POST   | `/api/entries`                    | Create entry (computes charge/GST/settlement) |
| entries     | DELETE | `/api/entries/{id}`               | Remove one                                    |
| entries     | DELETE | `/api/entries`                    | Clear all                                     |
| entries     | GET    | `/api/entries/totals`             | Aggregated totals                             |
| entries     | GET    | `/api/entries/grouped?by=merchant`| Group by merchant or company                 |
| entries     | GET    | `/api/entries/export.xlsx`        | Download filtered as Excel                   |
| recon       | GET    | `/api/recon?status=...`           | List recon entries                            |
| recon       | POST   | `/api/recon`                      | Manual single-txn verify (auto-classifies)   |
| recon       | PATCH  | `/api/recon/{id}`                 | Set reason / resolved / notes                 |
| recon       | DELETE | `/api/recon/{id}`                 | Remove one                                    |
| recon       | DELETE | `/api/recon`                      | Clear all                                     |
| recon       | GET    | `/api/recon/stats`                | KPI counts + total discrepancy               |
| recon       | POST   | `/api/recon/upload/bank`          | Upload bank settlement Excel (multipart)     |
| recon       | POST   | `/api/recon/upload/source`        | Upload internal payin Excel (multipart)      |
| recon       | GET    | `/api/recon/aggregated`           | Latest merchant-by-merchant comparison table |
| merchants   | GET    | `/api/merchants?search=&merchant=`| Browse merchants + companies                  |
| merchants   | GET    | `/api/merchants/summary`          | Compact list (name + company_count)          |
| merchants   | POST   | `/api/merchants/companies`        | Add a company (auto-creates merchant)        |
| merchants   | POST   | `/api/merchants/seed`             | Bulk seed (used by seed script)              |
| gmail       | GET    | `/api/gmail/config`               | Current sync settings                         |
| gmail       | PATCH  | `/api/gmail/config`               | Change search query / interval (runtime)     |
| gmail       | GET    | `/api/gmail/auth/start`           | Get the Google OAuth URL                     |
| gmail       | GET    | `/api/gmail/auth/callback`        | OAuth redirect target                         |
| gmail       | GET    | `/api/gmail/messages?query=...`   | Preview messages matching a query             |
| gmail       | POST   | `/api/gmail/sync`                 | Run a sync now, returns the new run          |
| gmail       | GET    | `/api/gmail/runs`                 | Sync history                                  |
| telegram    | GET    | `/api/telegram/config`            | Current bot/chat/cron settings               |
| telegram    | PATCH  | `/api/telegram/config`            | Runtime config update (reschedules jobs)     |
| telegram    | GET    | `/api/telegram/me`                | Verify bot token + identity                  |
| telegram    | POST   | `/api/telegram/sync`              | Pull updates, download bank files            |
| telegram    | GET    | `/api/telegram/runs`              | Inbound + outbound run history               |
| telegram    | POST   | `/api/telegram/send/message`      | Ad-hoc text/markdown to a chat               |
| telegram    | POST   | `/api/telegram/send/report`       | Send digest text + Excel report              |

---

## Gmail automation setup

1. **Google Cloud Console**
   - Create a project (or use an existing one).
   - Enable the **Gmail API**.
   - APIs & Services → OAuth consent screen → set up for "External" with your email as a test user.
   - APIs & Services → Credentials → **Create OAuth Client ID** → Application type: **Web application**.
   - Authorized redirect URI: `http://localhost:8000/api/gmail/auth/callback`
   - Download the JSON → save it as `backend/secrets/credentials.json`.

2. **Authorize the backend**
   - Start the API (`uvicorn app.main:app --reload`).
   - Hit `GET /api/gmail/auth/start` → returns an `auth_url`.
   - Open that URL in a browser, sign in, approve the read-only Gmail scope.
   - Google redirects back to `/api/gmail/auth/callback` and the backend saves
     `secrets/token.json`. Subsequent runs use this token (refresh handled).

3. **Configure the search query**
   - Tweak `GMAIL_SEARCH_QUERY` in `.env` — uses Gmail's normal search syntax.
     Example: `from:noreply@yourbank.in subject:"settlement" has:attachment newer_than:2d`
   - Or change it at runtime: `PATCH /api/gmail/config` with `{"search_query": "..."}`.

4. **Sync**
   - `POST /api/gmail/sync` runs a sync immediately and stores any attached
     `.xlsx` / `.xls` / `.csv` files as `BankUpload` (filenames containing
     "bank" / "settle" / "mid") or `SourceUpload`.
   - The scheduler also runs the sync every `GMAIL_SYNC_INTERVAL_MINUTES`
     minutes (defaults to 30).

5. **Inspect**
   - `GET /api/gmail/runs` shows the history (status, messages_seen,
     attachments_saved, uploads_created, error_message).
   - `GET /api/recon/uploads/bank` shows the parsed bank statements.
   - `GET /api/recon/aggregated` runs the merchant-by-merchant reconciliation
     against the latest source + bank uploads.

---

## Telegram automation setup

Two flows, both driven by a single bot:

1. **Inbound — pull bank statements from a Telegram group**
   - Create a bot via `@BotFather` on Telegram → he hands you a token like
     `123456789:ABCdefGHI...`. Paste it into `.env` as `TELEGRAM_BOT_TOKEN`.
   - Create a private Telegram group, add your bot to it.
   - **Important:** in BotFather → `/setprivacy` → choose **Disable** so the
     bot can see all messages in the group, not just commands.
   - Drop bank settlement `.xlsx` / `.csv` files into the group.
   - First time: hit `POST /api/telegram/sync` — it returns a run that includes
     `chat_id` for whatever chats were seen. Copy that and put it in
     `TELEGRAM_INBOUND_CHAT_IDS` (comma-separated allowlist) and
     `TELEGRAM_REPORT_CHAT_ID` (for outbound).
   - Set `TELEGRAM_SYNC_INTERVAL_MINUTES=10` and restart — the scheduler will
     poll the bot every 10 minutes and ingest new files into `BankUpload` /
     `SourceUpload` (same routing as Gmail: filenames containing
     `bank` / `settle` / `mid` → bank, else → source).

2. **Outbound — push reports to Telegram**
   - `POST /api/telegram/send/report` → builds a markdown digest of today's
     totals + open recon issues, sends it, then attaches today's Excel.
   - `POST /api/telegram/send/message` → ad-hoc text (handy for alerts).
   - Optional daily digest: set `TELEGRAM_DIGEST_CRON=30 18 * * *` to push the
     report every day at 18:30. Standard 5-field cron syntax.

Verify everything is wired with:

```bash
curl -s http://localhost:8000/api/telegram/config
curl -s http://localhost:8000/api/telegram/me      # tests bot token
curl -X POST http://localhost:8000/api/telegram/sync
curl -X POST http://localhost:8000/api/telegram/send/report \
     -H "Content-Type: application/json" \
     -d '{"include_excel": true}'
```

`telegram_runs` table audits every inbound poll, outbound message and digest
with status / error / counts — browse via `GET /api/telegram/runs`.

---

## Excel parser expectations

The parser is **header-driven** and case-insensitive — it accepts synonyms so
different bank templates still work.

### Bank settlement file

| Logical name | Synonyms it looks for |
|---|---|
| `mid`        | mid, merchant id, merchantid |
| `company`    | name, beneficiary, company, merchant name, party name |
| `amount`     | amount, credit, gross, txn amount, transaction amount |
| `fee`        | fee, commission, charge, mdr |
| `gst`        | gst, tax, igst, cgst+sgst, cgst, sgst |
| `settle`     | settle, net, net settle, settlement, net amount |
| `chargeback` | chargeback, refund, reversal, debit |

`Company` is required. Negative chargeback amounts get flipped to positive.

### Internal payin file

| Logical name | Synonyms |
|---|---|
| `company` | row labels, name, beneficiary, company, merchant name |
| `amount`  | sum of amount, amount, total amount, total |

Both required. Rows labelled "Grand Total" / "Total" are skipped.

If headers don't match, the endpoint returns 400 with a clear error message.

---

## Database

- Default: **SQLite** at `backend/settleops.db` (no setup required).
- Switch to Postgres by setting `DATABASE_URL=postgresql+psycopg://user:pass@host:5432/db`.
- Tables auto-create on first boot (lifespan handler calls `Base.metadata.create_all`).
- For prod: use Alembic migrations:

  ```bash
  alembic revision --autogenerate -m "init"
  alembic upgrade head
  ```

---

## Wiring the React frontend (next step)

The React app needs to swap its local `useState` arrays for calls to this API.
Drop-in starting point with `fetch` + `react-query`:

```js
// src/api.ts
const BASE = "http://localhost:8000";
export const api = {
  entries: () => fetch(`${BASE}/api/entries`).then(r => r.json()),
  createEntry: (body) => fetch(`${BASE}/api/entries`, {
    method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(body)
  }).then(r => r.json()),
  dashboard: () => fetch(`${BASE}/api/dashboard/stats`).then(r => r.json()),
  uploadBank: (file) => {
    const fd = new FormData(); fd.append("file", file);
    return fetch(`${BASE}/api/recon/upload/bank`, { method: "POST", body: fd }).then(r => r.json());
  },
  syncGmail: () => fetch(`${BASE}/api/gmail/sync`, { method: "POST" }).then(r => r.json()),
};
```

Then in App.jsx replace `useState([])` arrays + their setters with `useQuery`
+ `useMutation` calls. The UI / styling stays the same.

---

## Smoke test

```bash
# Health check
curl http://localhost:8000/api/health

# Create an entry
curl -X POST http://localhost:8000/api/entries \
  -H "Content-Type: application/json" \
  -d '{"merchant":"Aryan","company":"ECOMPANTHER","payin":10000,"gst_rate":18}'

# Get dashboard snapshot
curl http://localhost:8000/api/dashboard/stats | python -m json.tool
```
