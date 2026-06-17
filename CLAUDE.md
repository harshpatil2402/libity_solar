# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt
playwright install chromium
playwright install-deps   # Linux only

# Run development server
python run.py
# → http://localhost:5000

# Production (gunicorn)
gunicorn -w 2 -b 127.0.0.1:5000 run:app
```

Environment: copy `.env.example` to `.env` and fill in `SECRET_KEY`, `SUPABASE_URL`, `SUPABASE_KEY`, `SUPERADMIN_USER`, `SUPERADMIN_PASS`, `FAST2SMS_KEY`, `BASE_URL`.

Database: paste `schema.sql` into Supabase SQL Editor and run once.

## Architecture

Three-tier multi-tenant Flask app: **Superadmin → Agency → Client (no login)**.

**Role separation:**
- `Superadmin` (credentials from `.env`) — manages agencies via `/superadmin/*`
- `Agency` (credentials in Supabase `agencies` table) — manages clients, sends links, generates PDFs via `/agency/*`
- `Client` — no account; receives a tokenized SMS link (`/collect/<token>` or `/gps/<token>`), submits docs/photos once

**Plans:** `full` (client management + doc generation) or `docs_only` (doc generation only). Enforced by `@feature_required('client_mgmt')` decorator in `auth.py`.

**Auth:** Flask server-side sessions. `auth.py` provides `@superadmin_required` and `@agency_required` decorators. The hardcoded fallback agency (`admin`/`admin123`) in `auth.py` and `routes/agency.py` bypasses Supabase for local dev.

**Database layer (`db.py`):** Thin wrappers — `fetch_one`, `fetch_all`, `insert_row`, `update_row`, `delete_row` — around the Supabase Python client. All image data (Aadhaar, PAN, cheque, signature, agency logo, stamp) is stored as base64 strings in Supabase, not as files.

**Client pipeline statuses** (defined in `routes/agency.py::STATUSES`):
`New Lead → Link Sent → Info Collected → Documents Generated → Portal Applied → Sanctioned → Installation Done → GPS Photos Received → Final Submission Done → Completed`

**PDF generation (`utils/doc_engine.py`):**
- HTML templates in `input_docs/` use `{{placeholder}}` syntax (not Jinja)
- `preload_templates()` is called once at app startup to cache all 5 HTMLs in `_HTML_CACHE`
- `run_job()` runs in a background thread; progress is streamed via polling `/agency/api/job/<jid>/status`
- PDF rendering uses Playwright (Chromium) with WeasyPrint as fallback
- 5 PDFs are rendered in parallel via `ThreadPoolExecutor(max_workers=5)` then bundled into a ZIP
- Completed ZIP bytes are held in-memory in `jobs` dict (lost on server restart; `doc_jobs` table tracks history metadata only)

**WhatsApp (`utils/whatsapp.py`):** Opens WhatsApp Web in a persistent Chromium window (`wa_profile/` directory) with a pre-filled message. Non-headless — user must click Send manually. Requires prior WhatsApp Web login.

**SMS (`utils/helpers.py`):** Uses Fast2SMS API. If `FAST2SMS_KEY` is not set, messages are printed to console instead.

**Blueprint registration (`app.py`):** `sa_bp` (`/superadmin`), `agency_bp` (`/agency`), `public_bp` (no prefix — `/collect/<token>` and `/gps/<token>`). Root `/` redirects to agency login.
