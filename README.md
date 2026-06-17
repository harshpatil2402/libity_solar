# ☀️ LibityInfotech Solar SaaS v2

Multi-agency solar installation management platform with Supabase backend.

---

## Architecture

```
Superadmin (you)
  └── Creates agencies, sets plan + expiry, uploads logo/stamp
Agency
  └── Manages clients, sends collection links, generates PDFs
Client (no login)
  └── Opens link on phone → uploads Aadhaar, PAN, cheque, signature
```

**Plans:**
- `full` — Client management + Document generation
- `docs_only` — Document generation only

---

## Quick Start

### 1. Supabase Setup

1. Create a project at https://supabase.com
2. Go to SQL Editor → paste contents of `schema.sql` → Run
3. Copy your Project URL and anon/service key

### 2. Local Setup

```bash
cd solar_saas2
cp .env.example .env
# Edit .env with your Supabase URL, key, and credentials

pip install -r requirements.txt
playwright install chromium
playwright install-deps   # Linux only

python run.py
# → http://localhost:5000
```

### 3. Login

| Role | URL | Credentials |
|---|---|---|
| Superadmin | `/superadmin/login` | From `.env` SUPERADMIN_USER/PASS |
| Agency | `/agency/login` | Username/password you set when creating agency |

---

## User Flow

```
Superadmin creates agency
  → Sets plan (full / docs_only), expiry, logo, stamp

Agency logs in → New Client
  → Enters: Name, Mobile, kW, Final Amount

Agency → Send Info Collection Link
  → Client gets SMS with link
  → Client opens on phone:
      • Uploads Aadhaar (crop 11:7)
      • Uploads PAN (crop 16:9)
      • Uploads cancelled cheque (crop 2:1)
      • Draws/uploads signature (crop 3:1)
  → Client submits → status → "Info Collected"

Dashboard notification badge appears

Agency → Generate Documents
  → Aadhaar + signature auto-fetched from submission
  → Fill grid, technical, dates
  → 5 PDFs generated → ZIP download
      1. Commissioning Report
      2. Meter Testing
      3. Model Agreement
      4. Net Metering Agreement
      5. Work Completion Report

Agency continues pipeline:
  Portal Applied → Sanctioned → Installation Done
  → GPS Photo collected → Final Submission → Completed
```

---

## Project Structure

```
solar_saas2/
├── run.py                    ← Entry point
├── app.py                    ← Flask factory
├── config.py                 ← Config from .env
├── db.py                     ← Supabase client + helpers
├── auth.py                   ← Session decorators
├── schema.sql                ← Run once in Supabase SQL Editor
│
├── routes/
│   ├── superadmin.py         ← Agency CRUD, logo/stamp upload
│   ├── agency.py             ← Client mgmt, doc gen, API routes
│   └── public.py             ← Client collection + GPS photo
│
├── utils/
│   ├── helpers.py            ← SMS (Fast2SMS), token generator
│   └── doc_engine.py         ← PDF generation engine
│
├── input_docs/               ← HTML document templates (5 files)
├── static/images/            ← libitylogo.png, mahavitaran_logo.png
│
└── templates/
    ├── base.html             ← LibityInfotech theme (dark mode)
    ├── superadmin/           ← login, dashboard, agency_form, agency_clients
    ├── agency/               ← login, dashboard, new_client, client_detail,
    │                            generate, history
    └── public/               ← collect, gps, invalid, already_submitted
```

---

## Deployment (Hetzner CX32 — recommended)

```bash
# On Ubuntu 22.04 server
apt update && apt install python3-pip nginx certbot python3-certbot-nginx -y
pip install gunicorn
pip install -r requirements.txt
playwright install chromium && playwright install-deps

# Create systemd service
cat > /etc/systemd/system/solar.service << 'EOF'
[Unit]
Description=LibityInfotech Solar SaaS
After=network.target

[Service]
WorkingDirectory=/home/ubuntu/solar_saas2
EnvironmentFile=/home/ubuntu/solar_saas2/.env
ExecStart=/usr/local/bin/gunicorn -w 2 -b 127.0.0.1:5000 run:app
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable solar
systemctl start solar

# Nginx config
cat > /etc/nginx/sites-available/solar << 'EOF'
server {
    server_name yourdomain.com;
    client_max_body_size 50M;
    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_read_timeout 120s;
    }
}
EOF

ln -s /etc/nginx/sites-available/solar /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx
certbot --nginx -d yourdomain.com
```

---

## Environment Variables

| Variable | Description |
|---|---|
| `SECRET_KEY` | Flask session secret — use a long random string |
| `SUPABASE_URL` | Your Supabase project URL |
| `SUPABASE_KEY` | Supabase service role key (full access) |
| `SUPERADMIN_USER` | Your superadmin username |
| `SUPERADMIN_PASS` | Your superadmin password |
| `FAST2SMS_KEY` | Fast2SMS API key (optional — prints to console if missing) |
| `BASE_URL` | Public URL of your app (used in SMS links) |

---

## PDF Generation

Uses **Playwright** (Chromium) with **WeasyPrint** as fallback.

Install Playwright browsers once:
```bash
playwright install chromium
playwright install-deps  # Linux
```

Without Playwright installed, WeasyPrint is used automatically.
