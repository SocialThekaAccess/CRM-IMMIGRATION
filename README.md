# Immigration Management System (IMS) — CRM

A full-featured immigration consultancy CRM built with Python/Flask. Supports two companies (ABC & XYZ), multi-role access, lead management, enrollment tracking, finance, employee performance, and real-time competition system.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.11, Flask 3.x |
| Database | SQLite (local/dev), PostgreSQL (cloud/Render) |
| Frontend | Jinja2 HTML templates, vanilla JS, CSS |
| Excel Export | openpyxl |
| WSGI Server | Gunicorn (cloud), Waitress (LAN) |
| Deployment | Render (free tier), Railway |

---

## Project Structure

```
the-immigration-crm/
│
├── app.py                   # App entry point (local run)
├── wsgi.py                  # Gunicorn entry point (Render/cloud) + startup DB reset
├── legacy_core.py           # CORE FILE — Flask app, all routes, DB schema, templates
├── registry.py              # Top-level module registry
├── requirements.txt         # Python dependencies
├── render.yaml              # Render.com deployment config
├── Procfile                 # Heroku-style process config
│
├── modules/                 # Feature blueprints (Flask Blueprints)
│   ├── am_activity/         # AM (Assistant Manager) activity tracking
│   ├── bank_manager/        # Bank accounts + reconciliation
│   ├── direct_lead/         # Direct lead entry form
│   ├── fb_leads/            # Facebook Lead Ads integration
│   ├── leads/               # Lead upload, allocation, management
│   ├── lead_statuses/       # Lead status configuration
│   ├── no_chat/             # Chat disable module
│   ├── performance/         # Employee performance scoring
│   ├── report_tools/        # Excel/PDF/WhatsApp export tools
│   ├── runtime_lite/        # Lightweight runtime utilities
│   ├── security_settings/   # User management, OTP password reset
│   ├── system/              # System-level settings
│   └── registry.py          # Module registration
│
├── templates/               # Jinja2 HTML templates (on-disk)
│   ├── base.html            # Base layout (sidebar, nav)
│   ├── login_gaur.html      # Login page
│   ├── dashboard.html       # MD/GM executive dashboard
│   ├── leads.html           # Lead management
│   ├── ams.html             # Assistant Manager management
│   ├── cases.html           # Enrollment / payment / filing
│   ├── employees.html       # Employee directory
│   ├── reception.html       # Reception / visitor log
│   └── ...                  # Other pages
│
├── static/                  # Static assets (CSS, JS, images)
├── uploads/                 # File uploads (local only)
└── .env                     # Environment variables (never committed)
```

---

## Architecture Overview

```
Browser
   │
   ▼
Flask App (legacy_core.py)
   │
   ├── Routes / Views ──────────────────────────────────────────────
   │   ├── /                    → Login page
   │   ├── /dashboard           → Executive Dashboard (MD/GM)
   │   ├── /leads               → Lead upload & allocation
   │   ├── /ams                 → Assistant Manager management
   │   ├── /cases               → Enrollment / Payment / Filing
   │   ├── /employees           → Employee directory
   │   ├── /reception           → Visitor / reception log
   │   ├── /performance         → Performance analysis
   │   ├── /my-leads            → AM personal lead view
   │   └── ...
   │
   ├── Blueprints (modules/) ───────────────────────────────────────
   │   ├── /v4/leads/           → Lead management API
   │   ├── /v4/bank/            → Bank manager
   │   ├── /v4/fb-leads/        → Facebook Lead Ads webhook
   │   ├── /v4/report-tools/    → Excel/PDF export
   │   ├── /v4/performance/     → Performance module
   │   └── /security/           → Security settings
   │
   ├── Database Layer ──────────────────────────────────────────────
   │   ├── SQLite (local): mini_crm.db
   │   └── PostgreSQL (cloud): DATABASE_URL env var
   │
   └── Template Engine ─────────────────────────────────────────────
       ├── On-disk: templates/*.html (Jinja2)
       └── Embedded: _GAUR_V337_TEMPLATES dict (in legacy_core.py)
           └── Used as fallback / cloud override
```

---

## Database Schema (31 Tables)

| Table | Purpose |
|---|---|
| `users` | Portal users (MD, GM, AM, Reception) |
| `companies` | Company registry (SCIC=ABC, WWIC=XYZ) |
| `leads` | Immigration leads/prospects |
| `employee_master` | Employee records |
| `client_cases` | Enrollment, payment, filing cases |
| `lead_activity` | Lead follow-up activity log |
| `lead_documents` | Documents attached to leads |
| `employee_documents` | Employee documents |
| `employee_performance` | Monthly performance scores |
| `allocation_history` | Lead allocation audit trail |
| `finance_banks` | Bank accounts |
| `finance_transactions` | Financial transactions |
| `bank_recon_runs` | Bank reconciliation runs |
| `bank_statement_rows` | Imported bank statement rows |
| `bank_recon_matches` | Reconciliation matches |
| `visitors` | Reception visitor log |
| `employee_attendance` | Attendance records |
| `imports` | Lead import history |
| `chat_messages` | Internal chat |
| `broadcasts` | Broadcast messages |
| `activity_logs` | System activity audit |
| `competition_settings` | Competition configuration |
| `competition_events` | Competition enrollment events |
| `notification_reads` | Notification read status |
| `fb_settings` | Facebook Lead Ads config |
| `fb_fetch_log` | Facebook fetch log |
| `fb_form_map` | Facebook form field mapping |
| `password_reset_otps` | OTP tokens for password reset |
| `security_audit` | Security event log |

---

## User Roles

| Role | Access |
|---|---|
| `MD` | Full access — both companies, all reports, user management |
| `GM` | Company-level access — leads, cases, employees, finance |
| `AM` | Personal leads only — My Leads page |
| `RECEPTION` | Reception/visitor log only |

---

## Environment Variables

```env
DATABASE_URL=          # PostgreSQL URL (cloud) — empty = SQLite
immigration_crm_SECRET=       # Flask session secret key
CLOUD_MODE=0           # 0=SQLite, 1=PostgreSQL
STORE_FILES_IN_DB=0    # 0=filesystem, 1=DB BLOB
MAX_UPLOAD_MB=20       # Max file upload size

# Email (OTP password reset)
MAIL_USERNAME=
MAIL_PASSWORD=
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=465

# Facebook Lead Ads
FB_APP_SECRET=
FB_WEBHOOK_VERIFY_TOKEN=
FB_PAGE_ACCESS_TOKEN=
```

---

## Key Features

1. **Dual Company Dashboard** — Side-by-side ABC vs XYZ performance comparison
2. **Lead Management** — Upload via Excel, manual entry, Facebook Lead Ads integration
3. **AM Allocation** — Assign leads to Assistant Managers with tracking
4. **Enrollment/Cases** — Full enrollment lifecycle (payment, filing, visa status)
5. **Employee Management** — Directory, documents, attendance, performance scoring
6. **Finance Module** — Bank accounts, transactions, reconciliation
7. **Competition System** — Real-time live scoreboard between teams
8. **Excel/PDF Export** — Any page exportable to Excel with metadata
9. **OTP Security** — Email OTP for password reset, MD approval for AM deletion
10. **Facebook Lead Ads** — Webhook integration for automatic lead capture

---

## Running Locally

```bash
# 1. Create virtual environment
python -m venv venv
venv\Scripts\activate      # Windows
source venv/bin/activate   # Linux/Mac

# 2. Install dependencies
pip install -r requirements.txt

# 3. Copy env file
cp .env.example .env

# 4. Run
python app.py
# → http://localhost:5050
```

**Default Login (local):**
| Login ID | Password | Role |
|---|---|---|
| `md.admin` | `admin123` | MD |

---

## Deploying to Render

```bash
# render.yaml already configured
# Just connect GitHub repo on render.com → New Web Service

# Start Command:
gunicorn --workers 1 --timeout 120 --bind 0.0.0.0:$PORT wsgi:app

# Environment Variables to set on Render:
CLOUD_MODE=0
STORE_FILES_IN_DB=0
immigration_crm_SECRET=your-secret-key
```

**Live URL:** https://crm-immigration.onrender.com

**Render Login Credentials:**
| Login ID | Password | Role |
|---|---|---|
| `md.admin` | `Admin@2026` | MD |
| `gm.abc` | `ABC@GM2026#` | GM - ABC |
| `gm.xyz` | `XYZ@GM2026#` | GM - XYZ |
| `reception.abc` | `Welcome@123` | Reception - ABC |
| `reception.xyz` | `Welcome@123` | Reception - XYZ |

> ⚠️ Free tier spins down after 15 min inactivity. First request may take ~50 seconds.

---

## Module Details

### `legacy_core.py`
The monolithic core file containing:
- Flask app initialization
- All database schema definitions (`init_db()`)
- All primary routes (dashboard, leads, cases, employees, etc.)
- Embedded Jinja2 templates (`_GAUR_V337_TEMPLATES`) used as cloud overrides
- Context processors, authentication, session management
- Report brand context injection

### `wsgi.py`
Cloud entry point that:
1. Calls `init_db()` to create DB schema
2. Resets user credentials and company names on startup
3. Registers all blueprints
4. Exposes `application` for Gunicorn

### `modules/report_tools/`
Post-request middleware that injects WhatsApp Share / Print / Download Excel buttons into every HTML response for MD/GM users. Excel export endpoint at `/v4/report-tools/excel`.

### `modules/fb_leads/`
Facebook Lead Ads webhook integration:
- Webhook verification (`GET /fb/webhook`)
- Lead capture (`POST /fb/webhook`)
- Settings page for token/form configuration

### `modules/security_settings/`
- User CRUD (MD only)
- OTP-based password reset via Gmail SMTP
- Security audit log

---

## AI Prompt Context

This project is an **Immigration Consultancy CRM** with:
- **Backend:** Python 3.11 + Flask 3.x (monolithic `legacy_core.py` + Blueprint modules)
- **Database:** SQLite locally, PostgreSQL on cloud (auto-detected via `DATABASE_URL`)
- **Frontend:** Server-side rendered Jinja2 templates + vanilla JS (no React/Vue)
- **Auth:** Session-based (Flask sessions), role-based access (MD > GM > AM > RECEPTION)
- **Companies:** Two companies identified by codes `SCIC` (ABC) and `WWIC` (XYZ)
- **Deployment:** Render.com free tier via `wsgi.py` + `gunicorn`
- **Key constraint:** `sqlite3.Row` objects do NOT support `.get()` — use `row["key"]` or `row["key"] or default`
- **Templates:** Both on-disk (`templates/`) AND embedded in `_GAUR_V337_TEMPLATES` dict inside `legacy_core.py` — cloud uses embedded versions
- **No frontend framework** — all UI is Jinja2 + inline CSS/JS
