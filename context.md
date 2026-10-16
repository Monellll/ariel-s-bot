# Context — Multi-User Telegram Payment Reminder Bot

> **What this file is:** A single-source-of-truth summary of the project for any developer (or AI) picking it up. Read this before touching anything.

---

## 🎯 Purpose & Workflow

A simple Telegram bot supporting multiple independent users (e.g. teachers or class administrators). Each user manages their own Excel file, student roster, payment deadlines, and reminder history separately.

### 🔑 User Workflow:
1. **Admin Access:**
   - The user whose numeric Telegram ID matches `ADMIN_TELEGRAM_ID` in `.env` is auto-authorized on start.
2. **New Users (Teachers):**
   - When a new user opens the bot and sends `/start`, the bot detects they are unauthorized and asks for the **Admin's Telegram Chat ID**.
3. **Authorization:**
   - The user enters the numeric Admin Chat ID.
   - If correct: user is added to `data/authorized_users.json`, and the bot notifies the Admin.
   - If incorrect: access is rejected and they are prompted to try `/start` again.
4. **Excel Upload:**
   - Any authorized user can upload their `.xlsx` file.
   - The bot validates the document using `excel_handler.py`.
   - If valid: saved to `data/{user_id}/students.xlsx`, loaded into `_user_records[user_id]`.
   - If invalid: previous valid file/records are preserved; errors are reported to that user.
5. **Separate Reminders & Isolation:**
   - Each user has their own records in `_user_records[user_id]`.
   - Each user has their own reminder history in `data/{user_id}/sent_reminders.json`.
   - Daily check evaluates each user's records independently. Reminders are sent 5 days before deadlines.
   - Any warnings (students missing chat IDs, send failures) are sent directly to that specific user (`owner_chat_id`).
6. **Commands:**
   - `/start` or `/help`: Welcome message, Chat ID, and loaded record count (or auth prompt for new users).
   - `/status`: Shows current user's loaded student count and next 5 upcoming reminders (plus authorized users count for admin).

---

## 📁 Project Structure

```
Ariel's Bot/
├── prompt.md                              # Original build blueprint / spec
├── context.md                             # ← YOU ARE HERE
└── telegram-payment-reminder/             # The application
    ├── bot.py                             # Entry point: Multi-user Telegram handlers + APScheduler
    ├── excel_handler.py                   # Excel parsing & validation (pure Python, no Telegram)
    ├── reminder_scheduler.py              # Daily check logic, message formatting, JSON persistence
    ├── smoke_test.py                      # Offline unit/smoke tests (mock bot, 30 test cases)
    ├── requirements.txt                   # Python dependencies
    ├── .env                               # Real secrets (git-ignored)
    ├── .env.example                       # Template for .env
    ├── .gitignore
    ├── README.md                          # User-facing docs & setup instructions
    ├── Procfile                           # Railway / Heroku worker process
    ├── railway.toml                       # Railway deployment config
    ├── data/
    │   ├── authorized_users.json          # Persisted authorized user IDs (git-ignored)
    │   ├── {user_id}/                     # Per-user isolated directory (git-ignored)
    │   │   ├── students.xlsx              # That user's student roster
    │   │   └── sent_reminders.json        # That user's sent reminder history
    │   └── test_students.xlsx             # Sample test data
    ├── logs/
    │   └── bot.log                        # Rotating log file (git-ignored)
    └── venv/                              # Python virtual environment (git-ignored)
```

---

## 🧩 Module Responsibilities

### `bot.py` — Main Entry Point
- Loads env vars via `python-dotenv`.
- Sets up rotating file + console logging.
- Manages multi-user in-memory states:
  - `_user_records: dict[int, list[dict]]`
  - `_authorized_users: set[int]`
  - `_pending_auth: set[int]`
- Handles commands (`/start`, `/help`, `/status`) with per-user scope.
- Handles document uploads (`handle_document`) storing each user's `.xlsx` in `data/{user_id}/students.xlsx`.
- Handles text input (`handle_text`) to verify the Admin Chat ID for authorization.
- Runs `_run_daily_check` which iterates each user's roster and calls `reminder_scheduler.daily_check`.
- Startup migration: automatically migrates legacy single-user `data/students.xlsx` into `data/{ADMIN_TELEGRAM_ID}/students.xlsx` if present.

### `excel_handler.py` — Excel Parsing & Validation
- Pure Python + openpyxl. No Telegram dependency.
- Reads `.xlsx`, maps columns case-insensitively, validates each row.
- Required columns: `telegram_username`, `due_date`.
- Optional columns: `student_name` (fallback to username), `telegram_chat_id` (needed for delivery).
- Returns `(valid_records, invalid_count, error_messages)`.
- Skips blank rows silently. Reports bad rows without crashing.

### `reminder_scheduler.py` — Daily Check + Send Logic
- `daily_check(bot, records, owner_chat_id, ...)`:
  - For each record: skip past dates, skip if `days_until_due != reminder_days`.
  - If no `chat_id` → notify `owner_chat_id`, skip.
  - Check dedup key (`chat_id:due_date`) against user's `sent_reminders.json`.
  - Send message via `bot.send_message()`, persist key on success.
  - Report send failures to `owner_chat_id`.
- `get_upcoming_reminders()` — Used by `/status` to display the next N reminders.
- `load_sent_reminders()` / `save_sent_reminders()` — JSON persistence for dedup keys.

### `smoke_test.py` — Offline Tests
- 10 test suites (30 assertions) covering:
  - Valid rows, mixed valid/invalid, blank row skipping
  - Case-insensitive headers, missing chat_id handling
  - Date math, duplicate prevention
  - Mocked Telegram sending & error reporting
  - Multi-user isolation (separate files and histories)
  - Authorized user persistence & admin auto-authorization
- Run with: `.\venv\Scripts\python.exe smoke_test.py`

---

## ⚙️ Environment Variables

| Variable | Default | Description |
|---|---|---|
| `TELEGRAM_BOT_TOKEN` | *(required)* | Bot token from @BotFather |
| `ADMIN_TELEGRAM_ID` | *(required)* | Numeric Telegram user ID of the admin |
| `TIMEZONE` | `Asia/Yangon` | Timezone for date math and scheduling |
| `REMINDER_DAYS_BEFORE` | `5` | How many days before due date to send a reminder |

---

## 🛠️ Tech Stack

| Library | Purpose |
|---|---|
| `python-telegram-bot` ≥20.0 | Async Telegram bot framework |
| `openpyxl` ≥3.1.0 | Read `.xlsx` files |
| `APScheduler` ≥3.10.0 | Scheduled daily job |
| `python-dotenv` ≥1.0.0 | Load `.env` file |
| `zoneinfo` (stdlib) | Timezone handling (Python 3.9+) |

---

## 🚧 Current State & Known TODOs

> [!WARNING]
> **The bot is currently in TESTING MODE with two temporary overrides:**

1. **`REMINDER_DAYS_BEFORE` defaults to `0`** in `bot.py` line 50 — sends reminders for students due *today* instead of 5 days out. Change back to `"5"` for production.

2. **Scheduler runs every 3 minutes** (interval trigger) instead of once daily at 08:00 via cron in `bot.py`. Revert to cron for production:
```python
_scheduler.add_job(
    _run_daily_check,
    trigger="cron",
    hour=8,
    minute=0,
    kwargs={"bot": application.bot},
    id="daily_reminder_check",
    replace_existing=True,
)
```

---

*Last updated: 2026-10-16*
