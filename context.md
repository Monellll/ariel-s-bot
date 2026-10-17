# Context — Multi-User Telegram Payment Reminder Bot

> **What this file is:** A single-source-of-truth summary of the project for any developer (or AI) picking it up. Read this before touching anything.

---

## 🎯 Purpose & Workflow

A simple Telegram bot supporting multiple independent users (e.g. teachers or class administrators). Each user manages their own Excel file, student roster, payment deadlines, and reminder history separately.

### 🔑 Complete Workflow & Architecture:

1. **Admin Access:**
   - The user whose numeric Telegram ID matches `ADMIN_TELEGRAM_ID` in `.env` is auto-authorized on start.

2. **Student Auto-Registration (`@username` → `chat_id` capture):**
   - Because the Telegram Bot API cannot message users purely by username without a known numeric `chat_id` (and prior interaction), the bot automatically captures chat IDs.
   - When a student opens the bot and sends `/start`, the bot extracts their Telegram username and numeric chat ID and saves it to `data/chat_id_registry.json`.
   - The student receives a confirmation:
     > *"✅ You are now registered! You will receive automatic payment reminders when your teacher schedules them."*

3. **New Teachers (Authorization Flow):**
   - When a new unauthorized user sends `/start`, the message confirms registration and also prompts:
     > *"🎓 Are you a teacher/admin? Enter the Admin Chat ID below to get full access."*
   - Entering the correct `ADMIN_TELEGRAM_ID` adds them to `data/authorized_users.json` and notifies the Admin.
   - Incorrect entries reject access.

4. **Excel Upload (Teachers):**
   - Teachers only need to provide **`telegram_username`** and **`due_date`** (optional: `student_name`).
   - `telegram_chat_id` is **NOT required** in the Excel file!
   - On upload:
     - The file is validated via `excel_handler.py`.
     - Valid records are saved to `data/{user_id}/students.xlsx` and loaded into `_user_records[user_id]`.
     - `_resolve_chat_ids()` automatically cross-references `data/chat_id_registry.json` and maps `@username` to `chat_id`.
     - The bot reports: Total records, valid count, auto-matched count, and a list of students who haven't tapped `/start` yet.

5. **Separate Reminders & Isolation:**
   - Each teacher's data is isolated in `data/{user_id}/`.
   - Each user has their own reminder history in `data/{user_id}/sent_reminders.json`.
   - Scheduled checks independently evaluate each teacher's roster.
   - Before each check, `_resolve_chat_ids()` runs again so students who tapped `/start` after the upload are picked up dynamically.
   - Reminders are sent 5 days before deadlines (or today in testing mode).
   - Warnings (unregistered students, send failures) go to the file owner (`owner_chat_id`).

6. **Commands & Menu UI:**
   - `/start`, `/help`, or `/menu`: Interactive Control Panel with inline keyboard navigation:
     - `[📊 Roster & Status]` — View loaded count, Chat ID linkage, and next 5 upcoming reminders.
     - `[⚙️ Scheduler]` — Live countdown, check interval, and full schedule queue.
     - `[📨 Check & Send Now]` — Immediately evaluate and send due reminders.
     - `[🗑️ Delete Reminders]` — Interactive list with pagination and one-tap schedule deletion (persisted back to `students.xlsx`).
     - `[📜 Sent History]` — View history and clear sent history to re-test.
     - `[🔄 Refresh Menu]` — Live refresh of timer countdown and student counters.
   - `/delete` or `/removereminder`: Directly opens the interactive reminder deletion UI.
   - `/status`, `/scheduler`, `/check`, `/history`, `/clearsent`: Direct text shortcuts that also feature the interactive menu buttons.
   - `/users`: (Admin only) View authorized teachers and auto-captured student registry stats.

---

## 📁 Project Structure

```
Ariel's Bot/
├── prompt.md                              # Original build blueprint / spec
├── context.md                             # ← YOU ARE HERE
├── test_upload.xlsx                       # Sample test Excel without chat_id column
└── telegram-payment-reminder/             # The application
    ├── bot.py                             # Entry point: Multi-user Telegram handlers + Registry + APScheduler
    ├── excel_handler.py                   # Excel parsing & validation (pure Python, no Telegram)
    ├── reminder_scheduler.py              # Daily check logic, message formatting, JSON persistence
    ├── smoke_test.py                      # Offline unit/smoke tests (mock bot, 30 test cases)
    ├── create_test_xlsx.py                # Helper script to generate sample test Excel
    ├── requirements.txt                   # Python dependencies
    ├── .env                               # Real secrets (git-ignored)
    ├── .env.example                       # Template for .env
    ├── .gitignore
    ├── README.md                          # User-facing docs & setup instructions
    ├── Procfile                           # Railway / Heroku worker process
    ├── railway.toml                       # Railway deployment config
    ├── data/
    │   ├── authorized_users.json          # Persisted authorized user IDs (git-ignored)
    │   ├── chat_id_registry.json          # Persistent registry mapping @username -> chat_id (git-ignored)
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
  - `_chat_id_registry: dict[str, int]` (mapping lowercase `@username` to `chat_id`)
- **Chat ID Auto-Resolution**:
  - `_load_chat_id_registry()` / `_save_chat_id_registry()`
  - `_register_chat_id(username, chat_id)`: registers user whenever anyone runs `/start`.
  - `_resolve_chat_ids(records)`: automatically matches records to known chat IDs.
- Handles commands (`/start`, `/help`, `/status`) with per-user scope.
- Handles document uploads (`handle_document`) storing each user's `.xlsx` in `data/{user_id}/students.xlsx`.
- Handles text input (`handle_text`) to verify the Admin Chat ID for teacher authorization.
- Runs `_run_daily_check` which resolves missing IDs, iterates each user's roster, and calls `reminder_scheduler.daily_check`.
- Startup migration: automatically migrates legacy single-user `data/students.xlsx` into `data/{ADMIN_TELEGRAM_ID}/students.xlsx` if present.

### `excel_handler.py` — Excel Parsing & Validation
- Pure Python + openpyxl. No Telegram dependency.
- Reads `.xlsx`, maps columns case-insensitively, validates each row.
- Required columns: `telegram_username`, `due_date`.
- Optional columns: `student_name` (fallback to username), `telegram_chat_id` (auto-resolved if missing), `start_date` / `class_start_date` (defaults to July 14th if omitted).
- Returns `(valid_records, invalid_count, error_messages)`.
- Skips blank rows silently. Reports bad rows without crashing.

### `reminder_scheduler.py` — Daily Check + Send Logic
- Formats dynamic bilingual reminder messages (English + Myanmar):
  - English: `"Hi!\n\nJust a friendly reminder that since we started our class on {start_date_en} , the fee for this month was due on {due_date_en}.\n\nNo worries at all, just a gentle reminder.\n\nTake your time! 🤍\n\nHello, ချယ်ချယ်တို့ အတန်းလေးက {start_date_mm} ရက်နေ့က စတာမို့ {due_date_mm} ရက်နေ့ကျ သင်တန်းကြေးလေး ပြန်သွင်းပေးရပါမယ်ရှင့်\n\nမလောပါဘူးရှင် remind လုပ်တာပါနော်🤍"`
  - Automatically formats ordinal dates (e.g. `July 14th`, `October 14th`) and Myanmar date labels (`July 14 ရက်နေ့`, `October 14 ရက်နေ့`).
- `daily_check(bot, records, owner_chat_id, ...)`:
  - For each record: skip past dates, skip if `days_until_due != reminder_days`.
  - If no `chat_id` → notify `owner_chat_id`, skip.
  - Check dedup key (`chat_id:due_date`) against user's `sent_reminders.json`.
  - Send message via `bot.send_message()`, persist key on success.
  - Report send failures to `owner_chat_id`.
- `get_upcoming_reminders()` — Used by `/status` to display the next N reminders.
- `load_sent_reminders()` / `save_sent_reminders()` — JSON persistence for dedup keys.

### `smoke_test.py` — Offline Tests
- 13 test suites (50 assertions) covering:
  - Valid rows, mixed valid/invalid, blank row skipping
  - Case-insensitive headers, missing chat_id handling
  - Date math, duplicate prevention
  - Mocked Telegram sending & error reporting
  - Multi-user isolation (separate files and histories)
  - Authorized user persistence & admin auto-authorization
  - Dynamic bilingual reminder message formatting
  - `save_excel` serialization and schedule deletion persistence
  - Strict duplicate impossibility (in-run duplicates, 3x concurrent check race conditions, atomic persistence)
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

1. **`REMINDER_DAYS_BEFORE` defaults to `0`** in `bot.py` line 51 — sends reminders for students due *today* instead of 5 days out. Change back to `"5"` for production.

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
