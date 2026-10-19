# Context — Multi-User Telegram Payment Reminder Bot

> **What this file is:** A single-source-of-truth summary of the project for any developer (or AI) picking it up. Read this before touching anything.

---

## 🎯 Purpose & Workflow

A clean, friendly Telegram bot supporting multiple independent users (e.g. teachers or class administrators). Each user manages their own Excel file, student roster, payment deadlines, and reminder history separately.

### 🔑 Complete Workflow & Architecture:

1. **Teacher Access & Single Permanent Administrator:**
   - There is strictly **ONE Permanent Administrator**, defined via `ADMIN_TELEGRAM_ID` on Railway / `.env` (with `7462033455` as fallback permanent admin).
   - **Exclusive Administrator Rights:**
     - Only the ONE Permanent Administrator can perform actions regarding teachers:
       - `/authorize <user_id>` or `/authorize @username` — Grant teacher privileges directly by Telegram ID or username.
       - `/revoke <user_id>` — Revoke teacher privileges (the permanent administrator cannot be revoked).
       - `/teachers` — View the permanent administrator and all authorized teachers.
       - `[👥 Manage Teachers]` — Dashboard button visible exclusively to the Permanent Administrator with one-tap teacher revocation.
   - **Teacher Access (No Admin Actions):**
     - Teachers can gain access in two ways:
       1. **Entering Admin ID:** Sending the permanent admin ID into the chat (as plain text or via `/auth <Admin_ID>`).
       2. **Authorized by Admin:** Being authorized directly by the permanent administrator via `/authorize`.
     - In both cases, users receive **Teacher access only** (their Dashboard unlocks with student management tools).
     - Authorized teachers **CANNOT** take any action about other teachers (they cannot view `/teachers`, cannot `/authorize`, cannot `/revoke`, and do not have the `[👥 Manage Teachers]` button).
   - **Persistence:** Authorized teachers (`authorized_users.json`) are persisted to disk to survive bot restarts.
   - **Authorized Teachers:** Can manage their own student rosters, due dates, Excel uploads, student deletions, and payment reminders.

2. **Clean Student Experience:**
   - When a student opens the bot and sends `/start`, the bot captures their Telegram username and numeric chat ID automatically to `data/chat_id_registry.json`.
   - The student receives a warm, friendly confirmation without any technical jargon or admin prompts:
     > *"👋 Hello, [Name]! Welcome to Ariel's English Class! 🤍 You're all set! You will receive a gentle reminder here whenever your monthly tuition fee is coming up. No further action is needed from you. Have a wonderful day! ✨"*

3. **Streamlined Teacher Control Panel (Modern Card UX):**
   - When an authorized teacher sends `/start`, `/menu`, or `/roster`:
     - Clean visual sections: Personal Welcome, Class Overview (Total, Ready ✅, Waiting ⚠️), and Automation Schedule.
     - Action buttons:
       - `[📋 Student Roster]` — View list of students with clean two-line tree layout and relative due dates.
       - `[📨 Send Reminders Now]` — Instantly check and send reminders for due students.
       - `[🗑️ Remove Student]` — Interactive paginated deletion with one-tap removal and safety confirmation card.
       - `[📖 How to Use]` — Step-by-step Excel formatting and student onboarding guide.
       - `[👥 Manage Teachers]` — Admin-only teacher management panel.
       - `[🔄 Refresh Dashboard]` — Refreshes the dashboard stats.

4. **Excel Upload (Teachers):**
   - Teachers only need to provide **`student_name`**, **`telegram_username`**, and **`due_date`** (format: YYYY-MM-DD).
   - `telegram_chat_id` is **NOT required** in the Excel file!
   - On upload:
     - The file is validated via `excel_handler.py`.
     - Valid records are saved to `data/{user_id}/students.xlsx` and loaded into `_user_records[user_id]`.
     - `_resolve_chat_ids()` cross-references `data/chat_id_registry.json` and maps `@username` to `chat_id`.
     - The bot displays an instant **Upload Summary Card** showing total students, linked count, and names of pending students.

5. **Automated Reminders & Flexible Window:**
   - **Background Scheduler:** Dynamically configured via `CHECK_INTERVAL_MINUTES`:
     - If `CHECK_INTERVAL_MINUTES=0` (or not set): Runs in **Production Mode** once daily at 8:00 AM (`Asia/Yangon`).
     - If `CHECK_INTERVAL_MINUTES > 0`: Runs in **Test Mode** (e.g., checks every 3 minutes).
   - **Reminder Window:** Reminds students whose fee is due **within the reminder window (0 to `REMINDER_DAYS_BEFORE` days, default 5)** so students due today or within 5 days are never skipped.
   - **Bilingual Message:** Sends a polite bilingual reminder message (English + Myanmar).
   - **Teacher Alert on Auto-Send:** When the background job delivers reminders, it automatically notifies the teacher with the student names.
   - **Zero Duplicate Spam:** Canonical key `{chat_id}:{due_date}` saved atomically to `sent_reminders.json` prevents duplicate messages.

6. **Bulletproof Technical Resilience:**
   - **HTML Mode (`parse_mode="HTML"`)**: Prevents Telegram entity parsing errors caused by underscores in usernames (e.g. `@thiri_test`, `@myat_min`) or brackets in messages.
   - **`_safe_edit_message` Helper**: Calls `edit_message_text` with adaptive fallback across `query.edit_message_text`, `query.edit_text`, and `query.message.edit_text`, guaranteeing buttons always update the screen.
   - **Telegram Native Menu Button**: Commands registered with Telegram API (`set_my_commands`) so the bottom-left `[ ☰ Menu ]` button always works.
   - **`run_polling(drop_pending_updates=True)`**: Cleans out old stuck backlogs on startup.
   - **Single-Worker Cloud Deployment**: Uses `railway.toml` with `startCommand = "python bot.py"` (redundant `Procfile` removed) to eliminate duplicate instance conflicts (HTTP 409).
   - **Automated Tests:** 15 test suites and 60 assertions in `smoke_test.py` covering excel parsing, dates, deduplication, multi-user isolation, admin rights, and callbacks.

---

## 📁 Project Structure

```
Ariel's Bot/
├── prompt.md                              # Original build blueprint / spec
├── context.md                             # Project architecture & context summary
├── TESTING_GUIDE.md                       # Comprehensive 7-step live testing roadmap
├── test_upload.xlsx                       # Sample test Excel generated by create_test_xlsx.py
└── telegram-payment-reminder/             # The application
    ├── bot.py                             # Entry point: Multi-user Telegram handlers + Friendly UI + APScheduler
    ├── excel_handler.py                   # Excel parsing, validation & saving (pure Python, openpyxl)
    ├── reminder_scheduler.py              # Daily check logic, dynamic bilingual formatting, JSON persistence + lock
    ├── smoke_test.py                      # Offline unit/smoke tests (15 test suites, 60 assertions)
    ├── create_test_xlsx.py                # Helper script to generate sample test Excel
    ├── TESTING_GUIDE.md                   # Live self-testing guide
    ├── context.md                         # Synchronized project context
    ├── requirements.txt                   # Python dependencies
    ├── .env                               # Real secrets (git-ignored)
    ├── .env.example                       # Template for .env
    ├── .gitignore
    ├── README.md                          # User-facing docs & setup instructions
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

## ⚙️ Environment Variables

| Variable | Default | Description |
|---|---|---|
| `TELEGRAM_BOT_TOKEN` | *(required)* | Bot token from @BotFather |
| `ADMIN_TELEGRAM_ID` | *(required)* | Numeric Telegram user ID of the ONE permanent admin |
| `TIMEZONE` | `Asia/Yangon` | Timezone for date math and scheduling |
| `REMINDER_DAYS_BEFORE` | `5` | Maximum days before due date to trigger reminder |
| `CHECK_INTERVAL_MINUTES` | `0` | Scheduler mode: `0` = Daily 8:00 AM production cron; `> 0` = Test interval every X minutes |

---

## 🛠️ Deployment & Remote

* **Git Repository:** `https://github.com/Monellll/ariel-s-bot.git` (branch `main`)
* **Cloud Host:** Railway (`railway.toml: startCommand = "python bot.py"`)
* **Single-Instance Rule:** Keep local bot process stopped while Railway is running to avoid Telegram `HTTP 409 Conflict`.

---

*Last updated: 2026-10-18*
