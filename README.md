# 🤖 Multi-User Telegram Payment Reminder Bot

A Telegram bot that automatically sends students payment reminders **5 days before their deadline**, supporting multiple teachers/administrators independently with isolated student lists and schedules.

---

## ✨ Features

- **Multi-user isolation:** Each teacher/user uploads their own `.xlsx` file, managing their own student roster and reminder schedule.
- **Admin & authorization:** Configured Admin Chat ID has immediate access. Other users enter the Admin Chat ID once to get authorized.
- **Independent schedules & duplicate prevention:** Each user has their own `sent_reminders.json` history to prevent duplicates across restarts without interfering with other users.
- **Fail-safe uploads:** If an uploaded file is invalid, existing valid data is retained.
- **Fully local & lightweight:** No database server, no web dashboard, no heavy dependencies. Uses JSON and local filesystem storage.

---

## 📊 Required Excel Format

Each user's `.xlsx` file must contain these columns (order doesn't matter, header names are case-insensitive):

| Column | Required | Notes |
|---|---|---|
| `telegram_username` | ✅ Yes | e.g. `@K1K1Iskiki` — must start with `@` |
| `due_date` | ✅ Yes | Format: `YYYY-MM-DD` (e.g. `2026-11-05`) or native Excel date |
| `student_name` | ❌ Optional | Falls back to username if missing |
| `telegram_chat_id` | ❌ Optional but needed to deliver | Numeric ID — student gets this from `/start` |

> **Note:** A Telegram bot cannot message a student unless that student has opened the bot and sent `/start`. Have students send `/start`, share their Chat ID with you, and include it in your Excel file.

---

## ⚙️ Setup (Windows 11 / PowerShell)

### 1. Prerequisites

- Python 3.9 or later installed ([python.org](https://python.org))
- A Telegram Bot Token from [@BotFather](https://t.me/BotFather)
- Your numeric Telegram User ID (use [@userinfobot](https://t.me/userinfobot) to find it)

### 2. Install

```powershell
# Navigate into the project folder
cd telegram-payment-reminder

# Create virtual environment
python -m venv venv

# Activate it
.\venv\Scripts\Activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Configure

```powershell
# Copy the example env file
Copy-Item .env.example .env

# Edit it with your real values
notepad .env
```

Fill in your `.env`:

```env
TELEGRAM_BOT_TOKEN=your_bot_token_here
ADMIN_TELEGRAM_ID=123456789          # Your numeric Telegram user ID
TIMEZONE=Asia/Yangon
REMINDER_DAYS_BEFORE=5
```

### 4. Run

```powershell
python bot.py
```

> ⚠️ **The bot must stay running** for reminders to be sent. For 24/7 reliability, run it on a small VPS or cloud VM (Railway config files `railway.toml` and `Procfile` are included).

---

## 🤖 Bot Commands & Workflow

| Action | Who | Description |
|---|---|---|
| `/start` | Admin | Welcomes the admin, shows Chat ID and status |
| `/start` | New User | Prompts to enter the Admin's Chat ID for authorization |
| Text message | Pending User | Validates Admin Chat ID. If correct, authorizes the user |
| Document (`.xlsx`) | Authorized User | Validates and imports the user's Excel file |
| `/status` | Authorized User | Shows loaded records and upcoming reminders for this user |
| `/help` | Anyone | Shows help message |

---

## 📁 Multi-User Storage Layout

```
telegram-payment-reminder/
├── bot.py
├── excel_handler.py
├── reminder_scheduler.py
├── smoke_test.py
├── data/
│   ├── authorized_users.json          # List of authorized Telegram user IDs
│   └── {user_id}/                     # One folder per user
│       ├── students.xlsx              # That user's active roster
│       └── sent_reminders.json        # That user's reminder history
└── logs/
    └── bot.log
```

---

## 🧪 Testing

Run the test suite offline (no Telegram connection needed):

```powershell
.\venv\Scripts\python.exe smoke_test.py
```

---

*Built for Ariel's Bot | Python 3.9+ | python-telegram-bot v20+*
