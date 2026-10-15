# 🤖 Telegram Payment Reminder Bot

A simple Telegram bot that automatically sends each student a payment reminder **5 days before their deadline**, based on an Excel file uploaded by the administrator.

---

## ✨ Features

- Admin uploads a `.xlsx` file → bot schedules reminders automatically
- Students receive a Telegram message 5 days before their due date
- Duplicate prevention — each reminder is sent only once, even across restarts
- Fully local — no database, no web server, no Docker required

---

## 📊 Required Excel Format

Your `.xlsx` file must have these columns (column order doesn't matter, names are case-insensitive):

| Column | Required | Notes |
|---|---|---|
| `telegram_username` | ✅ Yes | e.g. `@K1K1Iskiki` — must start with `@` |
| `due_date` | ✅ Yes | Format: `YYYY-MM-DD` (e.g. `2026-11-05`) |
| `student_name` | ❌ Optional | Falls back to username if missing |
| `telegram_chat_id` | ❌ Optional but needed to send | Numeric ID — student gets this from `/start` |

> **Important:** Students must open the bot once and send `/start` to get their Chat ID. Without it, the bot cannot reach them.

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
EXCEL_FILE_PATH=data/students.xlsx
TIMEZONE=Asia/Yangon
REMINDER_DAYS_BEFORE=5
```

### 4. Run

```powershell
python bot.py
```

> ⚠️ **The bot must stay running** for reminders to be sent. If your computer is off at 08:00, no reminders will be sent that day.
>
> For reliable 24/7 reminders, consider running the bot on a cheap VPS or cloud VM — just copy the files, set up `.env`, and run `python bot.py` there. No Docker required.

---

## 🤖 Bot Commands

| Command | Who | Description |
|---|---|---|
| `/start` | Anyone | Shows how the bot works + displays your Chat ID |
| `/help` | Anyone | Same as `/start` |
| `/status` | Admin only | Shows loaded record count + next 5 upcoming reminders |

---

## 📁 Project Structure

```
telegram-payment-reminder/
├── bot.py                  # Main entry point
├── excel_handler.py        # Excel parsing & validation
├── reminder_scheduler.py   # Daily check + send logic
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
├── data/
│   ├── students.xlsx       # Uploaded by admin (not in git)
│   └── sent_reminders.json # Tracks sent reminders (not in git)
└── logs/
    └── bot.log             # Rotating log file
```

---

## 🔄 How It Works

1. Admin uploads `.xlsx` file to the bot in Telegram.
2. Bot validates the file and loads records into memory.
3. Every day at **08:00 Asia/Yangon**, the bot checks all records.
4. For any student whose `due_date` is exactly 5 days away, a reminder is sent.
5. Sent reminders are saved to `data/sent_reminders.json` to prevent duplicates.

---

## 🪵 Logs

Logs are written to `logs/bot.log` (rotating, max 5 MB × 3 files) and also printed to the console.

---

*Built for Ariel's Bot | Python 3.9+ | python-telegram-bot v20+*
