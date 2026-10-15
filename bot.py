"""
bot.py
------
Main entry point for the Telegram Payment Reminder Bot.

Wires together:
  - python-telegram-bot (async handlers)
  - APScheduler (daily reminder job at 08:00 Asia/Yangon)
  - excel_handler  (Excel parsing/validation)
  - reminder_scheduler (daily check + send logic)

Usage:
    python bot.py
"""

import asyncio
import logging
import os
import shutil
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path
from zoneinfo import ZoneInfo

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from dotenv import load_dotenv
from telegram import Update
from telegram.error import TelegramError
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

import excel_handler
import reminder_scheduler

# ---------------------------------------------------------------------------
# Load environment
# ---------------------------------------------------------------------------
load_dotenv()

BOT_TOKEN: str = os.getenv("TELEGRAM_BOT_TOKEN", "")
ADMIN_TELEGRAM_ID: int = int(os.getenv("ADMIN_TELEGRAM_ID", "0"))
EXCEL_FILE_PATH: str = os.getenv("EXCEL_FILE_PATH", "data/students.xlsx")
TIMEZONE: str = os.getenv("TIMEZONE", "Asia/Yangon")
REMINDER_DAYS_BEFORE: int = int(os.getenv("REMINDER_DAYS_BEFORE", "5"))

# Daily check time (24-hour, Asia/Yangon)
REMINDER_HOUR = 8
REMINDER_MINUTE = 0

SENT_REMINDERS_PATH = Path("data/sent_reminders.json")
LOGS_DIR = Path("logs")

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------


def setup_logging() -> None:
    """Configure rotating file + console logging."""
    LOGS_DIR.mkdir(exist_ok=True)
    log_file = LOGS_DIR / "bot.log"

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(file_handler)
    root.addHandler(console_handler)

    # Quieten noisy libraries
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("apscheduler").setLevel(logging.WARNING)


logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# In-memory state (loaded records, refreshed on Excel upload)
# ---------------------------------------------------------------------------
_loaded_records: list[dict] = []


def _get_records() -> list[dict]:
    return _loaded_records


def _set_records(records: list[dict]) -> None:
    global _loaded_records
    _loaded_records = records


# ---------------------------------------------------------------------------
# Admin guard
# ---------------------------------------------------------------------------


def _is_admin(update: Update) -> bool:
    return update.effective_user is not None and update.effective_user.id == ADMIN_TELEGRAM_ID


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/start — Explain bot + show caller's Chat ID."""
    user = update.effective_user
    chat_id = update.effective_chat.id
    name = user.first_name if user else "there"

    text = (
        f"👋 Hello, {name}!\n\n"
        "I'm a *Payment Reminder Bot*. I automatically send Telegram reminders "
        "to students 5 days before their payment deadline.\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"🪪 *Your Chat ID is:* `{chat_id}`\n"
        "━━━━━━━━━━━━━━━━━━━━\n\n"
        "Please share this Chat ID with your administrator so they can add it "
        "to the Excel file. Once added, you will receive an automatic reminder "
        "before your payment is due.\n\n"
        "If you are the administrator, upload the Excel file (.xlsx) here to schedule reminders."
    )
    await update.message.reply_text(text, parse_mode="Markdown")


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/help — Same as /start."""
    await cmd_start(update, context)


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/status — Admin only. Show loaded records + next 5 reminders."""
    if not _is_admin(update):
        await update.message.reply_text("❌ This command is for the administrator only.")
        return

    records = _get_records()
    total = len(records)
    with_chat_id = sum(1 for r in records if r.get("telegram_chat_id"))
    without_chat_id = total - with_chat_id

    lines = [
        "📊 *Bot Status*",
        f"Loaded records: {total}",
        f"✅ With Chat ID: {with_chat_id}",
        f"⚠️ Missing Chat ID: {without_chat_id}",
    ]

    if total == 0:
        lines.append("\n_No Excel file loaded yet. Upload a .xlsx file to begin._")
    else:
        upcoming = reminder_scheduler.get_upcoming_reminders(
            records, TIMEZONE, REMINDER_DAYS_BEFORE, n=5
        )
        if upcoming:
            lines.append("\n📅 *Next upcoming reminders:*")
            for rec in upcoming:
                lines.append(
                    f"• {rec['student_name']} ({rec['telegram_username']}) "
                    f"— due {rec['due_date']} (in {rec['days_until_due']} days)"
                )
        else:
            lines.append("\n_No upcoming reminders in the next period._")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle .xlsx file upload from admin."""
    if not _is_admin(update):
        await update.message.reply_text("❌ Only the administrator can upload files.")
        return

    doc = update.message.document
    if not doc:
        return

    # Reject non-xlsx files
    fname = doc.file_name or ""
    if not fname.lower().endswith(".xlsx"):
        await update.message.reply_text(
            "❌ Please upload a valid *.xlsx* file. Other formats are not supported.",
            parse_mode="Markdown",
        )
        return

    await update.message.reply_text("⏳ Receiving file and validating…")

    # Download to a temp file
    tmp_path = Path(tempfile.mktemp(suffix=".xlsx"))
    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(custom_path=str(tmp_path))
    except TelegramError as exc:
        logger.error("Failed to download uploaded file: %s", exc)
        await update.message.reply_text(f"❌ Could not download the file: {exc}")
        return

    # Validate the new file BEFORE replacing the old one
    try:
        valid_records, invalid_count, errors = excel_handler.load_excel(str(tmp_path))
    except (FileNotFoundError, ValueError) as exc:
        tmp_path.unlink(missing_ok=True)
        await update.message.reply_text(f"❌ Validation failed:\n{exc}")
        return

    if len(valid_records) == 0:
        tmp_path.unlink(missing_ok=True)
        await update.message.reply_text(
            "❌ The file contains no valid records. The previous file (if any) has been kept.\n\n"
            + (_format_errors(errors) if errors else "")
        )
        return

    # All good — replace the saved Excel file
    dest = Path(EXCEL_FILE_PATH)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(tmp_path), str(dest))

    # Update in-memory records
    _set_records(valid_records)
    logger.info(
        "Excel file updated. Valid: %d, Invalid: %d", len(valid_records), invalid_count
    )

    total = len(valid_records) + invalid_count
    reply_lines = [
        "✅ *Excel file imported successfully!*",
        f"Total records: {total}",
        f"Valid records: {len(valid_records)}",
        f"Invalid records: {invalid_count}",
        "",
        "Automatic payment reminders are now scheduled.",
    ]
    if errors:
        reply_lines += ["", "⚠️ *Issues found (rows skipped):*", _format_errors(errors)]

    await update.message.reply_text("\n".join(reply_lines), parse_mode="Markdown")


def _format_errors(errors: list[str]) -> str:
    """Format validation errors into a readable string (capped at 10 lines)."""
    shown = errors[:10]
    text = "\n".join(f"• {e}" for e in shown)
    if len(errors) > 10:
        text += f"\n… and {len(errors) - 10} more issues."
    return text


# ---------------------------------------------------------------------------
# Scheduled daily job
# ---------------------------------------------------------------------------


async def _run_daily_check(bot) -> None:
    """Wrapper called by APScheduler."""
    records = _get_records()
    if not records:
        logger.info("Daily check: no records loaded, skipping.")
        return
    await reminder_scheduler.daily_check(
        bot=bot,
        records=records,
        admin_chat_id=ADMIN_TELEGRAM_ID,
        timezone_str=TIMEZONE,
        reminder_days=REMINDER_DAYS_BEFORE,
        sent_reminders_path=SENT_REMINDERS_PATH,
    )


# ---------------------------------------------------------------------------
# Startup: load existing Excel (if any)
# ---------------------------------------------------------------------------


async def on_startup(application: Application) -> None:
    """Called once when the bot starts. Load existing Excel if present."""
    excel_path = Path(EXCEL_FILE_PATH)
    if excel_path.exists():
        try:
            records, invalid, _ = excel_handler.load_excel(str(excel_path))
            _set_records(records)
            logger.info(
                "Startup: loaded %d records from existing Excel (%d invalid).",
                len(records),
                invalid,
            )
        except Exception as exc:
            logger.warning("Startup: could not load existing Excel: %s", exc)
    else:
        logger.info("Startup: no existing Excel file found at %s.", excel_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    setup_logging()

    if not BOT_TOKEN:
        logger.critical("TELEGRAM_BOT_TOKEN is not set in .env. Exiting.")
        raise SystemExit(1)
    if not ADMIN_TELEGRAM_ID:
        logger.critical("ADMIN_TELEGRAM_ID is not set in .env. Exiting.")
        raise SystemExit(1)

    logger.info("Starting Ariel's Payment Reminder Bot…")
    logger.info("Admin ID: %s | Timezone: %s | Reminder days: %s", ADMIN_TELEGRAM_ID, TIMEZONE, REMINDER_DAYS_BEFORE)

    # Build the application
    app = Application.builder().token(BOT_TOKEN).post_init(on_startup).build()

    # Register handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(
        MessageHandler(filters.Document.ALL & filters.ChatType.PRIVATE, handle_document)
    )

    # Set up APScheduler
    scheduler = AsyncIOScheduler(timezone=ZoneInfo(TIMEZONE))
    scheduler.add_job(
        _run_daily_check,
        trigger="cron",
        hour=REMINDER_HOUR,
        minute=REMINDER_MINUTE,
        kwargs={"bot": app.bot},
        id="daily_reminder_check",
        replace_existing=True,
    )
    scheduler.start()
    logger.info(
        "Scheduler started. Daily check at %02d:%02d %s.",
        REMINDER_HOUR,
        REMINDER_MINUTE,
        TIMEZONE,
    )

    # Run (blocks until Ctrl+C)
    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES)
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot shutting down…")
    finally:
        if scheduler.running:
            scheduler.shutdown(wait=False)
        logger.info("Bot stopped cleanly.")


if __name__ == "__main__":
    main()
