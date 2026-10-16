"""
bot.py
------
Main entry point for the Telegram Payment Reminder Bot.

Multi-user support: each authorized user manages their own Excel file,
student records, and reminder history independently.

Wires together:
  - python-telegram-bot (async handlers)
  - APScheduler (daily reminder job at 08:00 Asia/Yangon)
  - excel_handler  (Excel parsing/validation)
  - reminder_scheduler (daily check + send logic)

Usage:
    python bot.py
"""

import json
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
TIMEZONE: str = os.getenv("TIMEZONE", "Asia/Yangon")
REMINDER_DAYS_BEFORE: int = int(os.getenv("REMINDER_DAYS_BEFORE", "0"))  # TODO: TESTING — change back to "5" for production

# Daily check time (24-hour, Asia/Yangon)
REMINDER_HOUR = 8
REMINDER_MINUTE = 0

DATA_DIR = Path("data")
AUTHORIZED_USERS_PATH = DATA_DIR / "authorized_users.json"
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
# Per-user data paths
# ---------------------------------------------------------------------------


def _user_data_dir(user_id: int) -> Path:
    """Return the data directory for a specific user."""
    return DATA_DIR / str(user_id)


def _user_excel_path(user_id: int) -> Path:
    """Return the Excel file path for a specific user."""
    return _user_data_dir(user_id) / "students.xlsx"


def _user_sent_reminders_path(user_id: int) -> Path:
    """Return the sent_reminders.json path for a specific user."""
    return _user_data_dir(user_id) / "sent_reminders.json"


# ---------------------------------------------------------------------------
# In-memory state
# ---------------------------------------------------------------------------

# Per-user loaded records: {user_id: [records]}
_user_records: dict[int, list[dict]] = {}

# Authorized user IDs (admin is always authorized)
_authorized_users: set[int] = set()

# Users awaiting authorization (sent /start but haven't provided admin Chat ID yet)
_pending_auth: set[int] = set()

_scheduler: AsyncIOScheduler | None = None

# ---------------------------------------------------------------------------
# Authorized users persistence
# ---------------------------------------------------------------------------


def _load_authorized_users() -> set[int]:
    """Load authorized user IDs from JSON. Admin is always included."""
    users = {ADMIN_TELEGRAM_ID}
    if AUTHORIZED_USERS_PATH.exists():
        try:
            with open(AUTHORIZED_USERS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                users.update(int(uid) for uid in data)
        except (json.JSONDecodeError, OSError, ValueError) as exc:
            logger.warning("Could not load authorized_users.json: %s", exc)
    return users


def _save_authorized_users() -> None:
    """Save authorized user IDs to JSON."""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(AUTHORIZED_USERS_PATH, "w", encoding="utf-8") as f:
            json.dump(sorted(_authorized_users), f, indent=2)
    except OSError as exc:
        logger.error("Could not save authorized_users.json: %s", exc)


# ---------------------------------------------------------------------------
# Guards
# ---------------------------------------------------------------------------


def _is_admin(update: Update) -> bool:
    return update.effective_user is not None and update.effective_user.id == ADMIN_TELEGRAM_ID


def _is_authorized(update: Update) -> bool:
    return update.effective_user is not None and update.effective_user.id in _authorized_users


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/start — Welcome authorized users, auth flow for new users."""
    user = update.effective_user
    chat_id = update.effective_chat.id
    user_id = user.id
    name = user.first_name if user else "there"

    if user_id in _authorized_users:
        # Already authorized (admin or previously approved)
        record_count = len(_user_records.get(user_id, []))
        text = (
            f"👋 Hello, {name}!\n\n"
            "I'm a *Payment Reminder Bot*. I automatically send Telegram reminders "
            "to students before their payment deadline.\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🪪 *Your Chat ID is:* `{chat_id}`\n"
            f"📊 *Your loaded records:* {record_count}\n"
            "━━━━━━━━━━━━━━━━━━━━\n\n"
            "Upload an Excel file (.xlsx) to schedule reminders for your students.\n"
            "Use /status to check your current records and upcoming reminders."
        )
        await update.message.reply_text(text, parse_mode="Markdown")
    else:
        # New user — ask for admin Chat ID
        _pending_auth.add(user_id)
        text = (
            f"👋 Hello, {name}!\n\n"
            "To use this bot, you need to be authorized.\n\n"
            "Please enter the *Admin Chat ID* to continue:"
        )
        await update.message.reply_text(text, parse_mode="Markdown")


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/help — Same as /start."""
    await cmd_start(update, context)


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/status — Show the calling user's own records + next 5 reminders."""
    if not _is_authorized(update):
        await update.message.reply_text("❌ You are not authorized. Send /start to begin.")
        return

    user_id = update.effective_user.id
    records = _user_records.get(user_id, [])
    total = len(records)
    with_chat_id = sum(1 for r in records if r.get("telegram_chat_id"))
    without_chat_id = total - with_chat_id

    lines = [
        "📊 *Your Status*",
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
    """Handle .xlsx file upload from authorized users."""
    if not _is_authorized(update):
        await update.message.reply_text("❌ You are not authorized. Send /start to begin.")
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

    user_id = update.effective_user.id

    # Download to a temp file
    tmp_path = Path(tempfile.mktemp(suffix=".xlsx"))
    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(custom_path=str(tmp_path))
    except TelegramError as exc:
        logger.error("Failed to download file from user %s: %s", user_id, exc)
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
            "❌ The file contains no valid records. Your previous file (if any) has been kept.\n\n"
            + (_format_errors(errors) if errors else "")
        )
        return

    # All good — replace this user's saved Excel file
    dest = _user_excel_path(user_id)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(tmp_path), str(dest))

    # Update this user's in-memory records
    _user_records[user_id] = valid_records
    logger.info(
        "Excel updated for user %s. Valid: %d, Invalid: %d",
        user_id, len(valid_records), invalid_count,
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


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle text messages — authorization flow for pending users."""
    user_id = update.effective_user.id

    if user_id not in _pending_auth:
        return  # No pending auth — ignore

    text = (update.message.text or "").strip()

    # Validate the entered Chat ID against the admin ID
    try:
        entered_id = int(text)
    except ValueError:
        await update.message.reply_text(
            "❌ That doesn't look like a valid Chat ID. Please enter a number."
        )
        return

    if entered_id != ADMIN_TELEGRAM_ID:
        _pending_auth.discard(user_id)
        logger.warning("User %s entered wrong admin Chat ID: %s", user_id, entered_id)
        await update.message.reply_text(
            "❌ Incorrect Admin Chat ID. Access denied.\n\n"
            "Send /start to try again."
        )
        return

    # Correct — authorize the user
    _pending_auth.discard(user_id)
    _authorized_users.add(user_id)
    _save_authorized_users()
    logger.info("User %s authorized successfully.", user_id)
    await update.message.reply_text(
        "✅ *Authorized!* You can now use the bot.\n\n"
        "Upload an Excel file (.xlsx) to schedule payment reminders for your students.\n"
        "Use /status to check your records.",
        parse_mode="Markdown",
    )

    # Notify admin about new authorized user
    if user_id != ADMIN_TELEGRAM_ID and ADMIN_TELEGRAM_ID:
        try:
            tg_user = update.effective_user
            user_info = tg_user.full_name or "New user"
            username_str = f" (@{tg_user.username})" if tg_user.username else ""
            await context.bot.send_message(
                chat_id=ADMIN_TELEGRAM_ID,
                text=f"ℹ️ *New User Authorized*\n{user_info}{username_str} (Chat ID: `{user_id}`)",
                parse_mode="Markdown",
            )
        except Exception as exc:
            logger.warning("Could not notify admin of new user: %s", exc)


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
    """Run daily check for ALL authorized users with loaded records."""
    if not _user_records:
        logger.info("Daily check: no users have loaded records, skipping.")
        return

    for user_id, records in _user_records.items():
        if not records:
            continue
        logger.info("Daily check for user %s (%d records).", user_id, len(records))
        try:
            await reminder_scheduler.daily_check(
                bot=bot,
                records=records,
                owner_chat_id=user_id,
                timezone_str=TIMEZONE,
                reminder_days=REMINDER_DAYS_BEFORE,
                sent_reminders_path=_user_sent_reminders_path(user_id),
            )
        except Exception as exc:
            logger.error("Daily check failed for user %s: %s", user_id, exc)


# ---------------------------------------------------------------------------
# Startup: load existing data for all authorized users
# ---------------------------------------------------------------------------


async def on_startup(application: Application) -> None:
    """Called once when the bot starts. Load authorized users and their Excel files."""
    global _scheduler, _authorized_users

    # Migrate legacy single-user files for admin if they exist
    if ADMIN_TELEGRAM_ID:
        legacy_excel = DATA_DIR / "students.xlsx"
        admin_excel = _user_excel_path(ADMIN_TELEGRAM_ID)
        if legacy_excel.exists() and not admin_excel.exists():
            admin_excel.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(legacy_excel), str(admin_excel))
            logger.info("Migrated legacy %s to %s", legacy_excel, admin_excel)

        legacy_sent = DATA_DIR / "sent_reminders.json"
        admin_sent = _user_sent_reminders_path(ADMIN_TELEGRAM_ID)
        if legacy_sent.exists() and not admin_sent.exists():
            admin_sent.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(legacy_sent), str(admin_sent))
            logger.info("Migrated legacy %s to %s", legacy_sent, admin_sent)

    # Load authorized users (admin is always included)
    _authorized_users = _load_authorized_users()
    logger.info("Startup: %d authorized user(s) loaded.", len(_authorized_users))

    # Load each authorized user's Excel file (if it exists)
    for user_id in _authorized_users:
        excel_path = _user_excel_path(user_id)
        if excel_path.exists():
            try:
                records, invalid, _ = excel_handler.load_excel(str(excel_path))
                _user_records[user_id] = records
                logger.info(
                    "Startup: loaded %d records for user %s (%d invalid).",
                    len(records), user_id, invalid,
                )
            except Exception as exc:
                logger.warning(
                    "Startup: could not load Excel for user %s: %s", user_id, exc
                )

    # Set up APScheduler (event loop is running here)
    # TODO: TESTING MODE — runs every 3 minutes. Revert to cron for production.
    _scheduler = AsyncIOScheduler(timezone=ZoneInfo(TIMEZONE))
    _scheduler.add_job(
        _run_daily_check,
        trigger="interval",
        minutes=3,
        kwargs={"bot": application.bot},
        id="daily_reminder_check",
        replace_existing=True,
    )
    _scheduler.start()
    logger.info("Scheduler started. TESTING MODE: checking every 3 minutes.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    setup_logging()

    if not BOT_TOKEN:
        logger.critical("TELEGRAM_BOT_TOKEN is not set. Exiting.")
        raise SystemExit(1)
    if not ADMIN_TELEGRAM_ID:
        logger.critical("ADMIN_TELEGRAM_ID is not set. Exiting.")
        raise SystemExit(1)

    logger.info("Starting Ariel's Payment Reminder Bot…")
    logger.info(
        "Admin ID: %s | Timezone: %s | Reminder days: %s",
        ADMIN_TELEGRAM_ID, TIMEZONE, REMINDER_DAYS_BEFORE,
    )

    # Build the application (scheduler starts inside on_startup)
    app = Application.builder().token(BOT_TOKEN).post_init(on_startup).build()

    # Register handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(
        MessageHandler(filters.Document.ALL & filters.ChatType.PRIVATE, handle_document)
    )
    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_text
        )
    )

    # Run (blocks until Ctrl+C)
    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES)
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot shutting down…")
    finally:
        if _scheduler and _scheduler.running:
            _scheduler.shutdown(wait=False)
        logger.info("Bot stopped cleanly.")


if __name__ == "__main__":
    main()
