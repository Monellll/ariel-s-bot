"""
reminder_scheduler.py
---------------------
Daily check logic for the Telegram Payment Reminder Bot.

Responsibilities:
  - Load / save sent_reminders.json (duplicate prevention)
  - Format reminder messages
  - Check which students are due in REMINDER_DAYS_BEFORE days
  - Send reminders via the Telegram bot
  - Report failures to admin
"""

import datetime
import json
import logging
import os
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

from telegram import Bot
from telegram.error import TelegramError

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants (overridden by env vars via bot.py)
# ---------------------------------------------------------------------------
SENT_REMINDERS_PATH = Path("data/sent_reminders.json")

REMINDER_TEMPLATE = """\
Hello {student_name}! 👋

This is a friendly reminder that your payment deadline is {due_date}, which is {days} days away.

Please make your payment before the deadline.

If you have already paid, please contact your class administrator.

Thank you!\
"""

# ---------------------------------------------------------------------------
# Sent Reminders JSON persistence
# ---------------------------------------------------------------------------


def load_sent_reminders(path: Path = SENT_REMINDERS_PATH) -> set[str]:
    """
    Load already-sent reminder keys from JSON.
    Returns an empty set if the file doesn't exist or is malformed.
    """
    if not path.exists():
        return set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return set(data)
        logger.warning("sent_reminders.json has unexpected format; resetting.")
        return set()
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Could not load sent_reminders.json: %s", exc)
        return set()


def save_sent_reminders(keys: set[str], path: Path = SENT_REMINDERS_PATH) -> None:
    """
    Save sent reminder keys to JSON (sorted list for readability).
    """
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(sorted(keys), f, indent=2)
    except OSError as exc:
        logger.error("Could not save sent_reminders.json: %s", exc)


# ---------------------------------------------------------------------------
# Message formatting
# ---------------------------------------------------------------------------


def format_reminder_message(
    student_name: str, due_date: datetime.date, reminder_days: int
) -> str:
    """Format the reminder message from the template."""
    due_date_str = due_date.strftime("%B %d, %Y")  # e.g. "November 05, 2026"
    return REMINDER_TEMPLATE.format(
        student_name=student_name,
        due_date=due_date_str,
        days=reminder_days,
    )


# ---------------------------------------------------------------------------
# Upcoming reminders (for /status command)
# ---------------------------------------------------------------------------


def get_upcoming_reminders(
    records: list[dict],
    timezone_str: str,
    reminder_days: int,
    n: int = 5,
) -> list[dict]:
    """
    Return the next `n` records sorted by days_until_due (ascending),
    where days_until_due > 0.  Used for the /status command display.
    """
    tz = ZoneInfo(timezone_str)
    today = datetime.datetime.now(tz).date()

    upcoming = []
    for rec in records:
        due_date: Optional[datetime.date] = rec.get("due_date")
        if not due_date:
            continue
        days_until = (due_date - today).days
        if days_until > 0:
            upcoming.append({**rec, "days_until_due": days_until})

    upcoming.sort(key=lambda r: r["days_until_due"])
    return upcoming[:n]


# ---------------------------------------------------------------------------
# Core daily check
# ---------------------------------------------------------------------------


async def daily_check(
    bot: Bot,
    records: list[dict],
    admin_chat_id: int,
    timezone_str: str,
    reminder_days: int,
    sent_reminders_path: Path = SENT_REMINDERS_PATH,
) -> None:
    """
    Main daily reminder check.

    For each record:
      - Skip if no chat_id (notify admin).
      - Skip if due_date already passed.
      - If days_until_due == reminder_days and not already sent → send + save key.
    """
    tz = ZoneInfo(timezone_str)
    today = datetime.datetime.now(tz).date()

    sent_keys = load_sent_reminders(sent_reminders_path)
    newly_sent: list[str] = []
    skipped_no_chat_id: list[dict] = []
    failed: list[tuple[dict, str]] = []

    logger.info(
        "Daily check started. Today: %s. Records: %d. Loaded sent keys: %d",
        today,
        len(records),
        len(sent_keys),
    )

    for rec in records:
        username: str = rec.get("telegram_username", "Unknown")
        student_name: str = rec.get("student_name", username)
        due_date: Optional[datetime.date] = rec.get("due_date")
        chat_id: Optional[int] = rec.get("telegram_chat_id")

        # Skip if no due_date (shouldn't happen after validation, but be safe)
        if due_date is None:
            logger.warning("Skipping %s: no due_date.", username)
            continue

        # Skip past due dates
        if due_date < today:
            continue

        days_until_due = (due_date - today).days

        # Only send when exactly REMINDER_DAYS_BEFORE
        if days_until_due != reminder_days:
            continue

        # No chat_id → report to admin, skip
        if chat_id is None:
            skipped_no_chat_id.append(rec)
            logger.warning(
                "No chat_id for %s (due %s) — cannot send reminder.", username, due_date
            )
            continue

        reminder_key = f"{chat_id}:{due_date}"

        # Duplicate prevention
        if reminder_key in sent_keys:
            logger.info("Already sent reminder for key %s — skipping.", reminder_key)
            continue

        # Send the reminder
        message = format_reminder_message(student_name, due_date, reminder_days)
        try:
            await bot.send_message(chat_id=chat_id, text=message)
            sent_keys.add(reminder_key)
            newly_sent.append(reminder_key)
            logger.info("Reminder sent to %s (chat_id=%s).", username, chat_id)
        except TelegramError as exc:
            failed.append((rec, str(exc)))
            logger.error(
                "Failed to send reminder to %s (chat_id=%s): %s",
                username,
                chat_id,
                exc,
            )

    # Persist after all sends
    if newly_sent:
        save_sent_reminders(sent_keys, sent_reminders_path)
        logger.info("Saved %d new reminder keys.", len(newly_sent))

    # Report skipped (no chat_id) to admin
    if skipped_no_chat_id:
        lines = ["⚠️ Students missing Chat ID (reminders NOT sent):\n"]
        for rec in skipped_no_chat_id:
            lines.append(
                f"• {rec.get('student_name')} ({rec.get('telegram_username')}) "
                f"— due {rec.get('due_date')}"
            )
        await _notify_admin(bot, admin_chat_id, "\n".join(lines))

    # Report failures to admin
    if failed:
        lines = ["❌ Failed to send reminders:\n"]
        for rec, err in failed:
            lines.append(
                f"• {rec.get('student_name')} ({rec.get('telegram_username')}) "
                f"— {err}"
            )
        await _notify_admin(bot, admin_chat_id, "\n".join(lines))

    logger.info(
        "Daily check complete. Sent: %d, Skipped (no chat_id): %d, Failed: %d",
        len(newly_sent),
        len(skipped_no_chat_id),
        len(failed),
    )


# ---------------------------------------------------------------------------
# Admin notification helper
# ---------------------------------------------------------------------------


async def _notify_admin(bot: Bot, admin_chat_id: int, text: str) -> None:
    """Send a message to the admin, logging any failure silently."""
    try:
        await bot.send_message(chat_id=admin_chat_id, text=text)
    except TelegramError as exc:
        logger.error("Could not notify admin (chat_id=%s): %s", admin_chat_id, exc)
