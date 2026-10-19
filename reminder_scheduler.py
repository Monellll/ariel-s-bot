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

import asyncio
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
# Concurrency & Deduplication Guards
# ---------------------------------------------------------------------------
_daily_check_lock = asyncio.Lock()
_memory_sent_cache: dict[str, set[str]] = {}

# ---------------------------------------------------------------------------
# Constants (overridden by env vars via bot.py)
# ---------------------------------------------------------------------------
SENT_REMINDERS_PATH = Path("data/sent_reminders.json")

REMINDER_TEMPLATE_EN = """\
Hi!

Just a friendly reminder that the fee for the class is due on {due_date_en}.

No worries at all, just a gentle reminder.

Take your time! 🤍\
"""

REMINDER_TEMPLATE_MM = """\
Hello, ချယ်ချယ်တို့ရဲ့သင်တန်းလေးကိုဒီလအတွက်  {due_date_mm} ရက်နေ့ နောက်ဆုံးထားပြီး ပေးသွင်းရပါမယ်ရှင့်။

မလောပါဘူးရှင် remind လုပ်တာပါနော်🤍\
"""

REMINDER_TEMPLATE_BOTH = f"""\
{REMINDER_TEMPLATE_EN}

{REMINDER_TEMPLATE_MM}\
"""

# Default template alias
REMINDER_TEMPLATE = REMINDER_TEMPLATE_BOTH

# ---------------------------------------------------------------------------
# Sent Reminders JSON persistence (Atomic + Cached)
# ---------------------------------------------------------------------------


def load_sent_reminders(path: Path = SENT_REMINDERS_PATH) -> set[str]:
    """
    Load already-sent reminder keys from JSON.
    Merges with in-memory cache to guarantee zero loss across calls.
    Returns an empty set if the file doesn't exist and cache is empty.
    """
    key_path = str(path.resolve()) if path.is_absolute() else str(path)
    cached = _memory_sent_cache.get(key_path, set())

    if not path.exists():
        return set(cached)

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            file_keys = set(data)
            merged = file_keys | cached
            _memory_sent_cache[key_path] = merged
            return set(merged)
        logger.warning("%s has unexpected format; preserving cache.", path)
        return set(cached)
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Could not load %s: %s (preserving cache)", path, exc)
        return set(cached)


def save_sent_reminders(keys: set[str], path: Path = SENT_REMINDERS_PATH) -> None:
    """
    Save sent reminder keys to JSON atomically using temporary file + os.replace.
    Guarantees no partial writes or zero-byte corrupted files.
    """
    key_path = str(path.resolve()) if path.is_absolute() else str(path)
    _memory_sent_cache[key_path] = set(keys)

    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(f".tmp.{os.getpid()}")
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(sorted(keys), f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
    except OSError as exc:
        logger.error("Could not save %s atomically: %s", path, exc)


# ---------------------------------------------------------------------------
# Message formatting
# ---------------------------------------------------------------------------


def _format_ordinal_date(d: datetime.date) -> str:
    """Format date as e.g. 'July 14th', 'October 14th', 'November 1st'."""
    day = d.day
    if 11 <= (day % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{d.strftime('%B')} {day}{suffix}"


def _format_myanmar_date(d: datetime.date) -> str:
    """Format date for Myanmar text e.g. 'July 14', 'October 14'."""
    return f"{d.strftime('%B')} {d.day}"


def format_reminder_message(
    student_name: str,
    due_date: datetime.date,
    reminder_days: int = 5,
    start_date: Optional[datetime.date] = None,
    language: str = "both",
) -> str:
    """
    Format the reminder message from the template.
    language can be: 'burmese' (or 'my', 'myanmar'), 'english' (or 'en'), or 'both' (bilingual).
    """
    due_date_en = _format_ordinal_date(due_date)
    due_date_mm = _format_myanmar_date(due_date)

    start_date_en = _format_ordinal_date(start_date) if start_date else ""
    start_date_mm = _format_myanmar_date(start_date) if start_date else ""

    lang_normalized = (language or "both").strip().lower()
    if lang_normalized in ("burmese", "my", "myanmar", "mm"):
        template = REMINDER_TEMPLATE_MM
    elif lang_normalized in ("english", "en"):
        template = REMINDER_TEMPLATE_EN
    else:
        template = REMINDER_TEMPLATE_BOTH

    return template.format(
        student_name=student_name,
        start_date_en=start_date_en,
        due_date_en=due_date_en,
        start_date_mm=start_date_mm,
        due_date_mm=due_date_mm,
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
    where days_until_due >= 0. Used for the /status command display.
    """
    tz = ZoneInfo(timezone_str)
    today = datetime.datetime.now(tz).date()

    upcoming = []
    for rec in records:
        due_date: Optional[datetime.date] = rec.get("due_date")
        if not due_date:
            continue
        days_until = (due_date - today).days
        if days_until >= 0:
            upcoming.append({**rec, "days_until_due": days_until})

    upcoming.sort(key=lambda r: r["days_until_due"])
    return upcoming[:n]


# ---------------------------------------------------------------------------
# Core daily check
# ---------------------------------------------------------------------------


async def daily_check(
    bot: Bot,
    records: list[dict],
    owner_chat_id: Optional[int] = None,
    timezone_str: str = "Asia/Yangon",
    reminder_days: int = 5,
    sent_reminders_path: Path = SENT_REMINDERS_PATH,
    admin_chat_id: Optional[int] = None,
    default_language: str = "both",
) -> dict:
    """
    Main daily reminder check.

    For each record:
      - Skip if no chat_id (notify owner/admin).
      - Skip if due_date already passed.
      - If days_until_due == reminder_days and not already sent → send + save key.

    Returns:
      dict with detailed breakdown of the check results:
      {
        "today": date,
        "reminder_days": int,
        "total_records": int,
        "sent": list[dict],
        "already_sent": list[dict],
        "not_due": list[dict],
        "past_due": list[dict],
        "no_chat_id": list[dict],
        "failed": list[tuple[dict, str]],
      }
    """
    notify_chat_id = owner_chat_id if owner_chat_id is not None else admin_chat_id
    tz = ZoneInfo(timezone_str)
    today = datetime.datetime.now(tz).date()

    async with _daily_check_lock:
        sent_keys = load_sent_reminders(sent_reminders_path)
        newly_sent: list[str] = []
        sent_records: list[dict] = []
        already_sent_records: list[dict] = []
        not_due_records: list[dict] = []
        past_records: list[dict] = []
        skipped_no_chat_id: list[dict] = []
        failed: list[tuple[dict, str]] = []
        seen_in_run: set[str] = set()

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
                past_records.append(rec)
                continue

            days_until_due = (due_date - today).days

            # Send when within the reminder window (0 to reminder_days, or 0 to 5 in test mode)
            max_days = reminder_days if reminder_days > 0 else 5
            if not (0 <= days_until_due <= max_days):
                not_due_records.append({**rec, "days_until_due": days_until_due})
                continue

            # No chat_id → report to owner/admin, skip
            if chat_id is None:
                skipped_no_chat_id.append(rec)
                logger.warning(
                    "No chat_id for %s (due %s) — cannot send reminder.", username, due_date
                )
                continue

            # Canonical reminder key
            due_date_iso = (
                due_date.strftime("%Y-%m-%d")
                if isinstance(due_date, (datetime.date, datetime.datetime))
                else str(due_date).strip()
            )
            reminder_key = f"{chat_id}:{due_date_iso}"

            # Duplicate prevention (both from persistent history and earlier rows in this run)
            if reminder_key in sent_keys or reminder_key in seen_in_run:
                already_sent_records.append(rec)
                logger.info("Already sent/processed reminder for key %s — skipping.", reminder_key)
                continue

            seen_in_run.add(reminder_key)

            # Send the reminder
            start_date: Optional[datetime.date] = rec.get("start_date")
            rec_lang = rec.get("language")
            chosen_lang = rec_lang if rec_lang else default_language
            message = format_reminder_message(
                student_name, due_date, reminder_days, start_date=start_date, language=chosen_lang
            )
            try:
                await bot.send_message(chat_id=chat_id, text=message)
                sent_keys.add(reminder_key)
                newly_sent.append(reminder_key)
                sent_records.append(rec)
                # Immediately persist atomically to disk so no interruption can ever cause duplication
                save_sent_reminders(sent_keys, sent_reminders_path)
                logger.info("Reminder sent & persisted for %s (chat_id=%s, key=%s).", username, chat_id, reminder_key)
            except TelegramError as exc:
                failed.append((rec, str(exc)))
                logger.error(
                    "Failed to send reminder to %s (chat_id=%s): %s",
                    username,
                    chat_id,
                    exc,
                )

    # Report skipped (no chat_id) to owner
    if skipped_no_chat_id and notify_chat_id:
        lines = ["⚠️ Students missing Chat ID (reminders NOT sent):\n"]
        for rec in skipped_no_chat_id:
            lines.append(
                f"• {rec.get('student_name')} ({rec.get('telegram_username')}) "
                f"— due {rec.get('due_date')}"
            )
        await _notify_owner(bot, notify_chat_id, "\n".join(lines))

    # Report failures to owner
    if failed and notify_chat_id:
        lines = ["❌ Failed to send reminders:\n"]
        for rec, err in failed:
            lines.append(
                f"• {rec.get('student_name')} ({rec.get('telegram_username')}) "
                f"— {err}"
            )
        await _notify_owner(bot, notify_chat_id, "\n".join(lines))

    logger.info(
        "Daily check complete. Sent: %d, Skipped (no chat_id): %d, Failed: %d",
        len(newly_sent),
        len(skipped_no_chat_id),
        len(failed),
    )

    return {
        "today": today,
        "reminder_days": reminder_days,
        "total_records": len(records),
        "sent": sent_records,
        "already_sent": already_sent_records,
        "not_due": not_due_records,
        "past_due": past_records,
        "no_chat_id": skipped_no_chat_id,
        "failed": failed,
    }


# ---------------------------------------------------------------------------
# Owner notification helper
# ---------------------------------------------------------------------------


async def _notify_owner(bot: Bot, chat_id: int, text: str) -> None:
    """Send a message to the user/owner, logging any failure silently."""
    try:
        await bot.send_message(chat_id=chat_id, text=text)
    except TelegramError as exc:
        logger.error("Could not notify user (chat_id=%s): %s", chat_id, exc)


# Backwards compatibility alias
_notify_admin = _notify_owner
