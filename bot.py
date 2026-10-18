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

import datetime
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
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.error import BadRequest, TelegramError
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

import excel_handler
import i18n
from i18n import t, get_user_language, set_user_language, SUPPORTED_LANGUAGES, LANG_EN, LANG_MY
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
CHAT_ID_REGISTRY_PATH = DATA_DIR / "chat_id_registry.json"
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

# Chat ID registry: maps lowercase @username -> chat_id (auto-captured when users /start)
_chat_id_registry: dict[str, int] = {}

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
# Chat ID registry persistence
# ---------------------------------------------------------------------------


def _load_chat_id_registry() -> dict[str, int]:
    """Load the @username -> chat_id registry from JSON."""
    if not CHAT_ID_REGISTRY_PATH.exists():
        return {}
    try:
        with open(CHAT_ID_REGISTRY_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {k.lower(): int(v) for k, v in data.items()}
    except (json.JSONDecodeError, OSError, ValueError) as exc:
        logger.warning("Could not load chat_id_registry.json: %s", exc)
    return {}


def _save_chat_id_registry() -> None:
    """Save the @username -> chat_id registry to JSON."""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(CHAT_ID_REGISTRY_PATH, "w", encoding="utf-8") as f:
            json.dump(_chat_id_registry, f, indent=2)
    except OSError as exc:
        logger.error("Could not save chat_id_registry.json: %s", exc)


def _register_chat_id(username: str | None, chat_id: int) -> None:
    """Register a @username -> chat_id mapping (if username is available)."""
    if not username:
        return
    key = f"@{username.lower().lstrip('@')}"
    if _chat_id_registry.get(key) != chat_id:
        _chat_id_registry[key] = chat_id
        _save_chat_id_registry()
        logger.info("Registered chat_id %s for %s", chat_id, key)


def _resolve_chat_ids(records: list[dict]) -> int:
    """
    Fill in missing telegram_chat_id from the registry.
    Returns the number of records resolved.
    """
    resolved = 0
    for rec in records:
        if rec.get("telegram_chat_id"):
            continue  # Already has a chat_id
        username = rec.get("telegram_username")
        if not username:
            continue
        key = username.lower()
        chat_id = _chat_id_registry.get(key)
        if chat_id:
            rec["telegram_chat_id"] = chat_id
            resolved += 1
            logger.info("Resolved chat_id for %s -> %s", username, chat_id)
    return resolved


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


# ---------------------------------------------------------------------------
# Menu UI & Helper Functions
# ---------------------------------------------------------------------------

PAGE_SIZE = 6


def _build_main_menu_keyboard(is_admin: bool = False, lang: str = LANG_MY) -> InlineKeyboardMarkup:
    """Build the primary navigation keyboard for the bot in the chosen language."""
    buttons = [
        [
            InlineKeyboardButton(t("btn_roster_status", lang), callback_data="menu:status"),
            InlineKeyboardButton(t("btn_scheduler", lang), callback_data="menu:scheduler"),
        ],
        [
            InlineKeyboardButton(t("btn_check_send", lang), callback_data="menu:check"),
            InlineKeyboardButton(t("btn_delete_reminders", lang), callback_data="menu:delete_list:0"),
        ],
        [
            InlineKeyboardButton(t("btn_sent_history", lang), callback_data="menu:history"),
            InlineKeyboardButton(t("btn_language", lang), callback_data="menu:language"),
        ],
        [
            InlineKeyboardButton(t("btn_refresh_menu", lang), callback_data="menu:home"),
        ],
    ]
    if is_admin:
        buttons.append(
            [InlineKeyboardButton(t("btn_admin_users", lang), callback_data="menu:users")]
        )
    return InlineKeyboardMarkup(buttons)


def _render_language_menu(user_id: int) -> tuple[str, InlineKeyboardMarkup]:
    """Render the language selection view with buttons for Burmese and English."""
    cur_lang = get_user_language(user_id)
    lang_name = SUPPORTED_LANGUAGES.get(cur_lang, cur_lang)
    text = t("lang_menu_text", cur_lang, cur_lang=lang_name)
    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    ("✅ " if cur_lang == LANG_MY else "") + "🇲🇲 မြန်မာ (Burmese)",
                    callback_data="set_lang:my",
                )
            ],
            [
                InlineKeyboardButton(
                    ("✅ " if cur_lang == LANG_EN else "") + "🇬🇧 English",
                    callback_data="set_lang:en",
                )
            ],
            [InlineKeyboardButton(t("btn_back_home", cur_lang), callback_data="menu:home")],
        ]
    )
    return text, keyboard


def _get_home_text(user_id: int, user_name: str, lang: str | None = None) -> str:
    """Return the dashboard summary text for the main menu."""
    if lang is None:
        lang = get_user_language(user_id)
    record_count = len(_user_records.get(user_id, []))
    tz = ZoneInfo(TIMEZONE)
    now = datetime.datetime.now(tz)
    now_str = now.strftime("%Y-%m-%d %H:%M")

    job = _scheduler.get_job("daily_reminder_check") if _scheduler else None
    countdown_str = "None"
    if job and job.next_run_time:
        diff = job.next_run_time.astimezone(tz) - now
        tot = max(0, int(diff.total_seconds()))
        m, s = divmod(tot, 60)
        h, m = divmod(m, 60)
        countdown_parts = []
        if h > 0:
            countdown_parts.append(f"{h}h")
        if m > 0:
            countdown_parts.append(f"{m}m")
        countdown_parts.append(f"{s}s")
        countdown_str = " ".join(countdown_parts)

    lang_label = SUPPORTED_LANGUAGES.get(lang, "🇲🇲 မြန်မာ")

    return t(
        "home_dashboard",
        lang,
        name=user_name,
        record_count=record_count,
        now_str=now_str,
        timezone=TIMEZONE,
        countdown_str=countdown_str,
        reminder_days=REMINDER_DAYS_BEFORE,
        lang_label=lang_label,
    )


def _render_status_text(user_id: int, lang: str | None = None) -> str:
    """Format roster status summary."""
    if lang is None:
        lang = get_user_language(user_id)
    records = _user_records.get(user_id, [])
    total = len(records)
    _resolve_chat_ids(records)
    with_chat_id = sum(1 for r in records if r.get("telegram_chat_id"))
    without_chat_id = total - with_chat_id

    now = datetime.datetime.now(ZoneInfo(TIMEZONE))
    now_str = now.strftime("%Y-%m-%d %H:%M")

    lines = [
        t("status_title", lang),
        t("status_total", lang, total=total),
        t("status_linked", lang, with_chat_id=with_chat_id),
        t("status_missing", lang, without_chat_id=without_chat_id),
        t("status_time", lang, now_str=now_str, timezone=TIMEZONE),
        t("status_offset", lang, reminder_days=REMINDER_DAYS_BEFORE),
    ]

    if total == 0:
        lines.append(t("status_empty", lang))
    else:
        upcoming = reminder_scheduler.get_upcoming_reminders(
            records, TIMEZONE, REMINDER_DAYS_BEFORE, n=5
        )
        if upcoming:
            lines.append(t("status_upcoming_header", lang))
            for rec in upcoming:
                days = rec["days_until_due"]
                if lang == LANG_MY:
                    due_label = "ယနေ့!" if days == 0 else f"{days} ရက်အလို"
                else:
                    due_label = "today!" if days == 0 else f"in {days} day(s)"
                uname = rec.get("telegram_username", "")
                sname = rec.get("student_name", uname)
                lines.append(
                    f"• {sname} (`{uname}`) — due `{rec['due_date']}` ({due_label})"
                )
        else:
            lines.append(t("status_no_upcoming", lang))

    return "\n".join(lines)


def _render_scheduler_text(user_id: int, lang: str | None = None) -> str:
    """Format scheduler engine state and active schedule data."""
    if lang is None:
        lang = get_user_language(user_id)
    tz = ZoneInfo(TIMEZONE)
    now = datetime.datetime.now(tz)
    today = now.date()
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    job = _scheduler.get_job("daily_reminder_check") if _scheduler else None
    is_running = _scheduler is not None and _scheduler.running

    status_str = "🟢 Active" if is_running else "🔴 Stopped"
    lines = [
        t("scheduler_title", lang),
        t("scheduler_status", lang, status=status_str),
        t("status_time", lang, now_str=now_str, timezone=TIMEZONE),
        t("status_offset", lang, reminder_days=REMINDER_DAYS_BEFORE),
    ]

    if job and job.next_run_time:
        next_run = job.next_run_time.astimezone(tz)
        diff = next_run - now
        total_seconds = max(0, int(diff.total_seconds()))
        minutes, seconds = divmod(total_seconds, 60)
        hours, minutes = divmod(minutes, 60)

        countdown = []
        if hours > 0:
            countdown.append(f"{hours}h")
        if minutes > 0:
            countdown.append(f"{minutes}m")
        countdown.append(f"{seconds}s")
        countdown_str = " ".join(countdown)

        trigger_desc = str(job.trigger)
        if "interval" in trigger_desc.lower():
            mode_desc = "`Interval (every 3 minutes)` *(Testing Mode)*"
        else:
            mode_desc = f"`Cron (Daily at {REMINDER_HOUR:02d}:{REMINDER_MINUTE:02d})` *(Production)*"

        lines += [
            t("scheduler_mode", lang, mode=mode_desc),
            t("scheduler_next_check", lang, next_run=next_run.strftime('%Y-%m-%d %H:%M:%S')),
            t("scheduler_countdown", lang, countdown=countdown_str),
        ]
    else:
        lines.append(t("scheduler_no_check", lang))

    lines.append(t("scheduler_active_data", lang))
    records = _user_records.get(user_id, [])
    _resolve_chat_ids(records)

    if not records:
        lines.append(t("status_empty", lang))
    else:
        sent_path = _user_sent_reminders_path(user_id)
        sent_keys = reminder_scheduler.load_sent_reminders(sent_path)

        due_on_target: list[dict] = []
        future_scheduled: list[dict] = []
        already_sent_count = 0
        missing_chat_id_count = 0

        for r in records:
            due = r.get("due_date")
            cid = r.get("telegram_chat_id")
            if not cid:
                missing_chat_id_count += 1
            if due:
                key = f"{cid}:{due}"
                if key in sent_keys:
                    already_sent_count += 1
                days = (due - today).days
                if days == REMINDER_DAYS_BEFORE:
                    due_on_target.append(
                        {**r, "days": days, "is_sent": key in sent_keys}
                    )
                elif days > REMINDER_DAYS_BEFORE:
                    future_scheduled.append(
                        {**r, "days": days, "is_sent": key in sent_keys}
                    )

        lines.append(t("scheduler_roster_summary", lang, total=len(records)))
        lines.append(t("scheduler_sent_summary", lang, sent=already_sent_count, missing=missing_chat_id_count))
        lines.append(t("scheduler_queue_header", lang, reminder_days=REMINDER_DAYS_BEFORE))
        if not due_on_target:
            lines.append(t("scheduler_queue_empty", lang))
        else:
            for r in due_on_target:
                uname = r.get("telegram_username", "")
                sname = r.get("student_name", uname)
                cid = r.get("telegram_chat_id")
                status = "✅ Sent" if r["is_sent"] else ("⏳ Pending" if cid else "⚠️ Missing Chat ID")
                lines.append(f"  • `{uname}` ({sname}) — `{r.get('due_date')}` [{status}]")

        future_scheduled.sort(key=lambda x: x["days"])
        lines.append("\n📅 *Upcoming Scheduled Reminders:*")
        if not future_scheduled:
            lines.append("  _No future reminders beyond the current target offset._")
        else:
            for r in future_scheduled[:10]:
                uname = r.get("telegram_username", "")
                sname = r.get("student_name", uname)
                cid = r.get("telegram_chat_id")
                cid_status = "✅ Linked" if cid else "⚠️ No Chat ID"
                lines.append(
                    f"  • `{uname}` ({sname}) — `{r.get('due_date')}` (in {r['days']}d) [{cid_status}]"
                )
            if len(future_scheduled) > 10:
                lines.append(
                    f"  _… and {len(future_scheduled) - 10} more upcoming._"
                )

    return "\n".join(lines)


def _render_history_text(user_id: int, lang: str | None = None) -> str:
    """Format sent reminder history."""
    if lang is None:
        lang = get_user_language(user_id)
    path = _user_sent_reminders_path(user_id)
    sent_keys = reminder_scheduler.load_sent_reminders(path)

    lines = [
        t("history_title", lang),
        t("history_total", lang, count=len(sent_keys)),
    ]

    if not sent_keys:
        lines.append(t("history_empty", lang))
    else:
        lines.append("\n*Recent entries (`chat_id:due_date`):*")
        for key in sorted(sent_keys, reverse=True)[:15]:
            lines.append(f"• `{key}`")
        if len(sent_keys) > 15:
            lines.append(f"… and {len(sent_keys) - 15} more.")

    return "\n".join(lines)


def _render_delete_list(
    user_id: int, page: int = 0, lang: str | None = None
) -> tuple[str, InlineKeyboardMarkup]:
    """Render list of reminder schedules with per-item delete buttons."""
    if lang is None:
        lang = get_user_language(user_id)
    records = _user_records.get(user_id, [])
    if not records:
        text = t("delete_empty", lang)
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")]]
        )
        return text, keyboard

    total_records = len(records)
    total_pages = max(1, (total_records + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(0, min(page, total_pages - 1))

    start_idx = page * PAGE_SIZE
    end_idx = min(start_idx + PAGE_SIZE, total_records)
    page_records = records[start_idx:end_idx]

    lines = [
        t("delete_title", lang),
        t("delete_showing", lang, start=start_idx+1, end=end_idx, total=total_records),
    ]

    buttons = []
    for i, rec in enumerate(page_records, start=start_idx):
        uname = rec.get("telegram_username", f"Student {i+1}")
        sname = rec.get("student_name", "")
        due = rec.get("due_date")
        lines.append(f"{i+1}. `{uname}` ({sname}) — Due: `{due}`")
        label = f"❌ {i+1}. {uname}"
        buttons.append(
            [InlineKeyboardButton(label, callback_data=f"del:ask:{i}:{page}")]
        )

    # Pagination navigation
    nav_buttons = []
    if page > 0:
        nav_buttons.append(
            InlineKeyboardButton(t("btn_prev", lang), callback_data=f"menu:delete_list:{page-1}")
        )
    if total_pages > 1:
        nav_buttons.append(
            InlineKeyboardButton(
                f"📄 {page+1}/{total_pages}", callback_data=f"menu:delete_list:{page}"
            )
        )
    if page < total_pages - 1:
        nav_buttons.append(
            InlineKeyboardButton(t("btn_next", lang), callback_data=f"menu:delete_list:{page+1}")
        )
    if nav_buttons:
        buttons.append(nav_buttons)

    buttons.append(
        [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")]
    )
    return "\n".join(lines), InlineKeyboardMarkup(buttons)


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/start — Welcome authorized users with Menu UI, register new students/teachers."""
    user = update.effective_user
    chat_id = update.effective_chat.id
    user_id = user.id
    name = user.first_name if user else "there"

    # Always register username -> chat_id for reminder delivery
    _register_chat_id(user.username, chat_id)
    lang = get_user_language(user_id)

    if user_id in _authorized_users:
        # Already authorized (admin or previously approved) -> show full Menu UI
        text = _get_home_text(user_id, name, lang)
        keyboard = _build_main_menu_keyboard(_is_admin(update), lang)
        await update.message.reply_text(
            text, reply_markup=keyboard, parse_mode="Markdown"
        )
    else:
        # New user — registered for reminders + offer teacher authorization
        _pending_auth.add(user_id)
        text = t("student_registered", lang, name=name)
        toggle_lang = LANG_EN if lang == LANG_MY else LANG_MY
        toggle_label = "🇬🇧 English" if lang == LANG_MY else "🇲🇲 မြန်မာ"
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton(f"🌐 Switch to {toggle_label}", callback_data=f"switch_start_lang:{toggle_lang}")]]
        )
        await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/help — Same as /start /menu."""
    await cmd_start(update, context)


async def cmd_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/menu — Display the interactive Menu UI."""
    await cmd_start(update, context)


async def cmd_language(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/language [burmese|english] — View or change bot language."""
    user = update.effective_user
    if not user:
        return
    user_id = user.id
    args = context.args or []
    if args:
        chosen = args[0].lower()
        if chosen in ("burmese", "my", "myanmar", "mm", "မြန်မာ"):
            new_code = set_user_language(user_id, LANG_MY)
        elif chosen in ("english", "en"):
            new_code = set_user_language(user_id, LANG_EN)
        else:
            await update.message.reply_text("❌ Please specify `burmese` (or `my`) or `english` (or `en`).", parse_mode="Markdown")
            return
        await update.message.reply_text(t("lang_changed", new_code), parse_mode="Markdown")
        return

    text, keyboard = _render_language_menu(user_id)
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/status — Show the calling user's own records + next 5 reminders."""
    if not _is_authorized(update):
        user_id = update.effective_user.id if update.effective_user else 0
        lang = get_user_language(user_id)
        await update.message.reply_text(t("not_authorized", lang))
        return

    user_id = update.effective_user.id
    lang = get_user_language(user_id)
    text = _render_status_text(user_id, lang)
    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(t("btn_send_now_short", lang), callback_data="menu:check"),
                InlineKeyboardButton(t("btn_delete_short", lang), callback_data="menu:delete_list:0"),
            ],
            [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
        ]
    )
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_check(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/check [all] — Manually trigger a reminder check right now and report results."""
    if not _is_authorized(update):
        user_id = update.effective_user.id if update.effective_user else 0
        lang = get_user_language(user_id)
        await update.message.reply_text(t("not_authorized", lang))
        return

    user_id = update.effective_user.id
    is_admin = _is_admin(update)
    lang = get_user_language(user_id)
    args = context.args or []
    check_all = is_admin and len(args) > 0 and args[0].lower() == "all"

    targets = list(_user_records.keys()) if check_all else [user_id]
    msg = await update.message.reply_text(t("check_running", lang))

    today = datetime.datetime.now(ZoneInfo(TIMEZONE)).date()
    lines = [
        t("check_report_title", lang),
        f"📅 Date: `{today}` (`{TIMEZONE}`)",
        f"🎯 Window: `{REMINDER_DAYS_BEFORE}` day(s) before deadline",
        "",
    ]

    total_sent = 0
    total_already_sent = 0
    total_not_due = 0
    total_missing_chat_id = 0
    total_failed = 0
    total_records_checked = 0

    for uid in targets:
        records = _user_records.get(uid, [])
        if not records:
            if not check_all:
                lines.append("ℹ️ _No student records loaded. Upload an Excel file first._")
            continue

        total_records_checked += len(records)
        _resolve_chat_ids(records)
        user_lang = get_user_language(uid)
        try:
            res = await reminder_scheduler.daily_check(
                bot=context.bot,
                records=records,
                owner_chat_id=uid,
                timezone_str=TIMEZONE,
                reminder_days=REMINDER_DAYS_BEFORE,
                sent_reminders_path=_user_sent_reminders_path(uid),
                default_language=user_lang,
            )
        except Exception as exc:
            logger.error("Manual check failed for user %s: %s", uid, exc)
            lines.append(f"❌ Check failed for user `{uid}`: {exc}")
            continue

        prefix = f"[User {uid}] " if check_all else ""

        if res["sent"]:
            lines.append(f"✅ *{prefix}Sent ({len(res['sent'])}):*")
            for r in res["sent"][:10]:
                lines.append(f"  • {r['student_name']} (`{r['telegram_username']}`) — due `{r['due_date']}`")
            total_sent += len(res["sent"])

        if res["already_sent"]:
            lines.append(f"ℹ️ *{prefix}Already Sent ({len(res['already_sent'])}):*")
            for r in res["already_sent"][:10]:
                lines.append(f"  • {r['student_name']} (`{r['telegram_username']}`) — due `{r['due_date']}`")
            total_already_sent += len(res["already_sent"])

        if res["not_due"]:
            lines.append(f"⏳ *{prefix}Not Due Yet ({len(res['not_due'])}):*")
            for r in res["not_due"][:10]:
                days = r.get("days_until_due", "?")
                lines.append(
                    f"  • {r['student_name']} (`{r['telegram_username']}`) — due `{r['due_date']}` "
                    f"(in {days}d; target={REMINDER_DAYS_BEFORE}d)"
                )
            total_not_due += len(res["not_due"])

        if res["no_chat_id"]:
            lines.append(f"⚠️ *{prefix}Missing Chat ID ({len(res['no_chat_id'])}):*")
            for r in res["no_chat_id"][:10]:
                lines.append(f"  • {r['student_name']} (`{r['telegram_username']}`)")
            total_missing_chat_id += len(res["no_chat_id"])

        if res["failed"]:
            lines.append(f"❌ *{prefix}Send Failed ({len(res['failed'])}):*")
            for r, err in res["failed"][:10]:
                lines.append(f"  • `{r.get('telegram_username')}`: {err}")
            total_failed += len(res["failed"])

    if total_records_checked > 0 and not total_sent and not total_already_sent and not total_not_due:
        lines.append(f"ℹ️ Checked {total_records_checked} record(s). None matched.")

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(t("btn_scheduler", lang), callback_data="menu:scheduler"),
                InlineKeyboardButton(t("btn_sent_history", lang), callback_data="menu:history"),
            ],
            [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
        ]
    )
    await msg.edit_text("\n".join(lines), reply_markup=keyboard, parse_mode="Markdown")


async def cmd_scheduler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/scheduler — Check scheduler state, countdown, and active schedule data."""
    if not _is_authorized(update):
        user_id = update.effective_user.id if update.effective_user else 0
        lang = get_user_language(user_id)
        await update.message.reply_text(t("not_authorized", lang))
        return

    user_id = update.effective_user.id
    lang = get_user_language(user_id)
    text = _render_scheduler_text(user_id, lang)
    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(t("btn_check_send", lang), callback_data="menu:check"),
                InlineKeyboardButton(t("btn_delete_short", lang), callback_data="menu:delete_list:0"),
            ],
            [
                InlineKeyboardButton(t("btn_refresh_queue", lang), callback_data="menu:scheduler"),
                InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home"),
            ],
        ]
    )
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_delete(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/delete — Interactive UI to delete individual reminder schedules."""
    if not _is_authorized(update):
        user_id = update.effective_user.id if update.effective_user else 0
        lang = get_user_language(user_id)
        await update.message.reply_text(t("not_authorized", lang))
        return

    user_id = update.effective_user.id
    lang = get_user_language(user_id)
    text, keyboard = _render_delete_list(user_id, page=0, lang=lang)
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_history(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/history — View sent reminders history for this user."""
    if not _is_authorized(update):
        user_id = update.effective_user.id if update.effective_user else 0
        lang = get_user_language(user_id)
        await update.message.reply_text(t("not_authorized", lang))
        return

    user_id = update.effective_user.id
    lang = get_user_language(user_id)
    text = _render_history_text(user_id, lang)
    keyboard = InlineKeyboardMarkup(
        [
            [InlineKeyboardButton(t("btn_clear_history", lang), callback_data="menu:clearsent")],
            [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
        ]
    )
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_clearsent(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/clearsent — Clear sent reminders history for this user (for testing)."""
    if not _is_authorized(update):
        user_id = update.effective_user.id if update.effective_user else 0
        lang = get_user_language(user_id)
        await update.message.reply_text(t("not_authorized", lang))
        return

    user_id = update.effective_user.id
    lang = get_user_language(user_id)
    path = _user_sent_reminders_path(user_id)
    sent_keys = reminder_scheduler.load_sent_reminders(path)
    count = len(sent_keys)

    reminder_scheduler.save_sent_reminders(set(), path)

    text = t("history_cleared", lang, count=count)
    keyboard = InlineKeyboardMarkup(
        [
            [InlineKeyboardButton(t("btn_send_now_short", lang), callback_data="menu:check")],
            [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
        ]
    )
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="Markdown")


async def cmd_users(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/users — Admin-only: list all authorized teachers and student registry stats."""
    if not _is_admin(update):
        await update.message.reply_text("❌ This command is only available to the Admin.")
        return

    lines = [
        "👥 *Authorized Users & Registry*",
        f"• *Admin Chat ID:* `{ADMIN_TELEGRAM_ID}`",
        f"• *Authorized Teachers ({len(_authorized_users)}):*",
    ]

    for uid in sorted(_authorized_users):
        rec_count = len(_user_records.get(uid, []))
        role = "👑 Admin" if uid == ADMIN_TELEGRAM_ID else "🎓 Teacher"
        lines.append(f"  - `{uid}` ({role}): {rec_count} student(s) loaded")

    lines += [
        "",
        f"• *Auto-Captured Student Chat IDs:* {len(_chat_id_registry)} registered",
    ]

    if _chat_id_registry:
        reg_usernames = list(_chat_id_registry.keys())[:10]
        lines.append("  Registered usernames: " + ", ".join(reg_usernames))
        if len(_chat_id_registry) > 10:
            lines.append(f"  … and {len(_chat_id_registry) - 10} more.")

    keyboard = InlineKeyboardMarkup(
        [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
    )
    await update.message.reply_text("\n".join(lines), reply_markup=keyboard, parse_mode="Markdown")


async def handle_callback_query(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Handle all interactive inline keyboard clicks for the menu UI."""
    query = update.callback_query
    if not query:
        return
    await query.answer()

    user_id = query.from_user.id
    data = query.data or ""

    # Allow unauthenticated users to toggle initial welcome language
    if data.startswith("switch_start_lang:"):
        target_lang = data.split(":")[1]
        set_user_language(user_id, target_lang)
        first_name = query.from_user.first_name if query.from_user else "there"
        text = t("student_registered", target_lang, name=first_name)
        toggle_lang = LANG_EN if target_lang == LANG_MY else LANG_MY
        toggle_label = "🇬🇧 English" if target_lang == LANG_MY else "🇲🇲 မြန်မာ"
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton(f"🌐 Switch to {toggle_label}", callback_data=f"switch_start_lang:{toggle_lang}")]]
        )
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass
        return

    if user_id not in _authorized_users:
        lang = get_user_language(user_id)
        try:
            await query.edit_text(t("not_authorized", lang))
        except BadRequest:
            pass
        return

    is_admin = user_id == ADMIN_TELEGRAM_ID
    lang = get_user_language(user_id)

    if data == "menu:home":
        first_name = query.from_user.first_name if query.from_user else "there"
        text = _get_home_text(user_id, first_name, lang)
        keyboard = _build_main_menu_keyboard(is_admin, lang)
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:language":
        text, keyboard = _render_language_menu(user_id)
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data.startswith("set_lang:"):
        chosen = data.split(":")[1]
        set_user_language(user_id, chosen)
        await query.answer(t("lang_changed", chosen), show_alert=False)
        first_name = query.from_user.first_name if query.from_user else "there"
        text = _get_home_text(user_id, first_name, chosen)
        keyboard = _build_main_menu_keyboard(is_admin, chosen)
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:status":
        text = _render_status_text(user_id, lang)
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(t("btn_send_now_short", lang), callback_data="menu:check"),
                    InlineKeyboardButton(t("btn_delete_short", lang), callback_data="menu:delete_list:0"),
                ],
                [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
            ]
        )
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:scheduler":
        text = _render_scheduler_text(user_id, lang)
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(t("btn_check_send", lang), callback_data="menu:check"),
                    InlineKeyboardButton(t("btn_delete_short", lang), callback_data="menu:delete_list:0"),
                ],
                [
                    InlineKeyboardButton(t("btn_refresh_queue", lang), callback_data="menu:scheduler"),
                    InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home"),
                ],
            ]
        )
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:check":
        records = _user_records.get(user_id, [])
        _resolve_chat_ids(records)
        today = datetime.datetime.now(ZoneInfo(TIMEZONE)).date()

        lines = [
            t("check_report_title", lang),
            f"📅 Date: `{today}` (`{TIMEZONE}`)",
            f"🎯 Window: `{REMINDER_DAYS_BEFORE}` day(s) before deadline",
            "",
        ]

        if not records:
            lines.append("ℹ️ _No student records loaded. Upload an Excel file first._")
        else:
            try:
                res = await reminder_scheduler.daily_check(
                    bot=context.bot,
                    records=records,
                    owner_chat_id=user_id,
                    timezone_str=TIMEZONE,
                    reminder_days=REMINDER_DAYS_BEFORE,
                    sent_reminders_path=_user_sent_reminders_path(user_id),
                    default_language=lang,
                )
                if res["sent"]:
                    lines.append(f"✅ *Sent ({len(res['sent'])}):*")
                    for r in res["sent"][:10]:
                        lines.append(
                            f"  • {r['student_name']} (`{r['telegram_username']}`) — due `{r['due_date']}`"
                        )
                if res["already_sent"]:
                    lines.append(f"ℹ️ *Already Sent ({len(res['already_sent'])}):*")
                    for r in res["already_sent"][:10]:
                        lines.append(
                            f"  • {r['student_name']} (`{r['telegram_username']}`) — due `{r['due_date']}`"
                        )
                if res["not_due"]:
                    lines.append(f"⏳ *Not Due Yet ({len(res['not_due'])}):*")
                    for r in res["not_due"][:10]:
                        days = r.get("days_until_due", "?")
                        lines.append(
                            f"  • {r['student_name']} (`{r['telegram_username']}`) — due `{r['due_date']}` (in {days}d)"
                        )
                if res["no_chat_id"]:
                    lines.append(f"⚠️ *Missing Chat ID ({len(res['no_chat_id'])}):*")
                    for r in res["no_chat_id"][:10]:
                        lines.append(f"  • {r['student_name']} (`{r['telegram_username']}`)")
                if res["failed"]:
                    lines.append(f"❌ *Send Failed ({len(res['failed'])}):*")
                    for r, err in res["failed"][:10]:
                        lines.append(f"  • `{r.get('telegram_username')}`: {err}")
            except Exception as exc:
                lines.append(f"❌ Check failed: {exc}")

        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(t("btn_scheduler", lang), callback_data="menu:scheduler"),
                    InlineKeyboardButton(t("btn_sent_history", lang), callback_data="menu:history"),
                ],
                [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
            ]
        )
        try:
            await query.edit_text("\n".join(lines), reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:history":
        text = _render_history_text(user_id, lang)
        keyboard = InlineKeyboardMarkup(
            [
                [InlineKeyboardButton(t("btn_clear_history", lang), callback_data="menu:clearsent")],
                [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
            ]
        )
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:clearsent":
        path = _user_sent_reminders_path(user_id)
        sent_keys = reminder_scheduler.load_sent_reminders(path)
        count = len(sent_keys)
        reminder_scheduler.save_sent_reminders(set(), path)

        text = t("history_cleared", lang, count=count)
        keyboard = InlineKeyboardMarkup(
            [
                [InlineKeyboardButton(t("btn_send_now_short", lang), callback_data="menu:check")],
                [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
            ]
        )
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data == "menu:users":
        if not is_admin:
            await query.answer("❌ Admin only", show_alert=True)
            return
        lines = [
            "👥 *Authorized Users & Registry*",
            f"• *Admin Chat ID:* `{ADMIN_TELEGRAM_ID}`",
            f"• *Authorized Teachers ({len(_authorized_users)}):*",
        ]
        for uid in sorted(_authorized_users):
            rec_count = len(_user_records.get(uid, []))
            role = "👑 Admin" if uid == ADMIN_TELEGRAM_ID else "🎓 Teacher"
            lines.append(f"  - `{uid}` ({role}): {rec_count} student(s) loaded")
        lines += [
            "",
            f"• *Auto-Captured Student Chat IDs:* {len(_chat_id_registry)} registered",
        ]
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")]]
        )
        try:
            await query.edit_text("\n".join(lines), reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data.startswith("menu:delete_list:"):
        page = int(data.split(":")[2])
        text, keyboard = _render_delete_list(user_id, page, lang)
        try:
            await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
        except BadRequest:
            pass

    elif data.startswith("del:ask:"):
        parts = data.split(":")
        idx = int(parts[2])
        page = int(parts[3])
        records = _user_records.get(user_id, [])
        if 0 <= idx < len(records):
            rec = records[idx]
            uname = rec.get("telegram_username", "")
            sname = rec.get("student_name", uname)
            due = rec.get("due_date")
            text = t(
                "delete_confirm_prompt",
                lang,
                student_name=sname,
                username=uname,
                due_date=due,
            )
            keyboard = InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            t("btn_confirm_delete", lang),
                            callback_data=f"del:confirm:{idx}:{page}",
                        ),
                        InlineKeyboardButton(
                            t("btn_cancel", lang),
                            callback_data=f"menu:delete_list:{page}",
                        ),
                    ],
                    [InlineKeyboardButton(t("btn_back_home", lang), callback_data="menu:home")],
                ]
            )
            try:
                await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
            except BadRequest:
                pass

    elif data.startswith("del:confirm:"):
        parts = data.split(":")
        idx = int(parts[2])
        page = int(parts[3])
        records = _user_records.get(user_id, [])
        if 0 <= idx < len(records):
            removed = records.pop(idx)
            # Save updated records to user's Excel file
            excel_path = _user_excel_path(user_id)
            excel_handler.save_excel(records, str(excel_path))
            uname = removed.get("telegram_username", "")
            due = removed.get("due_date")
            await query.answer(f"✅ Deleted reminder for {uname}!", show_alert=False)

            text_list, keyboard = _render_delete_list(user_id, page, lang)
            success_msg = t("delete_success", lang, username=f"`{uname}`", due_date=f"`{due}`")
            text = f"✅ {success_msg}\n\n" + text_list
            try:
                await query.edit_text(text, reply_markup=keyboard, parse_mode="Markdown")
            except BadRequest:
                pass



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

    # Auto-resolve missing chat_ids from the registry
    resolved = _resolve_chat_ids(valid_records)
    unresolved = sum(1 for r in valid_records if not r.get("telegram_chat_id"))

    total = len(valid_records) + invalid_count
    lang = get_user_language(user_id)
    reply_lines = [
        t("excel_saved_title", lang),
        t("excel_total_records", lang, total=total),
        t("excel_valid_records", lang, valid=len(valid_records)),
    ]
    if invalid_count > 0:
        reply_lines.append(f"• ⚠️ Invalid records: {invalid_count}")

    if resolved:
        reply_lines.append(t("excel_auto_matched", lang, matched=resolved))
    if unresolved:
        reply_lines.append(t("excel_unregistered", lang, unregistered=unresolved))
        # List unresolved usernames
        unresolved_names = [
            r.get("telegram_username", "?")
            for r in valid_records
            if not r.get("telegram_chat_id")
        ]
        note = "These students need to tap /start on this bot: " if lang == LANG_EN else "အောက်ပါကျောင်းသားများ /start နှိပ်ရန်လိုအပ်ပါသည်: "
        reply_lines.append(note + ", ".join(unresolved_names[:10]))
        if len(unresolved_names) > 10:
            reply_lines.append(f"… and {len(unresolved_names) - 10} more.")

    scheduled_note = "Automatic payment reminders are now scheduled." if lang == LANG_EN else "အလိုအလျောက် သတိပေးချက်များ သတ်မှတ်ပြီးပါပြီ။"
    reply_lines += ["", scheduled_note]

    # Immediately check & send any reminders that are due right now
    try:
        check_res = await reminder_scheduler.daily_check(
            bot=context.bot,
            records=valid_records,
            owner_chat_id=user_id,
            timezone_str=TIMEZONE,
            reminder_days=REMINDER_DAYS_BEFORE,
            sent_reminders_path=_user_sent_reminders_path(user_id),
            default_language=lang,
        )
        if check_res and check_res.get("sent"):
            sent_students = [
                f"{r.get('student_name', r.get('telegram_username'))} ({r.get('telegram_username')})"
                for r in check_res["sent"]
            ]
            reply_lines.append(
                f"\n📨 *Immediately sent {len(check_res['sent'])} reminder(s):*\n• "
                + "\n• ".join(sent_students[:5])
            )
    except Exception as exc:
        logger.error("Immediate check after upload failed for user %s: %s", user_id, exc)

    if errors:
        reply_lines += ["", "⚠️ *Issues found (rows skipped):*", _format_errors(errors)]

    reply_lines += ["\n💡 Tip: Use /menu or /status to view your dashboard."]
    await update.message.reply_text("\n".join(reply_lines), parse_mode="Markdown")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle text messages — authorization flow for pending users."""
    user_id = update.effective_user.id

    if user_id not in _pending_auth:
        return  # No pending auth — ignore

    lang = get_user_language(user_id)
    text = (update.message.text or "").strip()

    # Validate the entered Chat ID against the admin ID
    try:
        entered_id = int(text)
    except ValueError:
        await update.message.reply_text(
            "❌ That doesn't look like a valid Chat ID. Please enter a number."
            if lang == LANG_EN
            else "❌ မှန်ကန်သော Chat ID မဟုတ်ပါ။ ဂဏန်းသီးသန့် ရိုက်ထည့်ပေးပါ။"
        )
        return

    if entered_id != ADMIN_TELEGRAM_ID:
        _pending_auth.discard(user_id)
        logger.warning("User %s entered wrong admin Chat ID: %s", user_id, entered_id)
        await update.message.reply_text(t("auth_failed", lang), parse_mode="Markdown")
        return

    # Correct — authorize the user
    _pending_auth.discard(user_id)
    _authorized_users.add(user_id)
    _save_authorized_users()
    logger.info("User %s authorized successfully.", user_id)
    await update.message.reply_text(t("auth_success", lang), parse_mode="Markdown")

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
        # Try to resolve any newly registered chat_ids
        _resolve_chat_ids(records)
        user_lang = get_user_language(user_id)
        logger.info("Daily check for user %s (%d records, lang=%s).", user_id, len(records), user_lang)
        try:
            await reminder_scheduler.daily_check(
                bot=bot,
                records=records,
                owner_chat_id=user_id,
                timezone_str=TIMEZONE,
                reminder_days=REMINDER_DAYS_BEFORE,
                sent_reminders_path=_user_sent_reminders_path(user_id),
                default_language=user_lang,
            )
        except Exception as exc:
            logger.error("Daily check failed for user %s: %s", user_id, exc)


# ---------------------------------------------------------------------------
# Startup: load existing data for all authorized users
# ---------------------------------------------------------------------------


async def on_startup(application: Application) -> None:
    """Called once when the bot starts. Load authorized users and their Excel files."""
    global _scheduler, _authorized_users, _chat_id_registry

    # Load persistent user language preferences
    i18n.load_user_languages()
    logger.info("Startup: user language preferences loaded.")

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

    # Load chat ID registry
    _chat_id_registry = _load_chat_id_registry()
    logger.info("Startup: %d registered username(s) in chat ID registry.", len(_chat_id_registry))

    # Load each authorized user's Excel file (if it exists)
    for user_id in _authorized_users:
        excel_path = _user_excel_path(user_id)
        if excel_path.exists():
            try:
                records, invalid, _ = excel_handler.load_excel(str(excel_path))
                _user_records[user_id] = records
                # Auto-resolve chat_ids from registry
                resolved = _resolve_chat_ids(records)
                logger.info(
                    "Startup: loaded %d records for user %s (%d invalid, %d chat_ids resolved).",
                    len(records), user_id, invalid, resolved,
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
    app.add_handler(CommandHandler("menu", cmd_menu))
    app.add_handler(CommandHandler(["language", "lang", "burmese", "english"], cmd_language))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("check", cmd_check))
    app.add_handler(CommandHandler("runcheck", cmd_check))
    app.add_handler(CommandHandler("scheduler", cmd_scheduler))
    app.add_handler(CommandHandler("schedule", cmd_scheduler))
    app.add_handler(CommandHandler("scheduledata", cmd_scheduler))
    app.add_handler(CommandHandler("schedule_data", cmd_scheduler))
    app.add_handler(CommandHandler("delete", cmd_delete))
    app.add_handler(CommandHandler("removereminder", cmd_delete))
    app.add_handler(CommandHandler("history", cmd_history))
    app.add_handler(CommandHandler("clearsent", cmd_clearsent))
    app.add_handler(CommandHandler("resethistory", cmd_clearsent))
    app.add_handler(CommandHandler("users", cmd_users))
    app.add_handler(CallbackQueryHandler(handle_callback_query))
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
