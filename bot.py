"""
bot.py
------
Ariel's English Class — Telegram Payment Reminder Bot

A friendly, simple bot for teachers and students:
  - Teachers manage student payment deadlines via Excel (.xlsx).
  - Students tap /start once to register their username.
  - The bot automatically sends polite bilingual reminders 5 days before fees are due.
  - Teachers can view student status, trigger on-demand checks, and delete records via clean buttons.
"""

import datetime
import html
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
from telegram import BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, Update
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
import reminder_scheduler

# ---------------------------------------------------------------------------
# Load environment & Configuration
# ---------------------------------------------------------------------------
load_dotenv()

BOT_TOKEN: str = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
_raw_admin_id = os.getenv("ADMIN_TELEGRAM_ID", "").strip()
try:
    ADMIN_TELEGRAM_ID: int = int(_raw_admin_id) if _raw_admin_id else 7462033455
except ValueError:
    ADMIN_TELEGRAM_ID = 7462033455

# The single permanent administrator defined strictly by Railway / .env
ADMIN_ID: int = ADMIN_TELEGRAM_ID
ADMIN_IDS: set[int] = {ADMIN_ID}

TIMEZONE: str = os.getenv("TIMEZONE", "Asia/Yangon")
REMINDER_DAYS_BEFORE: int = int(os.getenv("REMINDER_DAYS_BEFORE", "5"))

# Daily automated check time (8:00 AM Asia/Yangon)
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

_user_records: dict[int, list[dict]] = {}
_authorized_users: set[int] = set()
_pending_auth: set[int] = set()
_chat_id_registry: dict[str, int] = {}
_scheduler: AsyncIOScheduler | None = None

PAGE_SIZE = 6

# ---------------------------------------------------------------------------
# Persistence helpers
# ---------------------------------------------------------------------------


def _load_authorized_users() -> set[int]:
    """Load authorized user IDs from JSON. Main admin is always included."""
    users = {ADMIN_ID}
    if AUTHORIZED_USERS_PATH.exists():
        try:
            with open(AUTHORIZED_USERS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                users.update(int(uid) for uid in data)
        except (json.JSONDecodeError, OSError, ValueError) as exc:
            logger.warning("Could not load authorized_users.json: %s", exc)
    users.add(ADMIN_ID)
    return users


def _save_authorized_users() -> None:
    """Save authorized user IDs to JSON."""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(AUTHORIZED_USERS_PATH, "w", encoding="utf-8") as f:
            json.dump(sorted(_authorized_users), f, indent=2)
    except OSError as exc:
        logger.error("Could not save authorized_users.json: %s", exc)


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
    """Register a student's @username -> chat_id mapping automatically."""
    if not username:
        return
    key = f"@{username.lower().lstrip('@')}"
    if _chat_id_registry.get(key) != chat_id:
        _chat_id_registry[key] = chat_id
        _save_chat_id_registry()
        logger.info("Registered chat_id %s for %s", chat_id, key)


def _resolve_chat_ids(records: list[dict]) -> int:
    """Fill in missing telegram_chat_id from the registry. Returns matched count."""
    resolved = 0
    for rec in records:
        if rec.get("telegram_chat_id"):
            continue
        username = rec.get("telegram_username")
        if not username:
            continue
        key = username.lower()
        chat_id = _chat_id_registry.get(key)
        if chat_id:
            rec["telegram_chat_id"] = chat_id
            resolved += 1
    return resolved


# ---------------------------------------------------------------------------
# Authorization Guards
# ---------------------------------------------------------------------------


def _is_admin(update: Update) -> bool:
    """Check if the user is the single permanent administrator configured in Railway."""
    return update.effective_user is not None and update.effective_user.id == ADMIN_ID


def _is_authorized(update: Update) -> bool:
    """Check if the user is an authorized teacher or the main administrator."""
    if update.effective_user is None:
        return False
    return update.effective_user.id == ADMIN_ID or update.effective_user.id in _authorized_users


# ---------------------------------------------------------------------------
# User-Friendly UI & Text Builders (HTML Mode)
# ---------------------------------------------------------------------------


def _build_main_menu_keyboard(is_admin: bool = False) -> InlineKeyboardMarkup:
    """Build a clean, user-friendly navigation keyboard."""
    buttons = [
        [
            InlineKeyboardButton("📋 Student Roster", callback_data="menu:roster"),
            InlineKeyboardButton("📨 Send Reminders Now", callback_data="menu:check"),
        ],
        [
            InlineKeyboardButton("🗑️ Remove Student", callback_data="menu:delete_list:0"),
            InlineKeyboardButton("📖 How to Use", callback_data="menu:help"),
        ],
    ]
    if is_admin:
        buttons.append(
            [InlineKeyboardButton("👥 Manage Teachers", callback_data="menu:teachers")]
        )
    buttons.append(
        [InlineKeyboardButton("🔄 Refresh Dashboard", callback_data="menu:home")]
    )
    return InlineKeyboardMarkup(buttons)


def _render_teachers_list() -> tuple[str, InlineKeyboardMarkup]:
    """Render the list of teachers and access control buttons for administrators."""
    id_to_uname = {cid: uname for uname, cid in _chat_id_registry.items()}
    admin_uname = id_to_uname.get(ADMIN_ID, "Admin")

    lines = [
        "👥 <b>Teacher & Administrator Access</b>",
        "━━━━━━━━━━━━━━━━━━━━━",
        f"👑 <b>Permanent Administrator:</b>\n  • <code>{ADMIN_ID}</code> ({admin_uname})",
        "",
        "🎓 <b>Authorized Teachers:</b>",
    ]

    other_teachers = sorted([uid for uid in _authorized_users if uid != ADMIN_ID])
    if other_teachers:
        for uid in other_teachers:
            uname = id_to_uname.get(uid, "Teacher")
            lines.append(f"  • <code>{uid}</code> ({uname})")
    else:
        lines.append("  <i>No additional teachers authorized yet.</i>")

    lines.append("\n━━━━━━━━━━━━━━━━━━━━━")
    lines.append("➕ <b>To add a new teacher:</b>")
    lines.append("Type: <code>/authorize &lt;Telegram_ID&gt;</code>")
    lines.append("<i>(or /authorize @username if they tapped Start)</i>")

    buttons = []
    for uid in other_teachers:
        uname = id_to_uname.get(uid, str(uid))
        buttons.append([
            InlineKeyboardButton(f"❌ Revoke {uname} ({uid})", callback_data=f"teacher:revoke:{uid}")
        ])

    buttons.append([InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")])
    return "\n".join(lines), InlineKeyboardMarkup(buttons)


def _get_home_text(user_id: int, user_name: str) -> str:
    """Friendly, modern dashboard summary for teachers."""
    records = _user_records.get(user_id, [])
    _resolve_chat_ids(records)
    total = len(records)
    linked = sum(1 for r in records if r.get("telegram_chat_id"))
    missing = total - linked

    safe_name = html.escape(user_name)
    lines = [
        "🎓 <b>Ariel's English Class</b>",
        "━━━━━━━━━━━━━━━━━━━━━",
        f"👋 Welcome, <b>{safe_name}</b>!\n",
        "📊 <b>Class Overview:</b>",
        f"• Total Students: <b>{total}</b>",
        f"• Connected on Telegram: <b>{linked}</b> ✅",
    ]
    if missing > 0:
        lines.append(f"• Waiting to Tap Start: <b>{missing}</b> ⚠️")

    lines += [
        "",
        "⏰ <b>Reminder Automation:</b>",
        "Automated checks run daily for students due within <b>5 days</b>.",
        "",
        "💡 <i>Tip: Send or drag & drop any <code>.xlsx</code> file anytime to update your roster.</i>",
        "Choose an action below:",
    ]
    return "\n".join(lines)


def _render_roster_text(user_id: int) -> str:
    """Clean, beautifully organized summary of student roster and status."""
    records = _user_records.get(user_id, [])
    total = len(records)
    if total == 0:
        return (
            "📋 <b>Your Student Roster</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━\n\n"
            "<i>No student roster loaded yet.</i>\n\n"
            "📤 <b>To get started:</b>\n"
            "Simply send or attach your Excel file (<code>.xlsx</code>) right here in this chat!"
        )

    _resolve_chat_ids(records)
    linked = sum(1 for r in records if r.get("telegram_chat_id"))
    missing = total - linked

    lines = [
        "📋 <b>Student Roster</b>",
        "━━━━━━━━━━━━━━━━━━━━━",
        f"📊 <b>Summary:</b> <b>{total}</b> total  •  ✅ <b>{linked}</b> ready  •  ⚠️ <b>{missing}</b> pending\n",
    ]

    today = datetime.datetime.now(ZoneInfo(TIMEZONE)).date()
    sorted_recs = sorted(records, key=lambda r: r.get("due_date") or datetime.date.max)

    for i, r in enumerate(sorted_recs[:15], 1):
        uname = html.escape(str(r.get("telegram_username", "")))
        sname = html.escape(str(r.get("student_name", uname)))
        due = r.get("due_date")
        has_chat = bool(r.get("telegram_chat_id"))

        if due:
            days = (due - today).days
            if days == 0:
                due_info = "Due Today 🚨"
            elif days == 1:
                due_info = f"Due tomorrow ({due})"
            elif days > 1:
                due_info = f"Due in {days}d ({due})"
            else:
                due_info = f"Overdue ({due})"
        else:
            due_info = "No date set"

        status_icon = "✅" if has_chat else "⚠️"
        lines.append(f"<b>{i}. {sname}</b> (<code>{uname}</code>)")
        lines.append(f"   {status_icon} <i>{due_info}</i>\n")

    if len(sorted_recs) > 15:
        lines.append(f"<i>… and {len(sorted_recs) - 15} more student(s).</i>\n")

    if missing > 0:
        lines.append("💡 <i>Students marked with ⚠️ simply need to open @ariel_english_bot and tap Start.</i>")

    return "\n".join(lines).strip()


# Alias for backward compatibility
_render_status_text = _render_roster_text


def _render_help_text() -> str:
    """Clear, step-by-step user guide."""
    return (
        "📖 <b>How to Use Ariel's Reminder Bot</b>\n"
        "━━━━━━━━━━━━━━━━━━━━━\n\n"
        "<b>1️⃣ Prepare Your Spreadsheet (.xlsx)</b>\n"
        "Your Excel file only needs 3 columns:\n"
        "• <code>student_name</code> — (e.g. <i>Thiri Hlaing</i>)\n"
        "• <code>telegram_username</code> — (e.g. <i>@thiri_test</i>)\n"
        "• <code>due_date</code> — (format: <i>YYYY-MM-DD</i>, e.g. <i>2026-10-25</i>)\n\n"
        "<b>2️⃣ Upload Anytime</b>\n"
        "Simply send or drag & drop your <code>.xlsx</code> file directly into this chat anytime! The bot automatically syncs it.\n\n"
        "<b>3️⃣ Student Linking</b>\n"
        "Tell your students to search for <b>@ariel_english_bot</b> and tap <b>Start</b>. The bot automatically links their account!\n\n"
        "<b>4️⃣ Automated Reminders</b>\n"
        "The bot automatically checks due dates and sends polite bilingual reminders (English + Myanmar) 5 days before payment is due."
    )


def _render_delete_list(
    user_id: int, page: int = 0
) -> tuple[str, InlineKeyboardMarkup]:
    """Render list of students with clean one-tap remove buttons."""
    records = _user_records.get(user_id, [])
    if not records:
        text = (
            "🗑️ <b>Remove a Student</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━\n\n"
            "<i>No students currently in your roster.</i>\n\n"
            "Upload an Excel file to add students."
        )
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
        )
        return text, keyboard

    total_records = len(records)
    total_pages = max(1, (total_records + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(0, min(page, total_pages - 1))

    start_idx = page * PAGE_SIZE
    end_idx = min(start_idx + PAGE_SIZE, total_records)
    page_records = records[start_idx:end_idx]

    lines = [
        "🗑️ <b>Remove a Student</b>",
        "━━━━━━━━━━━━━━━━━━━━━",
        f"Showing <b>{start_idx+1}–{end_idx}</b> of <b>{total_records}</b> student(s).",
        "Select a student below to remove them:\n",
    ]

    buttons = []
    for i, rec in enumerate(page_records, start=start_idx):
        uname = html.escape(str(rec.get("telegram_username", f"Student {i+1}")))
        sname = html.escape(str(rec.get("student_name", uname)))
        due = rec.get("due_date", "No date")
        lines.append(f"<b>{i+1}. {sname}</b> (<code>{uname}</code>) • Due: <code>{due}</code>")
        label = f"❌ Remove {sname}"[:30]
        buttons.append(
            [InlineKeyboardButton(label, callback_data=f"del:ask:{i}:{page}")]
        )

    # Pagination navigation
    nav_buttons = []
    if page > 0:
        nav_buttons.append(
            InlineKeyboardButton("⬅️ Prev", callback_data=f"menu:delete_list:{page-1}")
        )
    if total_pages > 1:
        nav_buttons.append(
            InlineKeyboardButton(
                f"📄 {page+1}/{total_pages}", callback_data=f"menu:delete_list:{page}"
            )
        )
    if page < total_pages - 1:
        nav_buttons.append(
            InlineKeyboardButton("Next ➡️", callback_data=f"menu:delete_list:{page+1}")
        )
    if nav_buttons:
        buttons.append(nav_buttons)

    buttons.append(
        [InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]
    )
    return "\n".join(lines), InlineKeyboardMarkup(buttons)



# ---------------------------------------------------------------------------
# Message Editor Helper (HTML Mode with Fallback)
# ---------------------------------------------------------------------------


async def _safe_edit_message(
    query,
    text: str,
    reply_markup: InlineKeyboardMarkup | None = None,
    parse_mode: str | None = "HTML",
) -> None:
    """Safely edit query message. Falls back to plain text if HTML parsing ever fails."""
    async def _do_edit(pm):
        if hasattr(query, "edit_message_text"):
            return await query.edit_message_text(text, reply_markup=reply_markup, parse_mode=pm)
        elif hasattr(query, "edit_text"):
            return await query.edit_text(text, reply_markup=reply_markup, parse_mode=pm)
        elif hasattr(query, "message") and hasattr(query.message, "edit_text"):
            return await query.message.edit_text(text, reply_markup=reply_markup, parse_mode=pm)

    try:
        await _do_edit(parse_mode)
    except BadRequest as exc:
        err = str(exc)
        if "Message is not modified" in err:
            return
        logger.warning("Parse error in edit_message_text (%s). Trying plain text fallback...", err)
        try:
            await _do_edit(None)
        except BadRequest as inner_exc:
            if "Message is not modified" not in str(inner_exc):
                logger.error("Failed to edit message in fallback: %s", inner_exc)
    except Exception as exc:
        logger.error("Unexpected error in _safe_edit_message: %s", exc)


# ---------------------------------------------------------------------------
# Command Handlers
# ---------------------------------------------------------------------------


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/start — Welcome teachers with Dashboard, or welcome students politely."""
    user = update.effective_user
    chat_id = update.effective_chat.id
    user_id = user.id
    name = user.first_name if user else "there"

    # Always capture student's @username -> chat_id automatically
    _register_chat_id(user.username, chat_id)

    if _is_authorized(update):
        # Authorized Teacher / Admin -> Show clean Control Panel
        text = _get_home_text(user_id, name)
        keyboard = _build_main_menu_keyboard(_is_admin(update))
        try:
            await update.message.reply_text(
                text, reply_markup=keyboard, parse_mode="HTML"
            )
        except BadRequest:
            await update.message.reply_text(
                text, reply_markup=keyboard, parse_mode=None
            )
    else:
        # Student -> Clean, polite welcome message (no developer jargon)
        safe_name = html.escape(name)
        text = (
            f"👋 <b>Hello, {safe_name}!</b>\n\n"
            "Welcome to <b>Ariel's English Class</b>! 🤍\n\n"
            "✅ <b>You're all set!</b>\n"
            "You will receive a gentle reminder here whenever your monthly tuition fee is coming up.\n\n"
            "No further action is needed from you. Have a wonderful day! ✨"
        )
        await update.message.reply_text(text, parse_mode="HTML")


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/help — Show how to use the bot."""
    if _is_authorized(update):
        text = _render_help_text()
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Back to Dashboard", callback_data="menu:home")]]
        )
        await update.message.reply_text(text, reply_markup=keyboard, parse_mode="HTML")
    else:
        await cmd_start(update, context)


async def cmd_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/menu — Open main dashboard."""
    if _is_authorized(update):
        await cmd_start(update, context)
    else:
        await update.message.reply_text(
            "🔒 <b>Teacher Access Required</b>\n\n"
            "The Menu & Control Panel is reserved for class administrators.\n\n"
            "To authorize this chat as a teacher, type:\n"
            "<code>/auth &lt;Admin_ID&gt;</code>\n"
            "or simply reply with your Admin ID.",
            parse_mode="HTML",
        )


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/status or /roster — View student roster."""
    if not _is_authorized(update):
        await update.message.reply_text("❌ You are not authorized. Send /start to begin.")
        return

    user_id = update.effective_user.id
    text = _render_roster_text(user_id)
    keyboard = InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("📨 Send Reminders Now", callback_data="menu:check")],
            [InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")],
        ]
    )
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="HTML")


def _build_check_report(res: dict, today: datetime.date) -> str:
    """Build a clean, user-friendly reminder check report."""
    lines = [
        "📨 <b>Reminder Check Report</b>",
        "━━━━━━━━━━━━━━━━━━━━━",
        f"📅 <b>Date:</b> <code>{today}</code>  •  🎯 <b>Window:</b> Next 5 days\n",
    ]

    if res.get("sent"):
        lines.append(f"✅ <b>Delivered Reminders ({len(res['sent'])}):</b>")
        for r in res["sent"][:10]:
            sname = html.escape(str(r.get("student_name", "")))
            uname = html.escape(str(r.get("telegram_username", "")))
            lines.append(f"  • <b>{sname}</b> (<code>{uname}</code>) — Due: <code>{r.get('due_date')}</code>")
        lines.append("")

    if res.get("already_sent"):
        lines.append(f"ℹ️ <b>Already Sent Earlier Today ({len(res['already_sent'])}):</b>")
        for r in res["already_sent"][:10]:
            sname = html.escape(str(r.get("student_name", "")))
            lines.append(f"  • {sname} — Due: <code>{r.get('due_date')}</code>")
        lines.append("")

    if res.get("not_due"):
        lines.append(f"⏳ <b>Not Due Yet ({len(res['not_due'])}):</b>")
        for r in res["not_due"][:6]:
            sname = html.escape(str(r.get("student_name", "")))
            days = r.get("days_until_due", "?")
            lines.append(f"  • {sname} — in {days} days")
        if len(res["not_due"]) > 6:
            lines.append(f"  <i>… and {len(res['not_due']) - 6} more.</i>")
        lines.append("")

    if res.get("no_chat_id"):
        lines.append(f"⚠️ <b>Waiting to Tap Start ({len(res['no_chat_id'])}):</b>")
        for r in res["no_chat_id"][:10]:
            uname = html.escape(str(r.get("telegram_username", "")))
            lines.append(f"  • <code>{uname}</code> (Cannot deliver until student taps Start)")
        lines.append("")

    if res.get("failed"):
        lines.append(f"❌ <b>Delivery Issues ({len(res['failed'])}):</b>")
        for r, err in res["failed"][:5]:
            uname = html.escape(str(r.get("telegram_username", "Student")))
            lines.append(f"  • <code>{uname}</code>: {html.escape(str(err))}")
        lines.append("")

    if not res.get("sent") and not res.get("already_sent"):
        lines.append("🎉 <i>No student reminders were pending right now!</i>")

    return "\n".join(lines).strip()


async def cmd_check(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/check — Manually send reminders for students due right now."""
    if not _is_authorized(update):
        await update.message.reply_text("❌ You are not authorized. Send /start to begin.")
        return

    user_id = update.effective_user.id
    records = _user_records.get(user_id, [])
    if not records:
        await update.message.reply_text(
            "ℹ️ <i>No student records loaded. Upload an Excel file first.</i>",
            parse_mode="HTML",
        )
        return

    _resolve_chat_ids(records)
    today = datetime.datetime.now(ZoneInfo(TIMEZONE)).date()
    msg = await update.message.reply_text("⏳ Checking reminder schedule…")

    try:
        res = await reminder_scheduler.daily_check(
            bot=context.bot,
            records=records,
            owner_chat_id=user_id,
            timezone_str=TIMEZONE,
            reminder_days=REMINDER_DAYS_BEFORE,
            sent_reminders_path=_user_sent_reminders_path(user_id),
        )
    except Exception as exc:
        await msg.edit_text(f"❌ Check failed: {exc}")
        return

    report_text = _build_check_report(res, today)
    keyboard = InlineKeyboardMarkup(
        [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
    )
    await msg.edit_text(report_text, reply_markup=keyboard, parse_mode="HTML")


async def cmd_delete(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/delete — Interactive deletion menu."""
    if not _is_authorized(update):
        await update.message.reply_text("❌ You are not authorized. Send /start to begin.")
        return

    user_id = update.effective_user.id
    text, keyboard = _render_delete_list(user_id, page=0)
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="HTML")


async def cmd_authorize(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/authorize <user_id> or /authorize @username — Main admin grants teacher access."""
    if not _is_admin(update):
        await update.message.reply_text("❌ Only the bot administrator can authorize new teachers.")
        return

    args = context.args or []
    if not args:
        await update.message.reply_text(
            "<b>Usage:</b> <code>/authorize &lt;Telegram_User_ID&gt;</code>\n\n"
            "<i>Tip: Ask the teacher to tap /start in the bot, or get their numeric ID from @userinfobot.</i>",
            parse_mode="HTML",
        )
        return

    raw_id = args[0].strip()
    target_id: int | None = None

    if raw_id.startswith("@"):
        target_username = raw_id.lower()
        target_id = _chat_id_registry.get(target_username)
        if not target_id:
            await update.message.reply_text(
                f"❌ <code>{html.escape(raw_id)}</code> hasn't tapped /start yet.\n"
                "Please have them send /start to @ariel_english_bot first, or use their numeric ID.",
                parse_mode="HTML",
            )
            return
    else:
        try:
            target_id = int(raw_id)
        except ValueError:
            await update.message.reply_text("❌ Please provide a valid numeric Telegram ID or @username.")
            return

    _authorized_users.add(target_id)
    _save_authorized_users()

    # Reverse lookup username
    id_to_uname = {cid: uname for uname, cid in _chat_id_registry.items()}
    uname = id_to_uname.get(target_id, "Teacher")

    await update.message.reply_text(
        f"✅ <b>Teacher Authorized!</b>\n\n"
        f"• <b>Telegram ID:</b> <code>{target_id}</code>\n"
        f"• <b>Username:</b> <code>{uname}</code>\n\n"
        "They now have full teacher privileges to upload rosters and send reminders.",
        parse_mode="HTML",
    )

    # Attempt to notify the authorized teacher directly
    try:
        await context.bot.send_message(
            chat_id=target_id,
            text=(
                "🎉 <b>Teacher Access Granted!</b>\n\n"
                "The administrator has authorized your account.\n"
                "Send /menu to open your class control panel!"
            ),
            parse_mode="HTML",
        )
    except Exception:
        pass


async def cmd_revoke(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/revoke <user_id> — Main admin removes teacher access."""
    if not _is_admin(update):
        await update.message.reply_text("❌ Only the bot administrator can revoke teacher access.")
        return

    args = context.args or []
    if not args:
        await update.message.reply_text(
            "<b>Usage:</b> <code>/revoke &lt;Telegram_User_ID&gt;</code>",
            parse_mode="HTML",
        )
        return

    try:
        target_id = int(args[0].strip().lstrip("@"))
    except ValueError:
        raw_user = args[0].strip().lower()
        target_id = _chat_id_registry.get(raw_user if raw_user.startswith("@") else f"@{raw_user}")
        if not target_id:
            await update.message.reply_text("❌ Please provide a valid numeric Telegram ID.")
            return

    if target_id == ADMIN_ID:
        await update.message.reply_text("❌ Cannot revoke the permanent administrator.")
        return

    if target_id in _authorized_users:
        _authorized_users.discard(target_id)
        _save_authorized_users()
        await update.message.reply_text(f"✅ Teacher access for <code>{target_id}</code> has been revoked.", parse_mode="HTML")
    else:
        await update.message.reply_text(f"ℹ️ User <code>{target_id}</code> is not in the authorized teachers list.", parse_mode="HTML")


async def cmd_teachers(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/teachers — View list of teachers and access control."""
    if not _is_admin(update):
        await update.message.reply_text("❌ Only the bot administrator can view the teachers management list.")
        return

    text, keyboard = _render_teachers_list()
    await update.message.reply_text(text, reply_markup=keyboard, parse_mode="HTML")


async def cmd_auth(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/auth — Admin helper command (redirects to /authorize)."""
    if _is_admin(update):
        if context.args:
            await cmd_authorize(update, context)
        else:
            await update.message.reply_text("👑 <b>Admin Usage:</b> <code>/authorize &lt;Telegram_ID&gt;</code> to grant teacher access.", parse_mode="HTML")
    else:
        await update.message.reply_text("🔒 Only the permanent administrator can grant teacher access.", parse_mode="HTML")


# ---------------------------------------------------------------------------
# Callback Query Handler (Interactive Buttons)
# ---------------------------------------------------------------------------


async def handle_callback_query(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Handle all button clicks with 100% reliability and HTML formatting."""
    query = update.callback_query
    if not query:
        return

    try:
        await query.answer()
    except Exception as exc:
        logger.debug("query.answer() exception: %s", exc)

    user_id = query.from_user.id
    if user_id not in ADMIN_IDS and user_id not in _authorized_users:
        logger.warning("Unauthorized button click from %s", user_id)
        await _safe_edit_message(query, "❌ You are not authorized. Send /start to begin.")
        return

    is_admin = user_id in ADMIN_IDS
    data = query.data or ""

    if data == "menu:home":
        first_name = query.from_user.first_name if query.from_user else "there"
        text = _get_home_text(user_id, first_name)
        keyboard = _build_main_menu_keyboard(is_admin)
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data in ("menu:roster", "menu:status"):
        text = _render_roster_text(user_id)
        keyboard = InlineKeyboardMarkup(
            [
                [InlineKeyboardButton("📨 Send Reminders Now", callback_data="menu:check")],
                [InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")],
            ]
        )
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data == "menu:check":
        records = _user_records.get(user_id, [])
        if not records:
            text = "ℹ️ <i>No student records loaded. Upload an Excel file first.</i>"
            keyboard = InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
            )
            await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")
            return

        _resolve_chat_ids(records)
        today = datetime.datetime.now(ZoneInfo(TIMEZONE)).date()

        try:
            res = await reminder_scheduler.daily_check(
                bot=context.bot,
                records=records,
                owner_chat_id=user_id,
                timezone_str=TIMEZONE,
                reminder_days=REMINDER_DAYS_BEFORE,
                sent_reminders_path=_user_sent_reminders_path(user_id),
            )
        except Exception as exc:
            text = f"❌ Check failed: {exc}"
            keyboard = InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
            )
            await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")
            return

        report_text = _build_check_report(res, today)
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
        )
        await _safe_edit_message(query, report_text, reply_markup=keyboard, parse_mode="HTML")

    elif data == "menu:help":
        text = _render_help_text()
        keyboard = InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")]]
        )
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data.startswith("menu:delete_list:"):
        page = int(data.split(":")[2])
        text, keyboard = _render_delete_list(user_id, page)
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data.startswith("del:ask:"):
        parts = data.split(":")
        idx = int(parts[2])
        page = int(parts[3])
        records = _user_records.get(user_id, [])
        if 0 <= idx < len(records):
            rec = records[idx]
            uname = html.escape(str(rec.get("telegram_username", "")))
            sname = html.escape(str(rec.get("student_name", uname)))
            due = rec.get("due_date")
            text = (
                "⚠️ <b>Confirm Student Removal</b>\n"
                "━━━━━━━━━━━━━━━━━━━━━\n\n"
                "Are you sure you want to remove this student?\n\n"
                f"👤 <b>Name:</b> <b>{sname}</b>\n"
                f"💬 <b>Username:</b> <code>{uname}</code>\n"
                f"📅 <b>Payment Due:</b> <code>{due}</code>\n\n"
                "<i>This will update your Excel roster and stop reminders for this student.</i>"
            )
            keyboard = InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "🗑️ Yes, Remove", callback_data=f"del:confirm:{idx}:{page}"
                        ),
                        InlineKeyboardButton(
                            "↩️ Cancel", callback_data=f"menu:delete_list:{page}"
                        ),
                    ],
                    [InlineKeyboardButton("⬅️ Back to Menu", callback_data="menu:home")],
                ]
            )
            await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data.startswith("del:confirm:"):
        parts = data.split(":")
        idx = int(parts[2])
        page = int(parts[3])
        records = _user_records.get(user_id, [])
        if 0 <= idx < len(records):
            removed = records.pop(idx)
            excel_path = _user_excel_path(user_id)
            excel_handler.save_excel(records, str(excel_path))
            sname = html.escape(str(removed.get("student_name", "Student")))
            uname = html.escape(str(removed.get("telegram_username", "")))
            try:
                await query.answer(f"✅ Removed {sname}!", show_alert=False)
            except Exception:
                pass

            text, keyboard = _render_delete_list(user_id, page)
            text = f"✅ <b>Removed {sname} ({uname}) successfully.</b>\n\n" + text
            await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data == "menu:teachers":
        if not is_admin:
            await query.answer("❌ Admin access required.", show_alert=True)
            return
        text, keyboard = _render_teachers_list()
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    elif data.startswith("teacher:revoke:"):
        if not is_admin:
            await query.answer("❌ Admin access required.", show_alert=True)
            return
        target_id = int(data.split(":")[2])
        if target_id in ADMIN_IDS:
            await query.answer("❌ Cannot revoke permanent admin.", show_alert=True)
            return
        _authorized_users.discard(target_id)
        _save_authorized_users()
        await query.answer(f"✅ Revoked {target_id}!", show_alert=False)
        text, keyboard = _render_teachers_list()
        text = f"✅ <b>Revoked teacher {target_id} successfully.</b>\n\n" + text
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")

    else:
        # Handles any legacy or unrecognized buttons by safely returning to home
        logger.info("Unrecognized callback '%s', returning to home.", data)
        first_name = query.from_user.first_name if query.from_user else "there"
        text = _get_home_text(user_id, first_name)
        keyboard = _build_main_menu_keyboard(is_admin)
        await _safe_edit_message(query, text, reply_markup=keyboard, parse_mode="HTML")


# ---------------------------------------------------------------------------
# Excel Upload Handler
# ---------------------------------------------------------------------------


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle .xlsx file uploads from authorized teachers."""
    if not _is_authorized(update):
        await update.message.reply_text("❌ You are not authorized. Send /start to begin.")
        return

    doc = update.message.document
    if not doc:
        return

    fname = doc.file_name or ""
    if not fname.lower().endswith(".xlsx"):
        await update.message.reply_text(
            "❌ Please upload a valid <b>.xlsx</b> Excel file.",
            parse_mode="HTML",
        )
        return

    await update.message.reply_text("⏳ Processing your student roster…")
    user_id = update.effective_user.id

    tmp_path = Path(tempfile.mktemp(suffix=".xlsx"))
    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(custom_path=str(tmp_path))
    except TelegramError as exc:
        logger.error("Failed to download file from user %s: %s", user_id, exc)
        await update.message.reply_text(f"❌ Could not download file: {exc}")
        return

    try:
        valid_records, invalid_count, errors = excel_handler.load_excel(str(tmp_path))
    except (FileNotFoundError, ValueError) as exc:
        tmp_path.unlink(missing_ok=True)
        await update.message.reply_text(f"❌ Validation error:\n{exc}")
        return

    if len(valid_records) == 0:
        tmp_path.unlink(missing_ok=True)
        await update.message.reply_text(
            "❌ The file contains no valid student records. Your previous list has been kept."
        )
        return

    # Replace saved Excel file
    dest = _user_excel_path(user_id)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(tmp_path), str(dest))
    _user_records[user_id] = valid_records

    # Auto-resolve chat IDs
    resolved = _resolve_chat_ids(valid_records)
    unresolved = sum(1 for r in valid_records if not r.get("telegram_chat_id"))

    lines = [
        "✅ <b>Student Roster Updated!</b>",
        "━━━━━━━━━━━━━━━━━━━━━",
        "📊 <b>Upload Summary:</b>",
        f"• Total Students Loaded: <b>{len(valid_records)}</b>",
        f"• Connected on Telegram: <b>{resolved}</b> ✅",
    ]
    if unresolved > 0:
        lines.append(f"• Waiting to Tap Start: <b>{unresolved}</b> ⚠️")
        missing_names = [
            f"<code>{html.escape(r.get('telegram_username', '?'))}</code>"
            for r in valid_records
            if not r.get("telegram_chat_id")
        ][:5]
        lines.append(f"  <i>(Pending: {', '.join(missing_names)})</i>")

    lines.append("\n⏰ <i>Automated reminders will run for students due within 5 days.</i>")
    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("📋 View Updated Roster", callback_data="menu:roster"),
                InlineKeyboardButton("⬅️ Dashboard", callback_data="menu:home"),
            ]
        ]
    )
    await update.message.reply_text("\n".join(lines), reply_markup=keyboard, parse_mode="HTML")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle plain text messages."""
    user = update.effective_user
    chat_id = update.effective_chat.id
    if user:
        _register_chat_id(user.username, chat_id)

    if _is_authorized(update):
        await update.message.reply_text(
            "💡 <i>Tip: Send an Excel file (<code>.xlsx</code>) anytime to update your student roster, or type /menu to open your dashboard.</i>",
            parse_mode="HTML",
        )
    else:
        await update.message.reply_text(
            "👋 Hello! Your account is registered with Ariel's English Class. You will receive gentle reminders here whenever tuition is due. ✨",
            parse_mode="HTML",
        )


# ---------------------------------------------------------------------------
# Scheduled Daily Job
# ---------------------------------------------------------------------------


async def _run_daily_check(bot) -> None:
    """Run automated check for all authorized teachers."""
    if not _user_records:
        return

    for user_id, records in _user_records.items():
        if not records:
            continue
        _resolve_chat_ids(records)
        try:
            res = await reminder_scheduler.daily_check(
                bot=bot,
                records=records,
                owner_chat_id=user_id,
                timezone_str=TIMEZONE,
                reminder_days=REMINDER_DAYS_BEFORE,
                sent_reminders_path=_user_sent_reminders_path(user_id),
            )
            if res.get("sent"):
                sent_names = [
                    html.escape(str(r.get("student_name", r.get("telegram_username", "Student"))))
                    for r in res["sent"]
                ]
                await bot.send_message(
                    chat_id=user_id,
                    text=(
                        f"⏰ <b>Automated Reminder Sent ({len(res['sent'])}):</b>\n"
                        + "\n".join(f"• {name}" for name in sent_names)
                    ),
                    parse_mode="HTML",
                )
        except Exception as exc:
            logger.error("Daily check failed for user %s: %s", user_id, exc)


# ---------------------------------------------------------------------------
# Startup & Main
# ---------------------------------------------------------------------------


async def on_startup(application: Application) -> None:
    """Load authorized users, chat ID registry, and student rosters."""
    global _scheduler, _authorized_users, _chat_id_registry

    _authorized_users = _load_authorized_users()
    _chat_id_registry = _load_chat_id_registry()

    # Register commands for the Telegram Menu button (bottom-left [Menu] button in chat)
    try:
        await application.bot.set_my_commands(
            [
                BotCommand("menu", "📋 Open Teacher Control Panel"),
                BotCommand("roster", "👥 View Student Roster & Due Dates"),
                BotCommand("check", "📨 Check & Send Reminders Now"),
                BotCommand("delete", "🗑️ Remove Student from Roster"),
                BotCommand("help", "📖 How to Use Ariel's Bot"),
                BotCommand("teachers", "👥 Manage Teachers (Admin)"),
                BotCommand("start", "🔄 Welcome / Restart Bot"),
            ]
        )
        logger.info("Bot commands successfully registered with Telegram.")
    except Exception as exc:
        logger.warning("Could not register bot commands: %s", exc)

    for user_id in _authorized_users:
        excel_path = _user_excel_path(user_id)
        if excel_path.exists():
            try:
                records, _, _ = excel_handler.load_excel(str(excel_path))
                _user_records[user_id] = records
                _resolve_chat_ids(records)
            except Exception as exc:
                logger.warning("Could not load Excel for user %s: %s", user_id, exc)

    # Start APScheduler: Production mode (daily 8:00 AM) or Test mode (interval)
    check_interval = int(os.getenv("CHECK_INTERVAL_MINUTES", "0"))
    _scheduler = AsyncIOScheduler(timezone=ZoneInfo(TIMEZONE))

    if check_interval > 0:
        # Test Mode (runs every X minutes)
        _scheduler.add_job(
            _run_daily_check,
            trigger="interval",
            minutes=check_interval,
            kwargs={"bot": application.bot},
            id="automated_reminder_check",
            replace_existing=True,
        )
        logger.info("Scheduler started in TEST MODE: checking every %d minutes (%s).", check_interval, TIMEZONE)
    else:
        # Production Mode (runs once daily at 8:00 AM Asia/Yangon)
        _scheduler.add_job(
            _run_daily_check,
            trigger="cron",
            hour=REMINDER_HOUR,
            minute=REMINDER_MINUTE,
            kwargs={"bot": application.bot},
            id="automated_reminder_check",
            replace_existing=True,
        )
        logger.info("Scheduler started in PRODUCTION MODE: checking daily at %02d:%02d (%s).", REMINDER_HOUR, REMINDER_MINUTE, TIMEZONE)

    _scheduler.start()


def main() -> None:
    setup_logging()

    if not BOT_TOKEN:
        logger.critical("TELEGRAM_BOT_TOKEN is not set. Exiting.")
        raise SystemExit(1)

    logger.info("Starting Ariel's English Class Reminder Bot…")

    app = Application.builder().token(BOT_TOKEN).post_init(on_startup).build()

    # Commands
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("menu", cmd_menu))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("roster", cmd_status))
    app.add_handler(CommandHandler("check", cmd_check))
    app.add_handler(CommandHandler("delete", cmd_delete))
    app.add_handler(CommandHandler("auth", cmd_auth))
    app.add_handler(CommandHandler("authorize", cmd_authorize))
    app.add_handler(CommandHandler("revoke", cmd_revoke))
    app.add_handler(CommandHandler("teachers", cmd_teachers))

    # Interactive buttons
    app.add_handler(CallbackQueryHandler(handle_callback_query))

    # Document upload
    app.add_handler(
        MessageHandler(filters.Document.ALL & filters.ChatType.PRIVATE, handle_document)
    )

    # Fallback text input
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_text)
    )

    # Run polling with drop_pending_updates=True to avoid stuck update loops
    try:
        app.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot shutting down…")
    finally:
        if _scheduler and _scheduler.running:
            _scheduler.shutdown(wait=False)
        logger.info("Bot stopped.")


if __name__ == "__main__":
    main()
