"""
i18n.py
-------
Internationalization (i18n) module for the Telegram Payment Reminder Bot.
Supports Burmese (မြန်မာဘာသာ) and English.
"""

import json
import logging
import os
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

DATA_DIR = Path("data")
LANGUAGES_FILE = DATA_DIR / "user_languages.json"

LANG_EN = "en"
LANG_MY = "my"

SUPPORTED_LANGUAGES = {
    LANG_EN: "🇬🇧 English",
    LANG_MY: "🇲🇲 မြန်မာ (Burmese)",
}

# ---------------------------------------------------------------------------
# In-memory language cache & persistence
# ---------------------------------------------------------------------------
_user_languages: dict[int, str] = {}


def load_user_languages() -> None:
    """Load user language preferences from disk."""
    global _user_languages
    if not LANGUAGES_FILE.exists():
        _user_languages = {}
        return
    try:
        with open(LANGUAGES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                _user_languages = {int(k): str(v) for k, v in data.items()}
    except (json.JSONDecodeError, OSError, ValueError) as exc:
        logger.warning("Could not load user_languages.json: %s", exc)
        _user_languages = {}


def save_user_languages() -> None:
    """Save user language preferences atomically."""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        temp_path = LANGUAGES_FILE.with_suffix(f".tmp.{os.getpid()}")
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(_user_languages, f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, LANGUAGES_FILE)
    except OSError as exc:
        logger.error("Could not save user_languages.json: %s", exc)


def get_user_language(user_id: int) -> str:
    """Get language for a user (defaults to Burmese 'my')."""
    if not _user_languages and LANGUAGES_FILE.exists():
        load_user_languages()
    return _user_languages.get(user_id, LANG_MY)


def set_user_language(user_id: int, lang: str) -> str:
    """Set language for a user ('en' or 'my'). Returns normalized code."""
    normalized = LANG_EN if lang.lower().startswith("en") else LANG_MY
    _user_languages[user_id] = normalized
    save_user_languages()
    return normalized


# ---------------------------------------------------------------------------
# Translations Dictionary
# ---------------------------------------------------------------------------
MESSAGES: dict[str, dict[str, str]] = {
    # Buttons
    "btn_roster_status": {
        "en": "📊 Roster & Status",
        "my": "📊 ကျောင်းသားစာရင်းနှင့် အခြေအနေ",
    },
    "btn_scheduler": {
        "en": "⚙️ Scheduler",
        "my": "⚙️ အချိန်ဇယား စောင့်ဆိုင်းစာရင်း",
    },
    "btn_check_send": {
        "en": "📨 Check & Send Now",
        "my": "📨 ယခုစစ်ဆေးပြီး ပို့မည်",
    },
    "btn_delete_reminders": {
        "en": "🗑️ Delete Reminders",
        "my": "🗑️ သတိပေးချက် ပယ်ဖျက်မည်",
    },
    "btn_sent_history": {
        "en": "📜 Sent History",
        "my": "📜 ပို့ပြီးသော မှတ်တမ်း",
    },
    "btn_language": {
        "en": "🌐 Language",
        "my": "🌐 ဘာသာစကား",
    },
    "btn_refresh_menu": {
        "en": "🔄 Refresh Menu",
        "my": "🔄 မီနူး အသစ်ပြန်ဖွင့်မည်",
    },
    "btn_back_home": {
        "en": "⬅️ Back to Menu",
        "my": "⬅️ မီနူးသို့ ပြန်သွားမည်",
    },
    "btn_admin_users": {
        "en": "👥 Admin: Users & Registry",
        "my": "👥 Admin: အသုံးပြုသူများနှင့် စာရင်း",
    },
    "btn_send_now_short": {
        "en": "📨 Send Now",
        "my": "📨 ယခုပို့မည်",
    },
    "btn_delete_short": {
        "en": "🗑️ Delete",
        "my": "🗑️ ပယ်ဖျက်မည်",
    },
    "btn_clear_history": {
        "en": "🗑️ Clear Sent History",
        "my": "🗑️ ပို့ပြီးမှတ်တမ်း ရှင်းလင်းမည်",
    },
    "btn_refresh_queue": {
        "en": "🔄 Refresh Queue",
        "my": "🔄 စာရင်း အသစ်ပြန်ဖွင့်မည်",
    },
    "btn_prev": {
        "en": "⬅️ Prev",
        "my": "⬅️ ရှေ့သို့",
    },
    "btn_next": {
        "en": "Next ➡️",
        "my": "နောက်သို့ ➡️",
    },
    "btn_confirm_delete": {
        "en": "🗑️ Yes, Delete",
        "my": "🗑️ ဟုတ်ကဲ့၊ ပယ်ဖျက်မည်",
    },
    "btn_cancel": {
        "en": "❌ Cancel",
        "my": "❌ မလုပ်တော့ပါ",
    },

    # Welcome / Auth
    "student_registered": {
        "en": (
            "👋 Hello, {name}!\n\n"
            "✅ *You are now registered!*\n"
            "You will receive automatic payment reminders when your "
            "teacher schedules them.\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "🎓 *Are you a teacher/admin?*\n"
            "Enter the *Admin Chat ID* below to get full access.\n"
            "━━━━━━━━━━━━━━━━━━━━"
        ),
        "my": (
            "👋 မင်္ဂလာပါ {name}!\n\n"
            "✅ *မှတ်ပုံတင်ခြင်း အောင်မြင်ပါသည်!*\n"
            "ဆရာ/ဆရာမများ သင်တန်းကြေး သတိပေးချက် ပို့သည့်အခါ အလိုအလျောက် ရရှိမည်ဖြစ်ပါသည်။\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "🎓 *သင်သည် ဆရာ/ဆရာမ (သို့မဟုတ်) Admin ဖြစ်ပါသလား?*\n"
            "အသုံးပြုခွင့်ရရန် *Admin Chat ID* ကို အောက်တွင် ရိုက်ထည့်ပေးပါ။\n"
            "━━━━━━━━━━━━━━━━━━━━"
        ),
    },
    "auth_success": {
        "en": (
            "🎉 *Authorization Successful!*\n\n"
            "You are now registered as an authorized teacher.\n"
            "Send an Excel (`.xlsx`) file anytime to load your students.\n\n"
            "Use /menu to open your interactive dashboard."
        ),
        "my": (
            "🎉 *အတည်ပြုခြင်း အောင်မြင်ပါသည်!*\n\n"
            "ဆရာ/ဆရာမအဖြစ် အသုံးပြုခွင့် ရရှိပါပြီ။\n"
            "ကျောင်းသားစာရင်း ထည့်သွင်းရန် Excel (`.xlsx`) ဖိုင်ကို အချိန်မရွေး ပို့နိုင်ပါသည်။\n\n"
            "ထိန်းချုပ်ခလုတ်များကို ကြည့်ရှုရန် /menu ကို နှိပ်ပါ။"
        ),
    },
    "auth_failed": {
        "en": (
            "❌ *Incorrect Admin Chat ID.*\n\n"
            "Access was not granted. Please contact the bot administrator."
        ),
        "my": (
            "❌ *Admin Chat ID မှားယွင်းနေပါသည်။*\n\n"
            "အသုံးပြုခွင့် မရရှိသေးပါ။ Administrator ထံ ဆက်သွယ်ပါ။"
        ),
    },
    "not_authorized": {
        "en": "❌ You are not authorized. Send /start to begin.",
        "my": "❌ အသုံးပြုခွင့် မရှိသေးပါ။ စတင်ရန် /start ကို နှိပ်ပါ။",
    },

    # Home / Dashboard
    "home_dashboard": {
        "en": (
            "👋 *Welcome, {name}!*\n\n"
            "🎯 *Payment Reminder Control Panel*\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "📊 *Students Loaded:* `{record_count}`\n"
            "⏱️ *Bot Time:* `{now_str}` (`{timezone}`)\n"
            "⚙️ *Next Timer Check:* ⏳ `in {countdown_str}`\n"
            "🎯 *Reminder Offset:* `{reminder_days} day(s)`\n"
            "🌐 *Language:* {lang_label}\n"
            "━━━━━━━━━━━━━━━━━━━━\n\n"
            "📤 *Tip:* Send any `.xlsx` file anytime to update your student roster.\n"
            "Choose an action below:"
        ),
        "my": (
            "👋 *မင်္ဂလာပါ {name}!*\n\n"
            "🎯 *သင်တန်းကြေး သတိပေးချက် ထိန်းချုပ်ခလုတ်များ*\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "📊 *ကျောင်းသားစာရင်း:* `{record_count}` ယောက်\n"
            "⏱️ *လက်ရှိအချိန်:* `{now_str}` (`{timezone}`)\n"
            "⚙️ *နောက်တစ်ကြိမ် စစ်ဆေးရန်ကျန်ချိန်:* ⏳ `{countdown_str}` အတွင်း\n"
            "🎯 *ကြိုတင်အသိပေးမည့်ရက်:* `{reminder_days}` ရက်\n"
            "🌐 *လက်ရှိဘာသာစကား:* {lang_label}\n"
            "━━━━━━━━━━━━━━━━━━━━\n\n"
            "📤 *အကြံပြုချက်:* ကျောင်းသားစာရင်း ပြင်ဆင်ရန် `.xlsx` ဖိုင်ကို အချိန်မရွေး ပို့နိုင်ပါသည်။\n"
            "လုပ်ဆောင်လိုသည်ကို ရွေးချယ်ပါ:"
        ),
    },

    # Language Menu
    "lang_menu_text": {
        "en": (
            "🌐 *Language Settings / ဘာသာစကား ရွေးချယ်မှု*\n\n"
            "• Current Language: *{cur_lang}*\n\n"
            "Please select your preferred language for the bot interface and reminders:"
        ),
        "my": (
            "🌐 *ဘာသာစကား ရွေးချယ်မှု / Language Settings*\n\n"
            "• လက်ရှိရွေးချယ်ထားသော ဘာသာစကား: *{cur_lang}*\n\n"
            "ဘော့တ်နှင့် သတိပေးချက်များတွင် အသုံးပြုလိုသော ဘာသာစကားကို ရွေးချယ်ပါ:"
        ),
    },
    "lang_changed": {
        "en": "✅ Language set to English!",
        "my": "✅ မြန်မာဘာသာသို့ ပြောင်းလဲပြီးပါပြီ!",
    },

    # Status
    "status_title": {
        "en": "📊 *Your Roster & Status*",
        "my": "📊 *ကျောင်းသားစာရင်းနှင့် အခြေအနေ*",
    },
    "status_total": {
        "en": "• Total Loaded: `{total}` student(s)",
        "my": "• စာရင်းသွင်းထားသော စုစုပေါင်း: `{total}` ယောက်",
    },
    "status_linked": {
        "en": "• ✅ Chat ID Linked: `{with_chat_id}`",
        "my": "• ✅ Chat ID ချိတ်ဆက်ပြီး: `{with_chat_id}` ယောက်",
    },
    "status_missing": {
        "en": "• ⚠️ Missing Chat ID: `{without_chat_id}`",
        "my": "• ⚠️ Chat ID မရှိသေးသူ: `{without_chat_id}` ယောက်",
    },
    "status_time": {
        "en": "• ⏱️ Bot Time: `{now_str}` (`{timezone}`)",
        "my": "• ⏱️ လက်ရှိအချိန်: `{now_str}` (`{timezone}`)",
    },
    "status_offset": {
        "en": "• 🎯 Target Offset: `REMINDER_DAYS_BEFORE = {reminder_days}` day(s)",
        "my": "• 🎯 ကြိုတင်အသိပေးရက်: `{reminder_days}` ရက်အလို",
    },
    "status_empty": {
        "en": "\n_No Excel file loaded yet. Upload a .xlsx file to begin._",
        "my": "\n_Excel ဖိုင် မတင်ရသေးပါ။ စတင်ရန် .xlsx ဖိုင် ပို့ပေးပါ။_",
    },
    "status_upcoming_header": {
        "en": "\n📅 *Next upcoming reminders:*",
        "my": "\n📅 *လာမည့် သတိပေးချက်များ:*",
    },
    "status_no_upcoming": {
        "en": "\n_No upcoming reminders in the next period._",
        "my": "\n_လောလောဆယ် လာမည့် သတိပေးချက် မရှိသေးပါ။_",
    },

    # Scheduler
    "scheduler_title": {
        "en": "⚙️ *Scheduler Engine*",
        "my": "⚙️ *အချိန်ဇယား စနစ်အခြေအနေ*",
    },
    "scheduler_status": {
        "en": "• Status: {status}",
        "my": "• အခြေအနေ: {status}",
    },
    "scheduler_mode": {
        "en": "• Schedule Mode: {mode}",
        "my": "• အချိန်ဇယား ပုံစံ: {mode}",
    },
    "scheduler_next_check": {
        "en": "• Next Check: `{next_run}`",
        "my": "• နောက်တစ်ကြိမ် စစ်ဆေးမည့်အချိန်: `{next_run}`",
    },
    "scheduler_countdown": {
        "en": "• Countdown: ⏳ *in {countdown}*",
        "my": "• ကျန်ရှိချိန်: ⏳ *{countdown} အတွင်း*",
    },
    "scheduler_no_check": {
        "en": "• Next Check: _None scheduled_",
        "my": "• နောက်တစ်ကြိမ် စစ်ဆေးမည့်အချိန်: _မသတ်မှတ်ရသေးပါ_",
    },
    "scheduler_active_data": {
        "en": "\n📋 *Active Schedule Data*",
        "my": "\n📋 *လက်ရှိ အချိန်ဇယား စာရင်း*",
    },
    "scheduler_roster_summary": {
        "en": "• Total Roster: {total} student(s)",
        "my": "• စုစုပေါင်း ကျောင်းသား: {total} ယောက်",
    },
    "scheduler_sent_summary": {
        "en": "• Already Sent: {sent} | Missing Chat ID: {missing}",
        "my": "• ပို့ပြီး: {sent} ယောက် | Chat ID မရှိသေး: {missing} ယောက်",
    },
    "scheduler_queue_header": {
        "en": "\n🎯 *Next Check Queue (due in {reminder_days}d):*",
        "my": "\n🎯 *လာမည့် အကြိမ်တွင် ပို့မည့်စာရင်း ({reminder_days} ရက်အလို):*",
    },
    "scheduler_queue_empty": {
        "en": "  _None due for the next check._",
        "my": "  _လာမည့်အကြိမ်တွင် ပို့ရန်မရှိသေးပါ။_",
    },

    # Check
    "check_running": {
        "en": "⏳ Running reminder check…",
        "my": "⏳ သတိပေးချက်များကို စစ်ဆေးနေပါသည်…",
    },
    "check_report_title": {
        "en": "🔍 *Reminder Check Report*",
        "my": "🔍 *သတိပေးချက် စစ်ဆေးမှု အစီရင်ခံစာ*",
    },

    # Delete
    "delete_title": {
        "en": "🗑️ *Delete a Reminder Schedule*",
        "my": "🗑️ *သတိပေးချက် အချိန်ဇယား ပယ်ဖျက်ခြင်း*",
    },
    "delete_showing": {
        "en": "Showing {start}–{end} of {total} student(s).\nTap a student button below to delete their reminder:\n",
        "my": "ကျောင်းသား {total} ယောက်အနက် {start} မှ {end} အထိ ဖော်ပြထားပါသည်။\nပယ်ဖျက်လိုသော ကျောင်းသားခလုတ်ကို နှိပ်ပါ:\n",
    },
    "delete_empty": {
        "en": (
            "🗑️ *Delete Reminder Schedule*\n\n"
            "_No active student reminders found in your roster._\n"
            "Upload an Excel (.xlsx) file to add student reminders."
        ),
        "my": (
            "🗑️ *သတိပေးချက် အချိန်ဇယား ပယ်ဖျက်ခြင်း*\n\n"
            "_လက်ရှိ ကျောင်းသားစာရင်း မရှိသေးပါ။_\n"
            "ကျောင်းသားများ ထည့်သွင်းရန် Excel (.xlsx) ဖိုင် ပေးပို့ပါ။"
        ),
    },
    "delete_confirm_prompt": {
        "en": (
            "⚠️ *Confirm Schedule Deletion*\n\n"
            "Are you sure you want to delete the reminder for:\n"
            "• Student: *{student_name}* (`{username}`)\n"
            "• Due Date: `{due_date}`\n\n"
            "This will remove the student from your scheduled reminders and update your Excel roster."
        ),
        "my": (
            "⚠️ *ပယ်ဖျက်ခြင်းကို အတည်ပြုပါ*\n\n"
            "အောက်ပါကျောင်းသား၏ သတိပေးချက်ကို ပယ်ဖျက်ရန် သေချာပါသလား:\n"
            "• ကျောင်းသားအမည်: *{student_name}* (`{username}`)\n"
            "• သတ်မှတ်ရက်: `{due_date}`\n\n"
            "ဤလုပ်ဆောင်ချက်သည် သတိပေးချက်စာရင်းနှင့် Excel ဖိုင်ထဲမှ အပြီးတိုင် ဖယ်ရှားပါမည်။"
        ),
    },
    "delete_success": {
        "en": "🗑️ Deleted schedule for {username} (Due: {due_date}).",
        "my": "🗑️ {username} (သတ်မှတ်ရက်: {due_date}) ၏ သတိပေးချက်ကို ပယ်ဖျက်ပြီးပါပြီ။",
    },

    # History
    "history_title": {
        "en": "📜 *Sent Reminders History*",
        "my": "📜 *ပို့ပြီးသော သတိပေးချက် မှတ်တမ်း*",
    },
    "history_total": {
        "en": "Total reminders recorded: `{count}`",
        "my": "မှတ်တမ်းတင်ထားသော စုစုပေါင်း: `{count}` ခု",
    },
    "history_empty": {
        "en": "\n_No reminders have been recorded as sent yet._",
        "my": "\n_ပေးပို့ထားသော သတိပေးချက် မှတ်တမ်း မရှိသေးပါ။_",
    },
    "history_cleared": {
        "en": "🗑️ *Sent reminders history cleared!* (`{count}` record(s) removed)",
        "my": "🗑️ *ပို့ပြီးမှတ်တမ်းကို ရှင်းလင်းပြီးပါပြီ!* (မှတ်တမ်း `{count}` ခု ဖယ်ရှားပြီး)",
    },

    # Excel Upload
    "excel_saved_title": {
        "en": "📁 *Excel File Imported Successfully!*",
        "my": "📁 *Excel ဖိုင် ထည့်သွင်းခြင်း အောင်မြင်ပါသည်!*",
    },
    "excel_total_records": {
        "en": "• Total records in file: `{total}`",
        "my": "• ဖိုင်ထဲရှိ စုစုပေါင်းစာရင်း: `{total}`",
    },
    "excel_valid_records": {
        "en": "• ✅ Valid records: `{valid}`",
        "my": "• ✅ မှန်ကန်သော စာရင်း: `{valid}`",
    },
    "excel_auto_matched": {
        "en": "• 🔗 Auto-matched Chat IDs: `{matched}`",
        "my": "• 🔗 Chat ID အလိုအလျောက် ချိတ်ဆက်ပြီး: `{matched}`",
    },
    "excel_unregistered": {
        "en": "• ⚠️ Unregistered students (no Chat ID): `{unregistered}`",
        "my": "• ⚠️ Chat ID မရှိသေးသော ကျောင်းသား: `{unregistered}`",
    },
}


def t(key: str, lang: str = LANG_MY, **kwargs: Any) -> str:
    """
    Get translated string for the given key and language.
    Falls back to English if the key or language is not found.
    """
    normalized_lang = LANG_EN if (lang or "").lower().startswith("en") else LANG_MY
    msg_dict = MESSAGES.get(key)
    if not msg_dict:
        return key

    text = msg_dict.get(normalized_lang) or msg_dict.get(LANG_EN, key)
    if kwargs:
        try:
            return text.format(**kwargs)
        except (KeyError, IndexError, ValueError):
            return text
    return text
