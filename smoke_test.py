"""
smoke_test.py
-------------
Offline smoke tests for excel_handler.py and reminder_scheduler.py.
No real Telegram calls are made.

Run with:
    python smoke_test.py
"""

import datetime
import io
import json
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

# Force UTF-8 output on Windows (avoids cp1252 UnicodeEncodeError with emoji)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import openpyxl

# Make sure project root is on path
sys.path.insert(0, str(Path(__file__).parent))

import excel_handler
import reminder_scheduler

PASS = "✅ PASS"
FAIL = "❌ FAIL"
results = []


def check(name: str, condition: bool, detail: str = "") -> None:
    status = PASS if condition else FAIL
    msg = f"  {status}  {name}"
    if detail:
        msg += f"\n         {detail}"
    print(msg)
    results.append(condition)


# ===========================================================================
# Helpers
# ===========================================================================


def make_xlsx(rows: list[list], headers: list[str]) -> str:
    """Create a temporary .xlsx file and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    return tmp


# ===========================================================================
# Test 1: Valid rows parse correctly
# ===========================================================================
print("\n--- Test 1: Valid rows ---")
xlsx = make_xlsx(
    headers=["telegram_username", "student_name", "due_date", "telegram_chat_id"],
    rows=[
        ["@Alice", "Alice Smith", "2026-12-01", 111111111],
        ["@Bob", "Bob Jones", "2026-11-20", 222222222],
    ],
)
records, invalid, errors = excel_handler.load_excel(xlsx)
os.unlink(xlsx)

check("Two valid rows loaded", len(records) == 2, f"got {len(records)}")
check("No invalid rows", invalid == 0, f"got {invalid}")
check("due_date is datetime.date", isinstance(records[0]["due_date"], datetime.date))
check("student_name preserved", records[0]["student_name"] == "Alice Smith")
check("chat_id is int", records[0]["telegram_chat_id"] == 111111111)

# ===========================================================================
# Test 2: Invalid rows are counted, valid rows still returned
# ===========================================================================
print("\n--- Test 2: Mixed valid/invalid rows ---")
xlsx = make_xlsx(
    headers=["telegram_username", "student_name", "due_date", "telegram_chat_id"],
    rows=[
        ["@Valid", "Valid Student", "2026-12-10", 333333333],
        ["no_at_sign", "Bad Username", "2026-12-11", 444444444],   # bad username
        ["@Also Valid", "Good", "not-a-date", 555555555],          # bad date
    ],
)
records, invalid, errors = excel_handler.load_excel(xlsx)
os.unlink(xlsx)

check("1 valid row returned", len(records) == 1, f"got {len(records)}")
check("2 invalid rows counted", invalid == 2, f"got {invalid}")
check("Errors list has entries", len(errors) >= 2, f"got {len(errors)}")

# ===========================================================================
# Test 3: Blank rows are silently skipped
# ===========================================================================
print("\n--- Test 3: Blank rows skipped ---")
xlsx = make_xlsx(
    headers=["telegram_username", "student_name", "due_date", "telegram_chat_id"],
    rows=[
        ["@Real", "Real Student", "2026-12-15", 666666666],
        [None, None, None, None],   # blank row
        ["", "", "", ""],           # blank row
    ],
)
records, invalid, errors = excel_handler.load_excel(xlsx)
os.unlink(xlsx)

check("1 valid row (blank rows skipped)", len(records) == 1, f"got {len(records)}")

# ===========================================================================
# Test 4: Case-insensitive column headers
# ===========================================================================
print("\n--- Test 4: Case-insensitive headers ---")
xlsx = make_xlsx(
    headers=["Telegram_Username", "Student_Name", "Due_Date", "Telegram_Chat_Id"],
    rows=[
        ["@CaseTest", "Case Test", "2026-12-20", 777777777],
    ],
)
records, invalid, errors = excel_handler.load_excel(xlsx)
os.unlink(xlsx)

check("Case-insensitive headers work", len(records) == 1, f"got {len(records)}")

# ===========================================================================
# Test 5: Missing telegram_chat_id is OK (not an error)
# ===========================================================================
print("\n--- Test 5: Missing chat_id is not an error ---")
xlsx = make_xlsx(
    headers=["telegram_username", "due_date"],
    rows=[
        ["@NoChatId", "2026-12-25"],
    ],
)
records, invalid, errors = excel_handler.load_excel(xlsx)
os.unlink(xlsx)

check("Row without chat_id column is valid", len(records) == 1, f"got {len(records)}")
check("chat_id is None", records[0]["telegram_chat_id"] is None)
check("student_name falls back to username", records[0]["student_name"] == "@NoChatId")

# ===========================================================================
# Test 6: Date math — reminder triggers on exact day
# ===========================================================================
print("\n--- Test 6: Date math (days_until_due) ---")
today = datetime.date.today()
five_days_later = today + datetime.timedelta(days=5)
ten_days_later = today + datetime.timedelta(days=10)

records_mock = [
    {
        "telegram_username": "@Alice",
        "student_name": "Alice",
        "due_date": five_days_later,
        "telegram_chat_id": 111111111,
    },
    {
        "telegram_username": "@Bob",
        "student_name": "Bob",
        "due_date": ten_days_later,
        "telegram_chat_id": 222222222,
    },
    {
        "telegram_username": "@OldDebt",
        "student_name": "OldDebt",
        "due_date": today - datetime.timedelta(days=1),  # past
        "telegram_chat_id": 333333333,
    },
]

upcoming = reminder_scheduler.get_upcoming_reminders(records_mock, "Asia/Yangon", 5, n=5)
check("get_upcoming_reminders returns future only", len(upcoming) == 2, f"got {len(upcoming)}")
check("Sorted ascending by days_until_due", upcoming[0]["days_until_due"] <= upcoming[1]["days_until_due"])

# ===========================================================================
# Test 7: Duplicate prevention
# ===========================================================================
print("\n--- Test 7: Duplicate prevention (sent_reminders.json) ---")
with tempfile.TemporaryDirectory() as tmpdir:
    json_path = Path(tmpdir) / "sent_reminders.json"

    # Empty at start
    keys = reminder_scheduler.load_sent_reminders(json_path)
    check("Empty set on first load", len(keys) == 0)

    # Add a key and save
    keys.add("111111111:2026-12-01")
    reminder_scheduler.save_sent_reminders(keys, json_path)

    # Reload
    keys2 = reminder_scheduler.load_sent_reminders(json_path)
    check("Key persists after reload", "111111111:2026-12-01" in keys2)
    check("Only 1 key saved", len(keys2) == 1)

# ===========================================================================
# Test 8: daily_check only sends to students due in exactly REMINDER_DAYS days
# ===========================================================================
print("\n--- Test 8: daily_check sends correct reminders (mocked bot) ---")

async def _run_daily_check_test():
    with tempfile.TemporaryDirectory() as tmpdir:
        json_path = Path(tmpdir) / "sent.json"
        today = datetime.date.today()
        records = [
            {  # Should send — exactly 5 days out
                "telegram_username": "@ShouldSend",
                "student_name": "ShouldSend",
                "due_date": today + datetime.timedelta(days=5),
                "telegram_chat_id": 999999999,
            },
            {  # Should NOT send — 10 days out
                "telegram_username": "@TooFar",
                "student_name": "TooFar",
                "due_date": today + datetime.timedelta(days=10),
                "telegram_chat_id": 888888888,
            },
            {  # Should NOT send — already past
                "telegram_username": "@Past",
                "student_name": "Past",
                "due_date": today - datetime.timedelta(days=1),
                "telegram_chat_id": 777777777,
            },
            {  # Should NOT send — no chat_id
                "telegram_username": "@NoChatId",
                "student_name": "NoChatId",
                "due_date": today + datetime.timedelta(days=5),
                "telegram_chat_id": None,
            },
        ]

        mock_bot = AsyncMock()
        mock_bot.send_message = AsyncMock()

        await reminder_scheduler.daily_check(
            bot=mock_bot,
            records=records,
            admin_chat_id=12345,
            timezone_str="Asia/Yangon",
            reminder_days=5,
            sent_reminders_path=json_path,
        )

        # send_message called: 1 reminder + 1 admin warning (no chat_id) = 2 calls
        calls = mock_bot.send_message.call_args_list
        sent_chat_ids = [c.kwargs.get("chat_id") or c.args[0] for c in calls]

        check(
            "Reminder sent to correct student",
            999999999 in sent_chat_ids,
            f"sent_chat_ids={sent_chat_ids}",
        )
        check(
            "Too-far student NOT messaged directly",
            888888888 not in sent_chat_ids,
            f"sent_chat_ids={sent_chat_ids}",
        )
        check(
            "Past student NOT messaged",
            777777777 not in sent_chat_ids,
            f"sent_chat_ids={sent_chat_ids}",
        )

        # Reload json — key should be saved
        saved = reminder_scheduler.load_sent_reminders(json_path)
        key = f"999999999:{today + datetime.timedelta(days=5)}"
        check("Sent key persisted to JSON", key in saved, f"saved={saved}")

        # Run again — should NOT resend
        mock_bot.send_message.reset_mock()
        await reminder_scheduler.daily_check(
            bot=mock_bot,
            records=records,
            admin_chat_id=12345,
            timezone_str="Asia/Yangon",
            reminder_days=5,
            sent_reminders_path=json_path,
        )
        second_calls = mock_bot.send_message.call_args_list
        second_student_calls = [
            c for c in second_calls if (c.kwargs.get("chat_id") or c.args[0]) == 999999999
        ]
        check("Duplicate NOT re-sent on second run", len(second_student_calls) == 0)


import asyncio
asyncio.run(_run_daily_check_test())

# ===========================================================================
# Summary
# ===========================================================================
print("\n" + "=" * 50)
passed = sum(results)
total = len(results)
print(f"Smoke tests: {passed}/{total} passed")
if passed == total:
    print("🎉 All tests passed!")
    sys.exit(0)
else:
    print(f"⚠️  {total - passed} test(s) failed.")
    sys.exit(1)
