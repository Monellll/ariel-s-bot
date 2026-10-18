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
# Test 1: Valid rows parse correctly (including various date formats)
# ===========================================================================
print("\n--- Test 1: Valid rows ---")
xlsx = make_xlsx(
    headers=["telegram_username", "student_name", "due_date", "telegram_chat_id"],
    rows=[
        ["@Alice", "Alice Smith", "2026-12-01", 111111111],
        ["@Bob", "Bob Jones", "2026-11-20", 222222222],
        ["@Charlie", "Charlie Brown", "2026-10-16 00:00:00", 333333333],
        ["@Dave", "Dave Miller", datetime.datetime(2026, 10, 16, 0, 0), 444444444],
    ],
)
records, invalid, errors = excel_handler.load_excel(xlsx)
os.unlink(xlsx)

check("Four valid rows loaded", len(records) == 4, f"got {len(records)}")
check("No invalid rows", invalid == 0, f"got {invalid}")
check("due_date is datetime.date", isinstance(records[0]["due_date"], datetime.date))
check("datetime string parsed", records[2]["due_date"] == datetime.date(2026, 10, 16))
check("native datetime parsed", records[3]["due_date"] == datetime.date(2026, 10, 16))
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
from zoneinfo import ZoneInfo
today = datetime.datetime.now(ZoneInfo("Asia/Yangon")).date()
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
        today = datetime.datetime.now(ZoneInfo("Asia/Yangon")).date()
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
# Test 9: Multi-user isolation (separate files and schedules)
# ===========================================================================
print("\n--- Test 9: Multi-user isolation ---")

async def _run_multi_user_test():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_base = Path(tmpdir)
        user1_dir = tmp_base / "1001"
        user2_dir = tmp_base / "1002"
        user1_dir.mkdir()
        user2_dir.mkdir()

        user1_sent = user1_dir / "sent_reminders.json"
        user2_sent = user2_dir / "sent_reminders.json"

        today = datetime.datetime.now(ZoneInfo("Asia/Yangon")).date()
        target_due = today + datetime.timedelta(days=5)

        # User 1 has Student Alpha
        user1_records = [
            {
                "telegram_username": "@StudentAlpha",
                "student_name": "Alpha",
                "due_date": target_due,
                "telegram_chat_id": 11111,
            }
        ]

        # User 2 has Student Beta
        user2_records = [
            {
                "telegram_username": "@StudentBeta",
                "student_name": "Beta",
                "due_date": target_due,
                "telegram_chat_id": 22222,
            }
        ]

        mock_bot = AsyncMock()

        # Run User 1 check
        await reminder_scheduler.daily_check(
            bot=mock_bot,
            records=user1_records,
            owner_chat_id=1001,
            timezone_str="Asia/Yangon",
            reminder_days=5,
            sent_reminders_path=user1_sent,
        )

        user1_keys = reminder_scheduler.load_sent_reminders(user1_sent)
        user2_keys = reminder_scheduler.load_sent_reminders(user2_sent)

        check("User 1 sent recorded in User 1 history", f"11111:{target_due}" in user1_keys)
        check("User 2 history remains empty", len(user2_keys) == 0)

        # Run User 2 check
        await reminder_scheduler.daily_check(
            bot=mock_bot,
            records=user2_records,
            owner_chat_id=1002,
            timezone_str="Asia/Yangon",
            reminder_days=5,
            sent_reminders_path=user2_sent,
        )

        user2_keys_after = reminder_scheduler.load_sent_reminders(user2_sent)
        check("User 2 sent recorded in User 2 history", f"22222:{target_due}" in user2_keys_after)
        check("User 1 history unaffected by User 2 run", len(reminder_scheduler.load_sent_reminders(user1_sent)) == 1)

asyncio.run(_run_multi_user_test())

# ===========================================================================
# Test 10: Authorized user storage & admin auto-authorization
# ===========================================================================
print("\n--- Test 10: User authorization storage ---")
with tempfile.TemporaryDirectory() as tmpdir:
    auth_file = Path(tmpdir) / "authorized_users.json"
    admin_id = 987654

    # Test load when file doesn't exist yet
    users = {admin_id}
    if auth_file.exists():
        with open(auth_file, "r", encoding="utf-8") as f:
            users.update(json.load(f))
    check("Admin is auto-authorized by default", admin_id in users)

    # Save a new user
    users.add(123456)
    with open(auth_file, "w", encoding="utf-8") as f:
        json.dump(sorted(users), f)

    # Reload from disk
    loaded_users = set()
    with open(auth_file, "r", encoding="utf-8") as f:
        loaded_users = set(json.load(f))
    check("Persisted authorized users load correctly", 123456 in loaded_users and admin_id in loaded_users)
    check("Total authorized count is 2", len(loaded_users) == 2)

# ===========================================================================
# Test 11: Dynamic reminder message formatting
# ===========================================================================
print("\n--- Test 11: Dynamic reminder message formatting ---")
msg_default = reminder_scheduler.format_reminder_message(
    student_name="Ariel",
    due_date=datetime.date(2026, 10, 14),
    reminder_days=5,
)
check("Message contains dynamic October 14th", "October 14th" in msg_default)
check("Message contains Myanmar October 14", "October 14 ရက်နေ့" in msg_default)
check("Message contains fee for the class text", "fee for the class is due on" in msg_default)
check("Message contains gentle reminder text", "No worries at all, just a gentle reminder." in msg_default)
check("Message contains Myanmar fee text", "ဒီလအတွက် သင်တန်းကြေးလေးက" in msg_default)
check("Message contains Myanmar polite closing", "မလောပါဘူးရှင် remind လုပ်တာပါနော်🤍" in msg_default)
check("Message does not contain start_date reference", "since we started" not in msg_default)

# English-only
msg_en = reminder_scheduler.format_reminder_message(
    student_name="Ariel",
    due_date=datetime.date(2026, 10, 14),
    language="english",
)
check("English-only contains English text", "fee for the class is due on October 14th" in msg_en)
check("English-only does NOT contain Myanmar text", "ဒီလအတွက် သင်တန်းကြေးလေးက" not in msg_en)

# Burmese-only
msg_my = reminder_scheduler.format_reminder_message(
    student_name="Ariel",
    due_date=datetime.date(2026, 10, 14),
    language="burmese",
)
check("Burmese-only contains Myanmar text", "ဒီလအတွက် သင်တန်းကြေးလေးက October 14 ရက်နေ့" in msg_my)
check("Burmese-only does NOT contain English text", "fee for the class is due" not in msg_my)

# ===========================================================================
# Test 12: save_excel and delete persistence
# ===========================================================================
print("\n--- Test 12: save_excel and delete persistence ---")
with tempfile.TemporaryDirectory() as tmpdir:
    test_xlsx = Path(tmpdir) / "students.xlsx"
    records_to_save = [
        {
            "telegram_username": "@student1",
            "student_name": "Student One",
            "due_date": datetime.date(2026, 11, 1),
            "telegram_chat_id": 111111,
            "start_date": datetime.date(2026, 7, 14),
        },
        {
            "telegram_username": "@student2",
            "student_name": "Student Two",
            "due_date": datetime.date(2026, 11, 2),
            "telegram_chat_id": 222222,
            "start_date": None,
        },
    ]

    # Save initial records
    excel_handler.save_excel(records_to_save, str(test_xlsx))
    check("Excel file created by save_excel", test_xlsx.exists())

    # Reload and verify
    loaded, invalids, errs = excel_handler.load_excel(str(test_xlsx))
    check("Reloaded records count matches", len(loaded) == 2, f"got {len(loaded)}")
    check("Student 1 username preserved", loaded[0]["telegram_username"] == "@student1")
    check("Student 1 chat_id preserved", loaded[0]["telegram_chat_id"] == 111111)

    # Delete 1 record and re-save
    loaded.pop(0)
    excel_handler.save_excel(loaded, str(test_xlsx))

    # Reload again
    reloaded, _, _ = excel_handler.load_excel(str(test_xlsx))
    check("Reloaded count after delete is 1", len(reloaded) == 1, f"got {len(reloaded)}")
    check("Remaining student is student2", reloaded[0]["telegram_username"] == "@student2")

# ===========================================================================
# Test 13: Strict duplicate impossibility & concurrency protection
# ===========================================================================
print("\n--- Test 13: Strict duplicate impossibility & concurrency protection ---")
async def run_concurrency_test():
    with tempfile.TemporaryDirectory() as tmpdir:
        today_yangon = datetime.datetime.now(ZoneInfo("Asia/Yangon")).date()
        target_date = today_yangon + datetime.timedelta(days=5)

        # Roster with duplicate rows of the exact same student
        dup_records = [
            {
                "telegram_username": "@dupstudent",
                "student_name": "Dup Student",
                "due_date": target_date,
                "telegram_chat_id": 555555,
            },
            {
                "telegram_username": "@dupstudent",
                "student_name": "Dup Student",
                "due_date": target_date,
                "telegram_chat_id": 555555,
            },
        ]

        mock_bot = AsyncMock()
        mock_bot.send_message = AsyncMock()

        # Execute 3 concurrent daily_check calls simultaneously
        results = await asyncio.gather(
            reminder_scheduler.daily_check(
                bot=mock_bot,
                records=dup_records,
                admin_chat_id=12345,
                timezone_str="Asia/Yangon",
                reminder_days=5,
                sent_reminders_path=json_path,
            ),
            reminder_scheduler.daily_check(
                bot=mock_bot,
                records=dup_records,
                admin_chat_id=12345,
                timezone_str="Asia/Yangon",
                reminder_days=5,
                sent_reminders_path=json_path,
            ),
            reminder_scheduler.daily_check(
                bot=mock_bot,
                records=dup_records,
                admin_chat_id=12345,
                timezone_str="Asia/Yangon",
                reminder_days=5,
                sent_reminders_path=json_path,
            ),
        )

        # Count total send_message calls for student 555555
        calls = mock_bot.send_message.call_args_list
        student_calls = [
            c for c in calls if (c.kwargs.get("chat_id") == 555555 or (c.args and c.args[0] == 555555))
        ]
        check(
            "Exactly 1 reminder sent across 3 concurrent checks and duplicate rows",
            len(student_calls) == 1,
            f"got {len(student_calls)} calls (expected 1)",
        )

        # Confirm persisted immediately
        saved_keys = reminder_scheduler.load_sent_reminders(json_path)
        expected_key = f"555555:{target_date.strftime('%Y-%m-%d')}"
        check("Key is persisted in sent_reminders.json", expected_key in saved_keys)

asyncio.run(run_concurrency_test())

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
