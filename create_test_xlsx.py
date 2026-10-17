"""
create_test_xlsx.py
-------------------
Helper script to generate a sample Excel file for testing Ariel's Payment Reminder Bot.

Usage:
    python create_test_xlsx.py
    python create_test_xlsx.py @YourTelegramUsername
"""

import os
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import openpyxl

# Bot timezone is Asia/Yangon; in testing mode REMINDER_DAYS_BEFORE=0 sends for today's date
today_yangon = datetime.now(ZoneInfo("Asia/Yangon")).date()
in_5_days = today_yangon + timedelta(days=5)
in_12_days = today_yangon + timedelta(days=12)

# Optional custom username passed as argument, e.g. python create_test_xlsx.py @myusername
my_username = sys.argv[1] if len(sys.argv) > 1 else "@Monellll"
if not my_username.startswith("@"):
    my_username = f"@{my_username}"

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Students"

headers = ["telegram_username", "student_name", "due_date", "start_date"]
ws.append(headers)

# 1. Due today -> triggers immediately in testing mode
ws.append([my_username, "Myat (Self Test)", today_yangon.strftime("%Y-%m-%d"), "2026-07-14"])

# 2. Due today with a custom start date -> tests custom start date formatting
ws.append(["@arealtest", "Bao (Test)", today_yangon.strftime("%Y-%m-%d"), "2026-08-01"])

# 3. Due in 5 days -> appears in upcoming queue
ws.append(["@thiri_test", "Thiri (Upcoming)", in_5_days.strftime("%Y-%m-%d"), "2026-07-14"])

# 4. Due in 12 days -> appears in future schedule
ws.append(["@kyi_test", "Kyi (Future)", in_12_days.strftime("%Y-%m-%d"), ""])

# Format column widths
for col in ws.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 16)

# Save both in root and in telegram-payment-reminder
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
out_path_root = os.path.join(root_dir, "test_upload.xlsx")
out_path_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_upload.xlsx")

wb.save(out_path_root)
wb.save(out_path_local)

# Configure stdout for UTF-8 on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

print("=" * 60)
print("[+] Test Excel file generated successfully!")
print(f"File path: {out_path_root}")
print("------------------------------------------------------------")
print(f"1. {my_username:15} | Due: {today_yangon} (TODAY -> Ready to test!)")
print(f"2. {'@arealtest':15} | Due: {today_yangon} (TODAY -> Custom start Aug 1)")
print(f"3. {'@thiri_test':15} | Due: {in_5_days} (In 5 days -> Upcoming queue)")
print(f"4. {'@kyi_test':15} | Due: {in_12_days} (In 12 days -> Future queue)")
print("=" * 60)
print("Tip: Run 'python create_test_xlsx.py @YourUsername' to put your username in row 1.")

