"""Quick script to generate a test Excel file for the bot."""
import os
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Students"

# Bot timezone is Asia/Yangon; in testing mode REMINDER_DAYS_BEFORE=0 sends for today's date
today_yangon = datetime.now(ZoneInfo("Asia/Yangon")).date()

ws.append(["telegram_username", "student_name", "due_date"])
ws.append(["@arealtest", "Bao", today_yangon.strftime("%Y-%m-%d")])

# Auto-width
for col in ws.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 2

out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "test_upload.xlsx")
wb.save(out_path)
print(f"Created: {out_path} with due_date = {today_yangon}")

