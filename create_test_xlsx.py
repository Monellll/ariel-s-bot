"""Quick script to generate a test Excel file for the bot."""
import openpyxl
from datetime import date
import os

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Students"

ws.append(["telegram_username", "student_name", "due_date"])
ws.append(["@arealtest", "Bao", date(2026, 10, 16)])

# Auto-width
for col in ws.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 2

out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "test_upload.xlsx")
wb.save(out_path)
print(f"Created: {out_path}")
