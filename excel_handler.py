"""
excel_handler.py
----------------
Reads and validates the student payment Excel file (.xlsx).
No Telegram dependencies — pure Python + openpyxl.

Returns:
    valid_records  : list of dicts with cleaned student data
    invalid_count  : number of rows that failed validation
    error_messages : list of human-readable error strings for admin
"""

import datetime
import logging
from pathlib import Path
from typing import Optional

import openpyxl

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------

StudentRecord = dict  # keys: telegram_username, student_name, due_date, telegram_chat_id

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def load_excel(file_path: str) -> tuple[list[StudentRecord], int, list[str]]:
    """
    Load and validate an Excel (.xlsx) file of student payment records.

    Args:
        file_path: Path to the .xlsx file.

    Returns:
        (valid_records, invalid_count, error_messages)

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file cannot be opened as a valid xlsx.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    try:
        wb = openpyxl.load_workbook(filename=str(path), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError(f"Cannot open Excel file: {exc}") from exc

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise ValueError("Excel file is empty.")

    # -----------------------------------------------------------------------
    # Parse header row
    # -----------------------------------------------------------------------
    raw_headers = rows[0]
    col_index = _map_columns(raw_headers)

    required = {"telegram_username", "due_date"}
    missing = required - col_index.keys()
    if missing:
        raise ValueError(
            f"Excel file is missing required column(s): {', '.join(sorted(missing))}. "
            f"Found columns: {list(col_index.keys())}"
        )

    # -----------------------------------------------------------------------
    # Parse data rows
    # -----------------------------------------------------------------------
    valid_records: list[StudentRecord] = []
    error_messages: list[str] = []
    invalid_count = 0
    seen_entries: set[tuple[str, str]] = set()

    for row_num, row in enumerate(rows[1:], start=2):
        # Skip entirely blank rows
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        errors, record = _parse_row(row_num, row, col_index)

        if errors:
            invalid_count += 1
            error_messages.extend(errors)
            logger.warning("Row %d skipped: %s", row_num, "; ".join(errors))
        elif record:
            # Duplicate row safeguard: skip identical student + due date entries in same sheet
            uname = record.get("telegram_username", "").lower()
            due_str = str(record.get("due_date"))
            dedup_tuple = (uname, due_str)
            if dedup_tuple in seen_entries:
                logger.warning(
                    "Row %d skipped: duplicate entry for %s with due_date %s",
                    row_num,
                    record.get("telegram_username"),
                    record.get("due_date"),
                )
                continue
            seen_entries.add(dedup_tuple)
            valid_records.append(record)

    logger.info(
        "Excel loaded: %d valid, %d invalid rows", len(valid_records), invalid_count
    )
    return valid_records, invalid_count, error_messages


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _map_columns(header_row: tuple) -> dict[str, int]:
    """
    Build a case-insensitive, whitespace-stripped column name → index map.
    """
    col_map: dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        normalized = str(cell).strip().lower()
        if normalized:
            col_map[normalized] = idx
    return col_map


def _get_cell(row: tuple, col_index: dict[str, int], col_name: str):
    """
    Safely retrieve a cell value by column name.
    Returns None if column doesn't exist or cell is blank.
    Preserves non-string types (e.g. datetime.date, datetime.datetime, int, float)
    so that callers like _parse_date() can use isinstance checks correctly.
    """
    idx = col_index.get(col_name)
    if idx is None or idx >= len(row):
        return None
    val = row[idx]
    if val is None:
        return None
    # Preserve native types (date/datetime/numeric) — do NOT stringify them here,
    # because _parse_date() relies on isinstance to handle them correctly.
    if isinstance(val, (datetime.date, datetime.datetime, int, float)):
        return val
    stripped = str(val).strip()
    return stripped if stripped else None


def _parse_row(
    row_num: int, row: tuple, col_index: dict[str, int]
) -> tuple[list[str], Optional[StudentRecord]]:
    """
    Validate a single data row.

    Returns:
        (errors, record) — errors is empty and record is populated on success.
    """
    errors: list[str] = []

    # --- telegram_username ---
    username_raw = _get_cell(row, col_index, "telegram_username")
    if not username_raw:
        errors.append(f"Row {row_num}: 'telegram_username' is missing.")
        username = None
    elif not username_raw.startswith("@") or " " in username_raw:
        errors.append(
            f"Row {row_num}: 'telegram_username' must start with '@' and contain no spaces. "
            f"Got: {username_raw!r}"
        )
        username = None
    else:
        username = username_raw

    # --- student_name (optional) ---
    name_raw = _get_cell(row, col_index, "student_name")
    student_name = name_raw if name_raw else (username or "Unknown")

    # --- due_date ---
    due_date_raw = _get_cell(row, col_index, "due_date")
    due_date = _parse_date(due_date_raw, row_num, errors)

    # --- telegram_chat_id (optional) ---
    chat_id_raw = _get_cell(row, col_index, "telegram_chat_id")
    telegram_chat_id = _parse_chat_id(chat_id_raw, row_num, errors)

    # --- start_date (optional) ---
    start_date_raw = _get_cell(row, col_index, "start_date") or _get_cell(
        row, col_index, "class_start_date"
    )
    start_date = _parse_date(start_date_raw, row_num, errors) if start_date_raw else None

    if errors:
        return errors, None

    # Warn (but don't reject) for past due dates
    if due_date and due_date < datetime.date.today():
        logger.warning(
            "Row %d: due_date %s is in the past for %s", row_num, due_date, username
        )

    record: StudentRecord = {
        "telegram_username": username,
        "student_name": student_name,
        "due_date": due_date,
        "telegram_chat_id": telegram_chat_id,
        "start_date": start_date,
    }
    return [], record


def _parse_date(
    raw: object, row_num: int, errors: list[str]
) -> Optional[datetime.date]:
    """
    Parse a date value from a cell.
    Accepts:
      - Python datetime.date / datetime.datetime (from openpyxl native date cells)
      - 'YYYY-MM-DD' strings, or datetimes with timestamps (e.g. 'YYYY-MM-DD HH:MM:SS')
      - Common date formats: DD/MM/YYYY, MM/DD/YYYY, YYYY/MM/DD, DD-MM-YYYY
    """
    if raw is None:
        errors.append(f"Row {row_num}: 'due_date' is missing.")
        return None

    # openpyxl may return date/datetime objects directly (data_only=True)
    if isinstance(raw, datetime.datetime):
        return raw.date()
    if isinstance(raw, datetime.date):
        return raw

    s = str(raw).strip()
    if not s:
        errors.append(f"Row {row_num}: 'due_date' is missing.")
        return None

    # Try parsing string formats (both date-only and datetime with timestamp)
    formats = (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %I:%M:%S %p",
        "%Y-%m-%d %I:%M %p",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%m/%d/%Y",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y/%m/%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%d-%m-%Y",
        "%d-%m-%Y %H:%M:%S",
    )
    for fmt in formats:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    # Also try ISO format parsing (handles ISO 8601 timestamps like "2026-10-16T00:00:00Z")
    try:
        return datetime.datetime.fromisoformat(s.replace("Z", "+00:00")).date()
    except (ValueError, TypeError):
        pass

    # If string has space or 'T', try parsing just the date portion
    for sep in (" ", "T"):
        if sep in s:
            date_part = s.split(sep)[0]
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"):
                try:
                    return datetime.datetime.strptime(date_part, fmt).date()
                except ValueError:
                    continue

    errors.append(
        f"Row {row_num}: 'due_date' could not be parsed: {raw!r}. "
        "Expected format: YYYY-MM-DD."
    )
    return None


def _parse_chat_id(
    raw: Optional[str], row_num: int, errors: list[str]
) -> Optional[int]:
    """
    Parse telegram_chat_id from a cell.
    Returns None if the column is missing or the cell is blank (allowed).
    Adds an error if a non-numeric, non-blank value is found.
    """
    if raw is None:
        return None  # Missing is OK — admin will be warned at send time

    # openpyxl may return a float for numeric cells (e.g., 123456789.0)
    if isinstance(raw, (int, float)):
        return int(raw)

    try:
        return int(str(raw).strip())
    except ValueError:
        errors.append(
            f"Row {row_num}: 'telegram_chat_id' must be a number. Got: {raw!r}"
        )
        return None


def save_excel(records: list[StudentRecord], file_path: str) -> None:
    """
    Save student records to an Excel (.xlsx) file.
    """
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Students"

    headers = [
        "telegram_username",
        "student_name",
        "due_date",
        "telegram_chat_id",
        "start_date",
    ]
    ws.append(headers)

    for rec in records:
        due = rec.get("due_date")
        if isinstance(due, (datetime.date, datetime.datetime)):
            due_str = due.strftime("%Y-%m-%d")
        else:
            due_str = str(due) if due is not None else ""

        start = rec.get("start_date")
        if isinstance(start, (datetime.date, datetime.datetime)):
            start_str = start.strftime("%Y-%m-%d")
        else:
            start_str = str(start) if start is not None else ""

        ws.append(
            [
                rec.get("telegram_username", ""),
                rec.get("student_name", ""),
                due_str,
                rec.get("telegram_chat_id") if rec.get("telegram_chat_id") is not None else "",
                start_str,
            ]
        )

    wb.save(str(path))
    wb.close()
    logger.info("Saved %d records to %s", len(records), path)

