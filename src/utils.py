# src/utils.py
# BSL SL Automation — Shared Helpers

import os
import logging
import re
from datetime import datetime, timedelta, date, time


def get_last_row(sheet) -> int:
    """Returns the index of the last row with data in an openpyxl sheet."""
    last_row = 0
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                last_row = cell.row
    return last_row


def get_yesterday_date() -> datetime:
    """
    Returns yesterday's date.
    If today is Monday, returns Friday (skip weekends).
    """
    today = datetime.today()
    if today.weekday() == 0:       # Monday → return Friday
        return today - timedelta(days=3)
    elif today.weekday() == 6:     # Sunday → return Friday
        return today - timedelta(days=2)
    else:
        return today - timedelta(days=1)


def get_output_filename(prefix: str) -> str:
    """
    Generates output filename with yesterday's date suffix.
    Example: SPM Productivity & Penetration Report_BSL-early&remedial_07212026.xlsx
    """
    date = get_yesterday_date()
    date_str = date.strftime("%m%d%Y")
    return f"{prefix}_{date_str}.xlsx"


def setup_logger(log_dir: str = "logs") -> logging.Logger:
    """
    Sets up a logger that writes to /logs/automation_log_{date}.txt
    Creates /logs/ directory if it doesn't exist.
    """
    os.makedirs(log_dir, exist_ok=True)
    date_str = datetime.today().strftime("%m%d%Y")
    log_file = os.path.join(log_dir, f"automation_log_{date_str}.txt")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("BSL_SL")


# ------------------------------------------------------------------ #
# Date cleaning — strips bogus "00:00:00" midnight times from dates
# ------------------------------------------------------------------ #
_MIDNIGHT_SUFFIX_RE = re.compile(
    r"^(?P<date>\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4})"
    r"\s+00:00:00(?:\.0+)?$"
)
_DATETIME_MIDNIGHT_RE = re.compile(
    r"^(?P<date>.+?)\s+00:00:00(?:\.0+)?$"
)


def strip_midnight_time(value):
    """
    Removes '00:00:00' midnight time from date values.

    - datetime/date       -> date (drops midnight time); datetimes with
                             a real time component keep date part only
                             when used for date columns via
                             clean_date_value(); bare time(0,0) -> "".
    - "YYYY-MM-DD 00:00:00" / "M/D/YYYY 00:00:00" -> date part only.
    - Other strings       -> stripped, unchanged.
    - None / NaN          -> "".
    """
    if value is None:
        return ""
    # pandas NaT / NaN
    try:
        import pandas as pd  # local import to avoid hard dep at import time
        if value is pd.NaT or (isinstance(value, float) and pd.isna(value)):
            return ""
    except Exception:
        pass
    if isinstance(value, datetime):
        if value.time() == time(0, 0, 0):
            return value.date()
        return value
    if isinstance(value, date):
        return value
    if isinstance(value, time):
        if value == time(0, 0, 0):
            return ""
        return value.strftime("%H:%M:%S")
    s = str(value).strip()
    if s in ("", "nan", "None", "NaT", "NaN", "nat"):
        return ""
    m = _MIDNIGHT_SUFFIX_RE.match(s)
    if m:
        return m.group("date")
    m2 = _DATETIME_MIDNIGHT_RE.match(s)
    if m2:
        # Only strip when the trailing time is exactly midnight.
        # e.g. "2026-01-15 00:00:00" -> "2026-01-15"
        # e.g. "01/15/2026 00:00:00" -> "01/15/2026"
        return m2.group("date").strip()
    return s


def clean_date_value(value):
    """
    Normalizes a DATE-column value (Date, PTP Date, Payment Date, DATE):
    - Blank/NaN -> "".
    - datetime with any time -> date part only (drops 00:00:00 AND real
      times — date columns never keep times).
    - "YYYY-MM-DD HH:MM:SS" / "M/D/YYYY HH:MM" strings -> date part only.
    - Bare date strings -> stripped, unchanged.
    """
    if value is None:
        return ""
    try:
        import pandas as pd
        if value is pd.NaT or (isinstance(value, float) and pd.isna(value)):
            return ""
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if s in ("", "nan", "None", "NaT", "NaN", "nat"):
        return ""
    # "2026-01-15 14:30:00" -> "2026-01-15"; "01/15/2026 10:00" -> "01/15/2026"
    if " " in s:
        head, _, tail = s.partition(" ")
        if re.match(r"^\d{1,2}:\d{2}", tail.strip()):
            return head.strip()
    return s