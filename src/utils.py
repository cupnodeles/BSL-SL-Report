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


# Display format for date cells written by the automation
# (m/d/yyyy shows 9/1/2026 style; Excel pads to 01/09/2026 with mm/dd/yyyy).
EXCEL_DATE_FMT = "m/d/yyyy"

# Display format for amount cells — real numbers, no green flag.
EXCEL_AMOUNT_FMT = "#,##0.00"

# Fallback display formats for the Penetration sheet, used only when the
# history row gives no usable format. They mimic the existing history
# style (yyyy-mm-dd dates, 0.00% percents, #,##0 balances).
PEN_DATE_FMT = "yyyy-mm-dd"
PEN_PCT_FMT = "0.00%"
PEN_INT_FMT = "#,##0"


def parse_date_value(value):
    """
    Parses a DATE-column value into a datetime.date for Excel output.
    Real Excel dates (not text) are required for pivot tables to
    recognize/group dates.

    - datetime -> date part; date -> as-is.
    - "2026-01-15", "01/15/2026", "2026-01-15 00:00:00" -> date.
    - Blank/NaN/unparseable -> None.
    """
    if value is None:
        return None
    try:
        import pandas as pd  # local import to avoid hard dep at import time
        if value is pd.NaT or (isinstance(value, float) and pd.isna(value)):
            return None
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if s in ("", "nan", "None", "NaT", "NaN", "nat"):
        return None
    try:
        import pandas as pd
        parsed = pd.to_datetime(s, errors="coerce")
        if parsed is None or pd.isna(parsed):
            return None
    except Exception:
        return None
    if isinstance(parsed, datetime):
        return parsed.date()
    try:
        return parsed.date()
    except Exception:
        return None


def parse_amount_value(value):
    """
    Parses an AMOUNT-column value into a float for Excel output.
    Real numbers (not text) keep pivot tables aggregating and avoid
    the green "number stored as text" flag.

    - Strips commas, spaces and currency symbols (₱, PHP, $).
    - "(1,234.56)" parentheses count as negative.
    - Blank/NaN/unparseable -> None (cell left blank, never crashes).
    NOTE: blank is NOT zero — callers decide how to treat None.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            import math
            if isinstance(value, float) and (
                math.isnan(value) or math.isinf(value)
            ):
                return None
        except Exception:
            pass
        return float(value)
    if isinstance(value, datetime):
        return None
    s = str(value).strip()
    if s in ("", "nan", "None", "NaT", "NaN", "nat", "-", "--", "N/A", "n/a"):
        return None
    try:
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1].strip()
        # Drop currency symbols/codes and spaces, keep digits . , -
        s = re.sub(r"(?i)^php\s*", "", s)     # PHP prefix
        s = re.sub(r"^[Pp](?=[\s\d])", "", s)  # P prefix (P 1,200 / P1,200)
        s = re.sub(r"[₱$\s]", "", s)
        s = s.replace(",", "")
        if s in ("", "-", ".", "-."):
            return None
        num = float(s)
        return -num if neg else num
    except Exception:
        return None


def parse_percent_value(value):
    """
    Parses a PERCENT-column value into a fraction for Excel output
    (0.0278 with a % format displays as 2.78%).

    - Literal "9.22%" -> 0.0922 (divided by 100, standard semantics).
    - Plain 9.2177 / "0.0278" -> kept as-is (dialer stores fractions).
    - Blank/NaN/unparseable -> None (cell left blank, never crashes).
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            import math
            if isinstance(value, float) and (
                math.isnan(value) or math.isinf(value)
            ):
                return None
        except Exception:
            pass
        return float(value)
    s = str(value).strip()
    if s in ("", "nan", "None", "NaT", "NaN", "nat", "-", "--", "N/A", "n/a"):
        return None
    try:
        has_pct = "%" in s
        s = s.replace("%", "")
        amt = parse_amount_value(s)
        if amt is None:
            return None
        return amt / 100.0 if has_pct else amt
    except Exception:
        return None


def _looks_like_date_text(s: str) -> bool:
    """True for ISO '2026-09-01' / '2026-09-01 00:00:00' style strings."""
    s = s.strip()
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}(\s+\d{1,2}:\d{2}(:\d{2})?(\.0+)?)?$", s):
        return True
    if re.match(
        r"^\d{1,2}/\d{1,2}/\d{4}(\s+\d{1,2}:\d{2}(:\d{2})?(\.0+)?)?$", s
    ):
        return True
    return False


def infer_cell_kind(cell) -> str:
    """
    Probes a history-row cell to decide how the new row's value in the
    same column should be typed ("follow the format").

    Returns one of: "percent" | "date" | "number" |
    "percent-history" | "date-history" | "number-history" | "text".
    - "percent"/"date"/"number": history cell is already typed — reuse
      its number format as-is.
    - "*-history": history cell is TEXT that looks like a percent, date
      or grouped number — callers should write a typed value with an
      explicit fallback format that looks identical (PEN_*_FMT).
    - "text": durations, names and anything else — write text unchanged.
    """
    try:
        v = cell.value
    except Exception:
        return "text"
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return "text"
    if isinstance(v, bool):
        return "text"
    if isinstance(v, datetime):
        return "date"
    if isinstance(v, date):
        return "date"
    if isinstance(v, (int, float)):
        try:
            fmt = str(getattr(cell, "number_format", "") or "")
        except Exception:
            fmt = ""
        if "%" in fmt:
            return "percent"
        if _is_date_format(fmt):
            return "date"
        return "number"
    if isinstance(v, str):
        s = v.strip()
        try:
            fmt = str(getattr(cell, "number_format", "") or "")
        except Exception:
            fmt = ""
        if "%" in fmt or (
            s.endswith("%")
            and parse_amount_value(s.replace("%", "")) is not None
        ):
            return "percent-history" if "%" not in fmt else "percent"
        if _is_date_format(fmt) or _looks_like_date_text(s):
            return "date" if _is_date_format(fmt) else "date-history"
        if "," in s and parse_amount_value(s) is not None:
            return "number-history"
        return "text"
    return "text"


def _is_date_format(fmt: str) -> bool:
    """Heuristic: does an Excel number-format string render dates?"""
    f = (fmt or "").lower()
    if not f or f in ("general", "@"):
        return False
    # Strip quoted literals and escaped chars, then look for date tokens.
    f = re.sub(r'"[^"]*"', "", f)
    f = f.replace("\\", "")
    return any(tok in f for tok in (
        "yyyy", "yy", "mmm", "mm", "dd", "d/m", "m/d",
    ))