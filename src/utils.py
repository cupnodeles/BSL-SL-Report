# src/utils.py
# BPI BL SL Automation — Shared Helpers

import os
import logging
from datetime import datetime, timedelta


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
    return logging.getLogger("BPI_BL_SL")