# src/dialer.py
# BPI BL SL Automation — Dialer Report Processing
# Source: DIALER REPORT BPI BUSINESS LOAN SL *.xlsx [1]
# Sheet:  Overall Combined Summary
# Extracts 19 columns: DATE to CALL DROP RATE (excludes CYCLE)

import pandas as pd
import logging

logger = logging.getLogger("BPI_BL_SL")

DIALER_SHEET = "Overall Combined Summary"

DIALER_COLUMNS = [
    "DATE", "CLIENT", "ACCOUNTS", "TOTAL DIALED",
    "PENETRATION RATE", "CONNECTED NU", "CONNECTED UNIQUE",
    "TOTAL RPC", "PTP", "CONNECTED % NU", "CONNECTED % UNIQUE",
    "RPC %", "PTP %", "TOTAL TALK TIME", "TALK TIME AVE",
    "TOTAL BALANCE", "NEG DROP", "SYSTEM DROP", "CALL DROP RATE"
]


def extract_dialer_data(file) -> pd.DataFrame:
    """
    Reads the Dialer Report Excel file [1].
    Finds the 'Overall Combined Summary' sheet.
    Returns a DataFrame with only the 19 required columns (excludes CYCLE).
    """
    logger.info("Reading Dialer Report...")

    try:
        df = pd.read_excel(file, sheet_name=DIALER_SHEET, header=None)
    except Exception as e:
        logger.error(f"Failed to read Dialer Report: {e}")
        raise ValueError(f"Could not read sheet '{DIALER_SHEET}': {e}")

    # Find the header row dynamically
    header_row_idx = None
    for idx, row in df.iterrows():
        row_values = [str(v).strip().upper() for v in row.values]
        if "CYCLE" in row_values and "DATE" in row_values:
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError("Could not find header row in Dialer Report.")

    # Re-read with correct header
    df = pd.read_excel(file, sheet_name=DIALER_SHEET, header=header_row_idx)
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Drop CYCLE column, keep only the 19 required columns
    missing = [c for c in DIALER_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Dialer Report: {missing}")

    df = df[DIALER_COLUMNS].dropna(how="all")

    logger.info(f"Dialer Report extracted: {len(df)} row(s) found.")
    return df