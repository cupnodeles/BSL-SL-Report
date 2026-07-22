# src/drr.py
# BSL SL Automation — Daily Remark Report Cleaning & Mapping
# Source: Daily_Remark_Report*.xlsx [3]

import pandas as pd
import numpy as np
import logging

logger = logging.getLogger("BSL_SL")

# Remark By values to delete [3]
REMARK_BY_DELETE = ["TLEYVA", "ZMONTANO"]

# Remark starting strings to delete [3]
# NOTE: No period at end of "Updates when case reassign to another collector"
# The actual data does not have a period [3]
REMARK_STARTS_DELETE = [
    "Sub Special Status",
    "System Auto Update",
    "Updates when case reassign to another collector"
]

# Remedial batch prefix [3]
REMEDIAL_PREFIX = "BPI BUSINESS LOAN REMEDIAL"

# Column mapping DRR -> Remedial & Early SL Template [4]
COLUMN_MAP = {
    "Date":              "Date",
    "Time":              "Time",
    "Debtor":            "Name",
    "Account No.":       "LAN",
    "Status":            "Status",
    "Remark":            "Remark",
    "Remark By":         "Remark By",
    "PTP Amount":        "PTP Amount",
    "PTP Date":          "PTP Date",
    "Claim Paid Amount": "Payment Amount",
    "Claim Paid Date":   "Payment Date",
    "Dialed Number":     "Dialed Number",
}

# Extra column for Early SL only
EARLY_EXTRA = {"Cycle": "Bucket"}


def _safe_str(series: pd.Series) -> pd.Series:
    """
    Converts a Series to string safely.
    Replaces NaN, None, 'nan', 'None', 'NaT' with empty string.
    """
    return (
        series
        .fillna("")
        .astype(str)
        .replace({"nan": "", "None": "", "NaT": "", "NaN": ""})
        .str.strip()
    )


def clean_drr(file) -> pd.DataFrame:
    """
    Reads and cleans the Daily Remark Report [3].
    - Deletes rows where Remark By is TLEYVA or ZMONTANO
    - Deletes rows where Remark starts with defined strings
    Returns cleaned DataFrame with NO blank rows.
    """
    logger.info("Reading Daily Remark Report...")

    try:
        df = pd.read_excel(file, dtype=str)
    except Exception as e:
        logger.error(f"Failed to read DRR: {e}")
        raise ValueError(f"Could not read DRR file: {e}")

    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Step 0: Aggressively fill ALL columns with empty string
    df = df.fillna("").astype(str)
    for col in df.columns:
        df[col] = _safe_str(df[col])

    initial_count = len(df)
    logger.info(f"DRR loaded: {initial_count} rows.")

    # Step 1: Delete rows by Remark By [3]
    remark_by_upper = df["Remark By"].str.upper()
    remark_by_mask  = remark_by_upper.isin(
        [r.upper() for r in REMARK_BY_DELETE]
    )
    remark_by_mask  = remark_by_mask.fillna(False).astype(bool)
    df = df[~remark_by_mask].copy()
    logger.info(f"Removed {initial_count - len(df)} rows by Remark By filter.")

    # Step 2: Delete rows by Remark starting strings [3]
    # Uses .startswith() without period to match both:
    # "Updates when case reassign to another collector"
    # "Updates when case reassign to another collector."
    count_before = len(df)

    def starts_with_any(remark: str) -> bool:
        remark = str(remark).strip()
        return any(remark.startswith(s) for s in REMARK_STARTS_DELETE)

    remark_mask = df["Remark"].apply(starts_with_any)
    remark_mask = remark_mask.fillna(False).astype(bool)
    df = df[~remark_mask].copy()
    logger.info(f"Removed {count_before - len(df)} rows by Remark filter.")

    # Step 3: Remove completely blank rows
    # This fixes the Ctrl+Shift+Down stopping early issue [5]
    count_before = len(df)
    df = df[df.apply(
        lambda row: any(str(v).strip() != "" for v in row), axis=1
    )].copy()
    logger.info(f"Removed {count_before - len(df)} blank rows.")
    logger.info(f"DRR cleaned: {len(df)} rows remaining.")

    return df.reset_index(drop=True)


def split_drr(df: pd.DataFrame):
    """
    Splits cleaned DRR into Remedial and Early SL DataFrames.
    Remedial : Batch No starts with 'BSL SL REMEDIAL' [3]
    Early SL : Everything else e.g. BSL SL EARLY SL_01/07/2026 [3]
    Returns (remedial_df, early_df)
    """
    df["Batch No"] = _safe_str(df["Batch No"])

    remedial_mask = df["Batch No"].str.upper().str.startswith(
        REMEDIAL_PREFIX.upper()
    )
    remedial_mask = remedial_mask.fillna(False).astype(bool)

    remedial_df = df[remedial_mask].copy()
    early_df    = df[~remedial_mask].copy()

    logger.info(
        f"Remedial rows: {len(remedial_df)} | Early SL rows: {len(early_df)}"
    )
    return remedial_df, early_df


def map_remedial(df: pd.DataFrame) -> pd.DataFrame:
    """
    Maps Remedial DRR columns to Productivity Template columns [4].
    Returns mapped DataFrame with only the required columns.
    """
    mapped = pd.DataFrame()
    for drr_col, tmpl_col in COLUMN_MAP.items():
        if drr_col in df.columns:
            mapped[tmpl_col] = _safe_str(df[drr_col]).values
        else:
            logger.warning(
                f"Column '{drr_col}' not found in DRR. Filling with empty."
            )
            mapped[tmpl_col] = ""
    return mapped.reset_index(drop=True)


def map_early(df: pd.DataFrame) -> pd.DataFrame:
    """
    Maps Early SL DRR columns to Productivity Template columns [4].
    Includes extra Bucket = Cycle column (Early SL only).
    Returns mapped DataFrame with all required columns.
    """
    mapped = map_remedial(df)

    for drr_col, tmpl_col in EARLY_EXTRA.items():
        if drr_col in df.columns:
            mapped[tmpl_col] = _safe_str(df[drr_col]).values
        else:
            logger.warning(
                f"Column '{drr_col}' not found in DRR. Filling with empty."
            )
            mapped[tmpl_col] = ""

    return mapped.reset_index(drop=True)