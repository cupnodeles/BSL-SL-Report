# src/ptp.py
# BSL SL Automation — PTP Monitoring Report
# Template: SPM PTP Monitoring Report_BSL-early_*.xlsx [2]

import io
import pandas as pd
import openpyxl
import logging
from src.utils import get_last_row

logger = logging.getLogger("BSL_SL")

PTP_STATUS_PREFIX  = "PTP"
SHEET_EARLY        = "Stat Result - BL Early SL"
SHEET_REMEDIAL     = "Stat Result - BL Remedial SL"

# Column mapping: Stat Result -> PTP Monitoring [2]
PTP_COLUMN_MAP = {
    "Date":          "Date",
    "Time":          "Time",
    "Name":          "Name",
    "LAN":           "LAN",
    "Status":        "Status",
    "Remark":        "Remark",
    "Remark By":     "Remark By",
    "PTP Amount":    "PTP Amount",
    "PTP Date":      "PTP Date",
    "Dialed Number": "Dialed Number",
    "Bucket":        "Bucket",
}


def _sheet_to_df(ws) -> pd.DataFrame:
    """
    Converts an openpyxl worksheet to a pandas DataFrame.
    Dynamically finds the header row by looking for 'Date' and 'Status'.
    Handles duplicate column names.
    Fills NaN with empty string.
    """
    data = list(ws.values)
    if not data:
        logger.warning(f"Sheet '{ws.title}' is empty.")
        return pd.DataFrame()

    # Dynamically find header row
    # Look for row that contains 'Date' and 'Status'
    header_idx = 0
    for idx, row in enumerate(data):
        row_vals = [str(v).strip() if v is not None else "" for v in row]
        if "Date" in row_vals and "Status" in row_vals:
            header_idx = idx
            logger.info(f"Found header row at index {idx} in sheet '{ws.title}'")
            break

    raw_headers = [
        str(h).strip() if h is not None else ""
        for h in data[header_idx]
    ]

    # Deduplicate column names
    seen = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    logger.info(f"Sheet '{ws.title}' headers: {headers}")

    rows = data[header_idx + 1:]
    if not rows:
        return pd.DataFrame(columns=headers)

    df = pd.DataFrame(rows, columns=headers)
    df = df.fillna("").astype(str)

    # Strip whitespace from all string columns
    for col in df.columns:
        df[col] = df[col].str.strip()

    return df


def _find_sheet(wb, keywords: list) -> str:
    """
    Dynamically finds a sheet by checking if any keyword
    is contained in the sheet name (case-insensitive).
    Logs all available sheets.
    """
    logger.info(f"Available sheets: {wb.sheetnames}")
    for sheet_name in wb.sheetnames:
        for keyword in keywords:
            if keyword.lower() in sheet_name.lower():
                logger.info(f"Matched sheet '{sheet_name}' with keyword '{keyword}'")
                return sheet_name
    raise ValueError(
        f"Could not find sheet with keywords {keywords}. "
        f"Available sheets: {wb.sheetnames}"
    )


def _find_ptp_data_sheet(wb) -> str:
    """
    Finds the PTP data sheet.
    Priority:
    1. Exact match 'PTP List' [2]
    2. Keyword match containing 'PTP'
    3. Header-based scan for Date + Name + LAN + Status
    4. First sheet fallback
    """
    logger.info(f"Searching PTP data sheet among: {wb.sheetnames}")

    # Priority 1: Exact match 'PTP List' [2]
    if "PTP List" in wb.sheetnames:
        logger.info("Found exact sheet: 'PTP List'")
        return "PTP List"

    # Priority 2: Keyword match
    keywords = ["PTP List", "PTP", "List", "Monitoring", "Data"]
    for sheet_name in wb.sheetnames:
        for keyword in keywords:
            if keyword.lower() in sheet_name.lower():
                logger.info(f"Found PTP sheet by keyword '{keyword}': '{sheet_name}'")
                return sheet_name

    # Priority 3: Header-based scan
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
            row_vals = [
                str(v).strip() if v is not None else ""
                for v in row
            ]
            if (
                "Date"   in row_vals and
                "Name"   in row_vals and
                "LAN"    in row_vals and
                "Status" in row_vals
            ):
                logger.info(f"Found PTP sheet by header scan: '{sheet_name}'")
                return sheet_name

    # Priority 4: Fallback to first sheet
    fallback = wb.sheetnames[0]
    logger.warning(f"Falling back to first sheet: '{fallback}'")
    return fallback


def extract_ptp_rows(prod_bytes: io.BytesIO) -> pd.DataFrame:
    """
    Reads the saved Productivity BytesIO [4].
    Finds Early SL and Remedial SL sheets dynamically.
    Filters rows where Status starts with 'PTP'.
    Returns combined PTP DataFrame mapped to PTP Monitoring columns [2].
    """
    logger.info("Extracting PTP rows from Productivity Report...")

    prod_bytes.seek(0)
    wb = openpyxl.load_workbook(prod_bytes)
    logger.info(f"Productivity sheets: {wb.sheetnames}")

    # Find Early SL sheet [4]
    early_sheet = _find_sheet(wb, [
        "Early SL", "Early", "BL Early", "Stat Result - BL Early"
    ])

    # Find Remedial SL sheet [4]
    remedial_sheet = _find_sheet(wb, [
        "Remedial SL", "Remedial", "BL Remedial", "Stat Result - BL Remedial"
    ])

    early_df    = _sheet_to_df(wb[early_sheet])
    remedial_df = _sheet_to_df(wb[remedial_sheet])

    logger.info(f"Early SL columns: {list(early_df.columns)}")
    logger.info(f"Remedial columns: {list(remedial_df.columns)}")
    logger.info(f"Early SL rows: {len(early_df)} | Remedial rows: {len(remedial_df)}")

    # Add Bucket to Remedial if missing
    if "Bucket" not in remedial_df.columns:
        remedial_df["Bucket"] = ""

    # Add Status to either df if missing (safety)
    if "Status" not in early_df.columns:
        logger.warning("'Status' column not found in Early SL. Adding empty column.")
        early_df["Status"] = ""
    if "Status" not in remedial_df.columns:
        logger.warning("'Status' column not found in Remedial SL. Adding empty column.")
        remedial_df["Status"] = ""

    # Align columns before concat to avoid reindexing error
    all_cols = list(dict.fromkeys(
        list(early_df.columns) + list(remedial_df.columns)
    ))
    early_df    = early_df.reindex(columns=all_cols, fill_value="")
    remedial_df = remedial_df.reindex(columns=all_cols, fill_value="")

    combined = pd.concat([early_df, remedial_df], ignore_index=True)

    logger.info(f"Combined columns: {list(combined.columns)}")
    logger.info(f"Combined rows: {len(combined)}")

    # Safety fill
    combined["Status"] = combined["Status"].fillna("").astype(str).str.strip()

    # Filter Status starts with PTP
    ptp_mask = combined["Status"].str.upper().str.startswith(
        PTP_STATUS_PREFIX.upper()
    )
    ptp_df = combined[ptp_mask.fillna(False)].copy()

    logger.info(f"PTP rows found: {len(ptp_df)}")

    # Map columns to PTP Monitoring [2]
    mapped = pd.DataFrame()
    for src_col, dst_col in PTP_COLUMN_MAP.items():
        if src_col in ptp_df.columns:
            mapped[dst_col] = ptp_df[src_col].values
        else:
            logger.warning(f"Column '{src_col}' not found in Stat Result. Filling empty.")
            mapped[dst_col] = ""

    logger.info(f"Mapped PTP columns: {list(mapped.columns)}")
    return mapped.reset_index(drop=True)


def populate_ptp(template_file, ptp_df: pd.DataFrame) -> io.BytesIO:
    """
    Appends PTP rows to the bottom of 'PTP List' sheet [2].
    Returns the modified workbook as BytesIO.
    """
    logger.info("Loading PTP Monitoring Template...")

    wb = openpyxl.load_workbook(template_file)
    logger.info(f"PTP template sheets: {wb.sheetnames}")

    # Find the PTP data sheet — prioritizes 'PTP List' [2]
    sheet_name = _find_ptp_data_sheet(wb)
    ws = wb[sheet_name]
    logger.info(f"Using PTP sheet: '{sheet_name}'")

    last_row = get_last_row(ws)
    next_row = last_row + 1

    logger.info(f"Pasting {len(ptp_df)} PTP rows at row {next_row}...")

    for r_idx, row in enumerate(ptp_df.itertuples(index=False), start=next_row):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    logger.info("PTP Monitoring Report populated successfully.")
    return output