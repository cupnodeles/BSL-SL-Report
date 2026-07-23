# src/ptp.py
# BPI BL SL Automation — PTP Monitoring Report
# Fix: Read Stat Result sheets by column POSITION not by header name
# since data is pasted positionally (A B C D)

import io
import pandas as pd
import openpyxl
import logging
from src.utils import get_last_row

logger = logging.getLogger("BPI_BL_SL")

PTP_STATUS_PREFIX = "PTP"

# Column position mapping for Stat Result sheets [4][5]
# These are the column positions (0-indexed) after positional paste
# Based on confirmed headers: Date, Time, Name, LAN, Status,
# Remark, Remark By, PTP Amount, PTP Date, Payment Amount,
# Payment Date, Dialed Number, Bucket (Early only)
STAT_RESULT_COL_POSITIONS = {
    "Date":           0,
    "Time":           1,
    "Name":           2,
    "LAN":            3,
    "Status":         4,
    "Remark":         5,
    "Remark By":      6,
    "PTP Amount":     7,
    "PTP Date":       8,
    "Payment Amount": 9,
    "Payment Date":   10,
    "Dialed Number":  11,
    "Bucket":         12,
}

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


def _find_sheet(wb, keywords: list) -> str:
    """
    Dynamically finds a sheet by checking if any keyword
    is contained in the sheet name (case-insensitive).
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


def _find_header_row(ws, keys: list) -> int:
    """
    Finds the row index (1-based) containing ALL keys as cell values.
    Returns that row index. Returns 1 if not found.
    """
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True),
        start=1
    ):
        row_vals = [str(v).strip() if v is not None else "" for v in row]
        if all(k in row_vals for k in keys):
            logger.info(f"Header row found at row {row_idx} in '{ws.title}'")
            return row_idx
    logger.warning(f"Header row not found in '{ws.title}'. Using row 1.")
    return 1


def _sheet_to_df_by_header(ws) -> pd.DataFrame:
    """
    Reads a worksheet into DataFrame using the ACTUAL header row.
    Finds the header row dynamically by looking for Date + Status + Name.
    This handles sheets where data starts after row 1 due to pivot tables.
    Deduplicates column names to prevent reindexing errors.
    """
    all_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

    if not all_rows:
        logger.warning(f"Sheet '{ws.title}' is empty.")
        return pd.DataFrame()

    # Find header row
    header_row_idx = 0
    for idx, row in enumerate(all_rows):
        row_vals = [str(v).strip() if v is not None else "" for v in row]
        if "Date" in row_vals and "Status" in row_vals and "Name" in row_vals:
            header_row_idx = idx
            logger.info(
                f"Header row found at index {idx} "
                f"(row {idx + 1}) in '{ws.title}'"
            )
            break

    raw_headers = [
        str(h).strip() if h is not None else ""
        for h in all_rows[header_row_idx]
    ]

    # Deduplicate column names
    seen    = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    logger.info(f"Sheet '{ws.title}' headers: {headers}")

    data_rows = all_rows[header_row_idx + 1:]
    if not data_rows:
        return pd.DataFrame(columns=headers)

    df = pd.DataFrame(data_rows, columns=headers)
    df = df.fillna("").astype(str)
    for col in df.columns:
        df[col] = df[col].str.strip()

    # Remove completely empty rows
    df = df[
        df.apply(lambda row: any(v != "" for v in row), axis=1)
    ].reset_index(drop=True)

    logger.info(f"Sheet '{ws.title}' loaded: {len(df)} data rows.")
    return df


def _find_ptp_data_sheet(wb) -> str:
    """
    Finds the PTP data sheet in the PTP Monitoring template [2].
    Priority:
    1. Exact 'PTP List'
    2. Keyword match containing 'PTP'
    3. Header scan for Date + Name + LAN + Status
    4. First sheet fallback
    """
    logger.info(f"Searching PTP sheet among: {wb.sheetnames}")

    # Priority 1: Exact match
    if "PTP List" in wb.sheetnames:
        logger.info("Found: 'PTP List'")
        return "PTP List"

    # Priority 2: Keyword match
    for sheet_name in wb.sheetnames:
        for keyword in ["PTP List", "PTP", "List", "Monitoring"]:
            if keyword.lower() in sheet_name.lower():
                logger.info(f"Found PTP sheet by keyword: '{sheet_name}'")
                return sheet_name

    # Priority 3: Header scan
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=15, values_only=True):
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
                logger.info(f"Found PTP sheet by header: '{sheet_name}'")
                return sheet_name

    # Priority 4: First sheet
    fallback = wb.sheetnames[0]
    logger.warning(f"Falling back to first sheet: '{fallback}'")
    return fallback


def extract_ptp_rows(prod_bytes: io.BytesIO) -> pd.DataFrame:
    """
    Reads the saved Productivity BytesIO [4][5].
    Finds Early SL and Remedial SL sheets dynamically.
    Reads data using actual header row (not positional).
    Filters rows where Status starts with 'PTP'.
    Returns combined PTP DataFrame mapped to PTP Monitoring columns [2].
    """
    logger.info("Extracting PTP rows from Productivity Report...")

    prod_bytes.seek(0)
    wb = openpyxl.load_workbook(prod_bytes)
    logger.info(f"Productivity sheets: {wb.sheetnames}")

    # Find Early SL sheet [4]
    early_sheet = _find_sheet(wb, [
        "Stat Result - BL Early SL",
        "Early SL",
        "Early",
        "BL Early"
    ])

    # Find Remedial SL sheet [4]
    remedial_sheet = _find_sheet(wb, [
        "Stat Result - BL Remedial SL",
        "Remedial SL",
        "Remedial",
        "BL Remedial"
    ])

    early_df    = _sheet_to_df_by_header(wb[early_sheet])
    remedial_df = _sheet_to_df_by_header(wb[remedial_sheet])

    logger.info(f"Early SL columns: {list(early_df.columns)}")
    logger.info(f"Remedial columns: {list(remedial_df.columns)}")
    logger.info(
        f"Early SL rows: {len(early_df)} | "
        f"Remedial rows: {len(remedial_df)}"
    )

    # Add Bucket to Remedial if missing
    if "Bucket" not in remedial_df.columns:
        remedial_df["Bucket"] = ""

    # Add Status if missing
    if "Status" not in early_df.columns:
        logger.warning("'Status' not in Early SL. Adding empty.")
        early_df["Status"] = ""
    if "Status" not in remedial_df.columns:
        logger.warning("'Status' not in Remedial SL. Adding empty.")
        remedial_df["Status"] = ""

    # Align columns before concat
    all_cols    = list(dict.fromkeys(
        list(early_df.columns) + list(remedial_df.columns)
    ))
    early_df    = early_df.reindex(columns=all_cols, fill_value="")
    remedial_df = remedial_df.reindex(columns=all_cols, fill_value="")

    combined = pd.concat([early_df, remedial_df], ignore_index=True)
    logger.info(f"Combined rows: {len(combined)}")
    logger.info(f"Combined columns: {list(combined.columns)}")

    # Fill and clean Status column
    combined["Status"] = (
        combined["Status"].fillna("").astype(str).str.strip()
    )

    # Filter PTP rows
    ptp_mask = combined["Status"].str.upper().str.startswith(
        PTP_STATUS_PREFIX.upper()
    )
    ptp_df = combined[ptp_mask.fillna(False)].copy()
    logger.info(f"PTP rows found: {len(ptp_df)}")

    if len(ptp_df) == 0:
        logger.warning(
            "No PTP rows found! Check if Status column contains "
            "PTP values. Sample statuses: "
            f"{combined['Status'].unique()[:10].tolist()}"
        )

    # Map columns to PTP Monitoring [2]
    mapped = pd.DataFrame()
    for src_col, dst_col in PTP_COLUMN_MAP.items():
        if src_col in ptp_df.columns:
            mapped[dst_col] = ptp_df[src_col].values
        else:
            logger.warning(
                f"Column '{src_col}' not found in Stat Result. "
                f"Filling empty."
            )
            mapped[dst_col] = ""

    logger.info(f"Mapped PTP columns: {list(mapped.columns)}")
    return mapped.reset_index(drop=True)


def populate_ptp(template_file, ptp_df: pd.DataFrame) -> io.BytesIO:
    """
    Appends PTP rows to the bottom of the PTP data sheet [2].
    Finds the correct sheet and header row dynamically.
    Returns modified workbook as BytesIO.
    """
    logger.info("Loading PTP Monitoring Template...")

    wb = openpyxl.load_workbook(template_file)
    logger.info(f"PTP template sheets: {wb.sheetnames}")

    sheet_name = _find_ptp_data_sheet(wb)
    ws         = wb[sheet_name]
    logger.info(f"Using PTP sheet: '{sheet_name}'")

    # Find header row in PTP sheet
    header_row = _find_header_row(ws, ["Date", "Name", "LAN", "Status"])
    last_row   = get_last_row(ws)

    # Use the greater of header_row or last_row to find next empty row
    next_row   = max(last_row, header_row) + 1
    logger.info(
        f"Header at row {header_row} | "
        f"Last data row: {last_row} | "
        f"Pasting at row: {next_row}"
    )

    logger.info(f"Pasting {len(ptp_df)} PTP rows at row {next_row}...")

    for r_idx, row in enumerate(
        ptp_df.itertuples(index=False), start=next_row
    ):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    logger.info("PTP Monitoring Report populated successfully.")
    return output