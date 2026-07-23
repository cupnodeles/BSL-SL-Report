# src/ptp.py
# BPI BL SL Automation — PTP Monitoring Report
# Fix: Correctly read Stat Result sheets after positional paste
# Fix: BytesIO position reset before reading
# Fix: Column matching by header name after finding header row

import io
import pandas as pd
import openpyxl
import logging
from src.utils import get_last_row

logger = logging.getLogger("BPI_BL_SL")

PTP_STATUS_PREFIX = "PTP"

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
    Tries exact match first, then partial.
    """
    logger.info(f"Available sheets: {wb.sheetnames}")

    # Exact match first
    for name in keywords:
        if name in wb.sheetnames:
            logger.info(f"Exact match: '{name}'")
            return name

    # Partial match
    for sheet_name in wb.sheetnames:
        for keyword in keywords:
            if keyword.lower() in sheet_name.lower():
                logger.info(f"Partial match '{keyword}' -> '{sheet_name}'")
                return sheet_name

    raise ValueError(
        f"Could not find sheet with keywords {keywords}. "
        f"Available: {wb.sheetnames}"
    )


def _sheet_to_df_positional(ws) -> pd.DataFrame:
    """
    Reads worksheet into DataFrame.
    Finds the actual header row by scanning for Date + Time + Name + LAN + Status.
    Uses those headers to name the columns.
    This correctly handles sheets where data was pasted positionally (A B C D)
    but the header row still has proper names [4][5].
    """
    all_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

    if not all_rows:
        logger.warning(f"Sheet '{ws.title}' is empty.")
        return pd.DataFrame()

    # Find header row — look for row with Date + Status + Name
    header_idx = None
    for idx, row in enumerate(all_rows):
        row_vals = [str(v).strip() if v is not None else "" for v in row]
        if (
            "Date"   in row_vals and
            "Status" in row_vals and
            "Name"   in row_vals
        ):
            header_idx = idx
            logger.info(
                f"Header row found at index {idx} "
                f"(row {idx + 1}) in '{ws.title}': "
                f"{row_vals[:8]}"
            )
            break

    if header_idx is None:
        logger.warning(
            f"No header row found in '{ws.title}'. "
            f"Using row 0 as header."
        )
        header_idx = 0

    # Build headers — deduplicate
    raw_headers = [
        str(h).strip() if h is not None else ""
        for h in all_rows[header_idx]
    ]
    seen    = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    logger.info(f"'{ws.title}' headers: {headers}")

    # Build DataFrame from data rows only
    data_rows = all_rows[header_idx + 1:]
    if not data_rows:
        logger.warning(f"No data rows in '{ws.title}'.")
        return pd.DataFrame(columns=headers)

    df = pd.DataFrame(data_rows, columns=headers)

    # Fill NaN and convert to string
    df = df.fillna("").astype(str)
    for col in df.columns:
        df[col] = df[col].str.strip()

    # Remove completely empty rows
    df = df[
        df.apply(lambda row: any(v != "" for v in row), axis=1)
    ].reset_index(drop=True)

    logger.info(f"'{ws.title}': {len(df)} data rows loaded.")
    return df


def _find_ptp_data_sheet(wb) -> str:
    """
    Finds the PTP data sheet in the PTP Monitoring template [2].
    Priority:
    1. Exact 'PTP List'
    2. Keyword match containing 'PTP' or 'List'
    3. Header scan for Date + Name + LAN + Status
    4. First sheet fallback
    """
    logger.info(f"PTP template sheets: {wb.sheetnames}")

    # Priority 1: Exact 'PTP List' [2]
    if "PTP List" in wb.sheetnames:
        logger.info("Found: 'PTP List'")
        return "PTP List"

    # Priority 2: Keyword match
    for sheet_name in wb.sheetnames:
        for keyword in ["PTP List", "PTP", "List", "Monitoring"]:
            if keyword.lower() in sheet_name.lower():
                logger.info(f"Found by keyword '{keyword}': '{sheet_name}'")
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
                logger.info(f"Found by header scan: '{sheet_name}'")
                return sheet_name

    # Priority 4: First sheet
    fallback = wb.sheetnames[0]
    logger.warning(f"Falling back to first sheet: '{fallback}'")
    return fallback


def extract_ptp_rows(prod_bytes: io.BytesIO) -> pd.DataFrame:
    """
    Reads the saved Productivity BytesIO [4][5].
    Finds Early SL and Remedial SL sheets dynamically.
    Reads using actual header row found by scanning.
    Filters rows where Status starts with 'PTP'.
    Returns combined mapped DataFrame for PTP Monitoring [2].
    """
    logger.info("Extracting PTP rows from Productivity Report...")

    # CRITICAL: Always seek to start before reading
    prod_bytes.seek(0)
    wb = openpyxl.load_workbook(prod_bytes, data_only=True)
    logger.info(f"Productivity sheets: {wb.sheetnames}")

    # Find Early SL sheet [4]
    early_sheet = _find_sheet(wb, [
        "Stat Result - BL Early SL",
        "Early SL",
        "BL Early",
        "Early"
    ])

    # Find Remedial SL sheet [4]
    remedial_sheet = _find_sheet(wb, [
        "Stat Result - BL Remedial SL",
        "Remedial SL",
        "BL Remedial",
        "Remedial"
    ])

    logger.info(f"Reading Early SL: '{early_sheet}'")
    early_df    = _sheet_to_df_positional(wb[early_sheet])

    logger.info(f"Reading Remedial SL: '{remedial_sheet}'")
    remedial_df = _sheet_to_df_positional(wb[remedial_sheet])

    logger.info(f"Early SL cols: {list(early_df.columns)}")
    logger.info(f"Remedial cols: {list(remedial_df.columns)}")
    logger.info(
        f"Early rows: {len(early_df)} | "
        f"Remedial rows: {len(remedial_df)}"
    )

    # Add Bucket to Remedial if missing
    if "Bucket" not in remedial_df.columns:
        remedial_df["Bucket"] = ""

    # Add Status if missing
    if "Status" not in early_df.columns:
        logger.warning("'Status' missing in Early SL — adding empty.")
        early_df["Status"] = ""
    if "Status" not in remedial_df.columns:
        logger.warning("'Status' missing in Remedial SL — adding empty.")
        remedial_df["Status"] = ""

    # Align columns before concat to avoid reindexing error
    all_cols    = list(dict.fromkeys(
        list(early_df.columns) + list(remedial_df.columns)
    ))
    early_df    = early_df.reindex(columns=all_cols, fill_value="")
    remedial_df = remedial_df.reindex(columns=all_cols, fill_value="")

    combined = pd.concat([early_df, remedial_df], ignore_index=True)
    logger.info(f"Combined rows: {len(combined)}")

    # Clean Status
    combined["Status"] = (
        combined["Status"].fillna("").astype(str).str.strip()
    )

    # Log sample statuses for debugging
    sample_statuses = combined["Status"].unique()[:15].tolist()
    logger.info(f"Sample statuses: {sample_statuses}")

    # Filter PTP rows
    ptp_mask = combined["Status"].str.upper().str.startswith(
        PTP_STATUS_PREFIX.upper()
    )
    ptp_df = combined[ptp_mask.fillna(False)].copy()
    logger.info(f"PTP rows found: {len(ptp_df)}")

    if len(ptp_df) == 0:
        logger.warning(
            f"NO PTP rows found! "
            f"Sample statuses were: {sample_statuses}"
        )

    # Map columns to PTP Monitoring template [2]
    mapped = pd.DataFrame()
    for src_col, dst_col in PTP_COLUMN_MAP.items():
        if src_col in ptp_df.columns:
            mapped[dst_col] = ptp_df[src_col].values
        else:
            logger.warning(f"'{src_col}' not found — filling empty.")
            mapped[dst_col] = ""

    logger.info(f"Mapped PTP cols: {list(mapped.columns)}")
    logger.info(f"Mapped PTP rows: {len(mapped)}")
    return mapped.reset_index(drop=True)


def populate_ptp(template_file, ptp_df: pd.DataFrame) -> io.BytesIO:
    """
    Appends PTP rows to the bottom of 'PTP List' sheet [2].
    Finds header row in PTP sheet dynamically.
    Pastes data by matching column positions to header names.
    Returns modified workbook as BytesIO.
    """
    logger.info("Loading PTP Monitoring Template...")

    if isinstance(template_file, io.BytesIO):
        template_file.seek(0)

    wb = openpyxl.load_workbook(template_file)
    logger.info(f"PTP template sheets: {wb.sheetnames}")

    sheet_name = _find_ptp_data_sheet(wb)
    ws         = wb[sheet_name]
    logger.info(f"Using PTP sheet: '{sheet_name}'")

    # Find header row in PTP sheet
    all_rows   = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
    header_idx = 0
    for idx, row in enumerate(all_rows):
        row_vals = [str(v).strip() if v is not None else "" for v in row]
        if "Date" in row_vals and "Name" in row_vals and "LAN" in row_vals:
            header_idx = idx
            logger.info(f"PTP header row at index {idx} (row {idx+1})")
            break

    header_row  = header_idx + 1  # 1-based
    ptp_headers = [
        str(h).strip() if h is not None else ""
        for h in all_rows[header_idx]
    ]
    logger.info(f"PTP sheet headers: {ptp_headers}")

    # Find last data row strictly below header
    last_data_row = header_row
    for row_idx in range(header_row + 1, ws.max_row + 1):
        cell_val = ws.cell(row=row_idx, column=1).value
        if cell_val is not None and str(cell_val).strip() != "":
            last_data_row = row_idx

    next_row = last_data_row + 1
    logger.info(
        f"Header at row {header_row} | "
        f"Last data row: {last_data_row} | "
        f"Pasting at row: {next_row}"
    )

    # Build column index map: header_name -> col_index (1-based)
    col_index_map = {
        name: idx + 1
        for idx, name in enumerate(ptp_headers)
        if name
    }
    logger.info(f"PTP col index map: {col_index_map}")

    # Paste PTP rows
    logger.info(f"Pasting {len(ptp_df)} PTP rows at row {next_row}...")
    for r_idx, row_data in enumerate(
        ptp_df.itertuples(index=False), start=next_row
    ):
        row_dict = dict(zip(ptp_df.columns, row_data))
        for col_name, value in row_dict.items():
            if col_name in col_index_map:
                c_idx = col_index_map[col_name]
                ws.cell(row=r_idx, column=c_idx).value = value

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    logger.info("PTP Monitoring Report populated successfully.")
    return output