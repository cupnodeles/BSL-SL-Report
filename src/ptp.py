# src/ptp.py
# BPI BL SL Automation — PTP Monitoring Report
# Fix: Copy row format from row above when pasting PTP rows [2]
# Fix: Correctly read Stat Result sheets after positional paste [4][5]
# Fix: BytesIO position reset before reading

import io
import copy
import re
import pandas as pd
import openpyxl
import logging
from datetime import datetime, date, timedelta
from openpyxl.utils import get_column_letter, range_boundaries
from src.utils import (
    get_last_row, clean_date_value, strip_midnight_time,
    parse_date_value, parse_amount_value, month_key, month_label,
    EXCEL_DATE_FMT, EXCEL_AMOUNT_FMT,
)

logger = logging.getLogger("BPI_BL_SL")

PTP_STATUS_PREFIX = "PTP"

# Extra computed column in the PTP Monitoring template.
# Filled with VALUES (not formulas): WORKDAY(PTP Date, 2).
HOLDING_DATE_COL = "Holding Date"
HOLDING_WORKDAYS = 2

# Date columns in the PTP flow — time component is always dropped.
PTP_DATE_COLS = {"Date", "PTP Date"}

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
    Dynamically finds a sheet by exact name first, then partial match.
    Logs all available sheets.
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
    Handles sheets where data was pasted positionally (A B C D) [4][5].
    Deduplicates column names to prevent reindexing errors.
    Removes completely empty rows.
    """
    all_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

    if not all_rows:
        logger.warning(f"Sheet '{ws.title}' is empty.")
        return pd.DataFrame()

    # Find header row
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

    # Build headers with deduplication
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
    df = df.fillna("").astype(str)
    for col in df.columns:
        df[col] = df[col].str.strip()

    # Strip midnight "00:00:00" from date columns so PTP output
    # never carries a bogus time component.
    def _clean_cell(s):
        c = clean_date_value(s)
        if isinstance(c, (datetime, date)):
            return c.isoformat()
        return c

    def _clean_time_cell(s):
        if not isinstance(s, str):
            s = "" if s is None else str(s)
        s = s.strip()
        m = re.match(r"^\d{4}-\d{2}-\d{2}[T ](\d{2}:\d{2}(?::\d{2})?)", s)
        if m:
            t = m.group(1)
            return t if len(t) == 8 else f"{t}:00"
        m2 = re.match(r"^\d{1,2}/\d{1,2}/\d{4}\s+(\d{2}:\d{2}(?::\d{2})?)", s)
        if m2:
            t = m2.group(1)
            return t if len(t) == 8 else f"{t}:00"
        return s

    for col in df.columns:
        if col in PTP_DATE_COLS:
            df[col] = df[col].apply(_clean_cell)
        elif col == "Time":
            df[col] = df[col].apply(_clean_time_cell)

    # Remove completely empty rows
    df = df[
        df.apply(lambda row: any(v != "" for v in row), axis=1)
    ].reset_index(drop=True)

    logger.info(f"'{ws.title}': {len(df)} data rows loaded.")
    return df


def _find_ptp_data_sheet(wb) -> str:
    """
    Finds the PTP data sheet in PTP Monitoring template [2].
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

    # Priority 4: First sheet fallback
    fallback = wb.sheetnames[0]
    logger.warning(f"Falling back to first sheet: '{fallback}'")
    return fallback


def extract_ptp_rows(prod_bytes: io.BytesIO):
    """
    Reads the saved Productivity BytesIO [4][5].
    Finds Early SL and Remedial SL sheets dynamically.
    Reads using actual header row found by scanning.
    Filters rows where Status starts with 'PTP'.
    Drops rows with blank PTP Date AND zero PTP Amount (reported).
    Returns (mapped_df, removed_df) — removed_df holds dropped rows
    with a 'Reason' column for display in the UI.
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
            f"Sample statuses: {sample_statuses}"
        )

    # Map columns to PTP Monitoring template [2]
    mapped = pd.DataFrame()
    for src_col, dst_col in PTP_COLUMN_MAP.items():
        if src_col in ptp_df.columns:
            mapped[dst_col] = ptp_df[src_col].values
        else:
            logger.warning(f"'{src_col}' not found — filling empty.")
            mapped[dst_col] = ""

    # Final safety: ensure date cols carry no time component.
    def _final_date(s):
        c = clean_date_value(s)
        if isinstance(c, (datetime, date)):
            return c.isoformat()
        return c

    for _col in PTP_DATE_COLS:
        if _col in mapped.columns:
            mapped[_col] = mapped[_col].apply(_final_date)

    # ---------------------------------------------------------- #
    # Drop rows with blank PTP Date AND zero PTP Amount.
    # Blank amount is NOT zero — only an explicit 0/0.00 counts.
    # Dropped rows are returned separately for the UI report.
    # ---------------------------------------------------------- #
    removed = pd.DataFrame()
    if "PTP Date" in mapped.columns and "PTP Amount" in mapped.columns:
        blank_date = mapped["PTP Date"].apply(
            lambda s: (
                True
                if s is None
                else str(s).strip() in ("", "nan", "None", "NaT", "NaN", "nat")
            )
        )
        def _is_zero_amount(v) -> bool:
            """True only for an explicit 0/0.00 (blank/unparseable = False)."""
            try:
                amt = parse_amount_value(v)
            except Exception:
                return False
            return amt is not None and amt == 0

        zero_amt = mapped["PTP Amount"].apply(_is_zero_amount)
        drop_mask = (blank_date & zero_amt).fillna(False).astype(bool)
        if bool(drop_mask.any()):
            removed = mapped[drop_mask].copy()
            # 1-based position in the pulled PTP set, for the UI report
            removed.insert(0, "Source Row #", removed.index + 1)
            removed["Reason"] = "Blank PTP Date + zero PTP Amount"
            mapped = mapped[~drop_mask].copy()
            logger.warning(
                f"Dropped {len(removed)} PTP row(s) with blank PTP Date "
                f"and zero PTP Amount."
            )
    else:
        logger.warning(
            "PTP Date/PTP Amount columns missing — skipping blank/zero filter."
        )

    logger.info(f"Mapped PTP cols: {list(mapped.columns)}")
    logger.info(f"Mapped PTP rows: {len(mapped)}")
    if len(removed):
        logger.info(f"Removed PTP rows: {len(removed)}")
    return mapped.reset_index(drop=True), removed.reset_index(drop=True)


def _add_workdays(start: date, n: int) -> date:
    """Adds n workdays (Mon–Fri) to start. Mirrors Excel WORKDAY(d, n)."""
    d = start
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def _holding_date_value(ptp_date_value):
    """
    Computes WORKDAY(PTP Date, 2) in Python so 'Holding Date' is
    pasted as a VALUE (real Excel date), not a formula.
    Blank/unparseable PTP Date -> "" (cell left blank).
    """
    if ptp_date_value is None:
        return ""
    if isinstance(ptp_date_value, datetime):
        base = ptp_date_value.date()
    elif isinstance(ptp_date_value, date):
        base = ptp_date_value
    else:
        s = str(ptp_date_value).strip()
        if s in ("", "nan", "None", "NaT", "NaN", "nat"):
            return ""
        try:
            parsed = pd.to_datetime(s, errors="coerce")
        except Exception:
            return ""
        if parsed is None or pd.isna(parsed):
            return ""
        base = parsed.date() if isinstance(parsed, datetime) else parsed
        if not isinstance(base, date):
            return ""
    return _add_workdays(base, HOLDING_WORKDAYS)


def _fill_holding_date(ws, col_index_map: dict,
                       first_row: int, last_row: int):
    """
    Fills 'Holding Date' with VALUES (WORKDAY of each row's PTP Date
    + 2 workdays) for rows first_row..last_row. Skipped silently when
    the template has no 'Holding Date' column.
    """
    if last_row < first_row:
        return
    holding_col = col_index_map.get(HOLDING_DATE_COL)
    if not holding_col:
        logger.info("No 'Holding Date' column — skipping fill.")
        return
    ptp_date_col = col_index_map.get("PTP Date")
    if not ptp_date_col:
        logger.warning("No 'PTP Date' column — cannot compute Holding Date.")
        return
    filled = 0
    for r in range(first_row, last_row + 1):
        ptp_val = ws.cell(row=r, column=ptp_date_col).value
        val = _holding_date_value(ptp_val)
        cell = ws.cell(row=r, column=holding_col)
        cell.value = val
        if isinstance(val, date):
            cell.number_format = EXCEL_DATE_FMT
        filled += 1
    logger.info(
        f"Holding Date values filled for {filled} row(s) "
        f"(rows {first_row} to {last_row})."
    )


def _expand_tables(ws, header_row: int, new_last_row: int):
    """
    Expands every Excel Table on ws so its ref covers header_row..
    new_last_row (keeps existing columns). Required so appended PTP
    rows join the table — structured refs like [@[PTP Date]] and
    filters/pivots keep working.
    """
    try:
        tables = list(ws.tables.values())
    except Exception as e:
        logger.warning(f"Could not list tables: {e}")
        return
    if not tables:
        logger.info("No Excel tables to expand.")
        return
    for tbl in tables:
        if isinstance(tbl, str):
            continue
        try:
            old_ref = tbl.ref
            min_col, min_row, max_col, max_row = range_boundaries(old_ref)
        except Exception as e:
            logger.warning(f"Could not parse table ref: {e}")
            continue
        if new_last_row <= max_row and min_row == header_row:
            continue
        new_ref = (
            f"{get_column_letter(min_col)}{min(header_row, min_row)}:"
            f"{get_column_letter(max_col)}{max(new_last_row, max_row)}"
        )
        try:
            tbl.ref = new_ref
            if getattr(tbl, "autoFilter", None) is not None:
                try:
                    tbl.autoFilter.ref = new_ref
                except Exception:
                    pass
            logger.info(f"Expanded table '{tbl.displayName}': "
                        f"{old_ref} -> {new_ref}")
        except Exception as e:
            logger.warning(f"Could not expand table: {e}")


# Shared month helpers live in src.utils (also used by productivity.py
# for the Penetration monthly reset).
_month_key = month_key
_month_label = month_label


def _incoming_latest_month(ptp_df: pd.DataFrame):
    """Latest (year, month) across incoming PTP Date (fallback Date)."""
    for col in ("PTP Date", "Date"):
        if col in ptp_df.columns and len(ptp_df):
            months = [
                m for m in (_month_key(v) for v in ptp_df[col].tolist())
                if m is not None
            ]
            if months:
                return max(months)
    return None


def _existing_probe(ws, header_row: int, col_index_map: dict):
    """
    Scans current PTP List rows for the latest (year, month) and the
    data-row count. Prefers the PTP Date column, falls back to Date.
    Returns (latest_month_or_None, data_row_count).
    """
    target_col = col_index_map.get("PTP Date") or col_index_map.get("Date")
    if not target_col:
        return None, 0
    months = []
    count = 0
    for row_idx in range(header_row + 1, ws.max_row + 1):
        first = ws.cell(row=row_idx, column=1).value
        if first is None or str(first).strip() == "":
            continue
        count += 1
        m = _month_key(ws.cell(row=row_idx, column=target_col).value)
        if m is not None:
            months.append(m)
    return (max(months) if months else None), count


def _reset_ptp_data(ws, header_row: int) -> int:
    """
    Monthly reset: clears all data values below the header row and
    shrinks every Excel Table to header-only (so filters show no ghost
    blank rows). Headers, formats and table styles are untouched.
    Returns the number of cleared data rows.
    """
    cleared = 0
    for row_idx in range(header_row + 1, ws.max_row + 1):
        has_data = False
        for c_idx in range(1, ws.max_column + 1):
            v = ws.cell(row=row_idx, column=c_idx).value
            if v is not None and str(v).strip() != "":
                has_data = True
                ws.cell(row=row_idx, column=c_idx).value = None
        if has_data:
            cleared += 1
    try:
        tables = list(ws.tables.values())
    except Exception as e:
        logger.warning(f"Could not list tables for reset: {e}")
        return cleared
    for tbl in tables:
        if isinstance(tbl, str):
            continue
        try:
            min_col, _, max_col, _ = range_boundaries(tbl.ref)
            new_ref = (
                f"{get_column_letter(min_col)}{header_row}:"
                f"{get_column_letter(max_col)}{header_row}"
            )
            tbl.ref = new_ref
            if getattr(tbl, "autoFilter", None) is not None:
                try:
                    tbl.autoFilter.ref = new_ref
                except Exception:
                    pass
            logger.info(f"Reset table '{tbl.displayName}' to {new_ref}")
        except Exception as e:
            logger.warning(f"Could not reset table: {e}")
    return cleared


def populate_ptp(template_file, ptp_df: pd.DataFrame,
                 new_month: bool = False):
    """
    Appends PTP rows to the bottom of 'PTP List' sheet [2].
    Copies row format from the row above (fixes formatting mismatch).
    Finds header row dynamically.
    Pastes data by matching column positions to header names.
    Auto-fills 'Holding Date' with VALUES (WORKDAY of PTP Date + 2).
    Expands Excel Table(s) to cover new rows.
    Strips midnight "00:00:00" from date values.
    Monthly reset: when new_month is True, or incoming data's month is
    newer than existing rows' month, all existing data rows are cleared
    first (tables shrunk to header-only, then re-expanded).
    Returns (BytesIO, info dict) with reset details for the UI.
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
            logger.info(f"PTP header row at index {idx} (row {idx + 1})")
            break

    header_row  = header_idx + 1  # 1-based
    ptp_headers = [
        str(h).strip() if h is not None else ""
        for h in all_rows[header_idx]
    ]
    logger.info(f"PTP sheet headers: {ptp_headers}")

    # Build column index map: header_name -> col_index (1-based)
    col_index_map = {
        name: idx + 1
        for idx, name in enumerate(ptp_headers)
        if name
    }
    logger.info(f"PTP col index map: {col_index_map}")

    # ---------------------------------------------------------- #
    # Monthly reset decision (PTP List only).
    # Manual toggle wins; otherwise auto-fire when incoming data's
    # latest month is newer than existing rows' latest month.
    # Never wipes when there is nothing incoming to replace with.
    # ---------------------------------------------------------- #
    reset_info = {
        "reset": False, "cleared": 0, "mode": "same-month",
        "incoming": None, "existing": None, "existing_rows": 0,
    }
    incoming_month, (existing_month, existing_rows) = (
        _incoming_latest_month(ptp_df),
        _existing_probe(ws, header_row, col_index_map),
    )
    reset_info["existing_rows"] = existing_rows
    if incoming_month is not None:
        reset_info["incoming"] = _month_label(incoming_month)
    if existing_month is not None:
        reset_info["existing"] = _month_label(existing_month)

    do_reset, reset_mode = False, "same-month"
    if len(ptp_df) == 0:
        reset_mode = "no-incoming"
        logger.info("No incoming PTP rows — monthly reset disabled.")
    elif new_month and existing_rows > 0:
        do_reset, reset_mode = True, "manual"
    elif new_month:
        reset_mode = "manual-empty"
        logger.info("New Month toggle ON but PTP sheet is already empty.")
    elif (
        incoming_month is not None
        and existing_month is not None
        and incoming_month > existing_month
    ):
        do_reset, reset_mode = True, "auto"

    if do_reset:
        cleared = _reset_ptp_data(ws, header_row)
        reset_info.update(
            {"reset": True, "cleared": cleared, "mode": reset_mode}
        )
        logger.warning(
            f"Monthly PTP reset ({reset_mode}): cleared {cleared} "
            f"existing row(s) "
            f"(existing {reset_info['existing']} -> "
            f"incoming {reset_info['incoming']})."
        )
    else:
        reset_info["mode"] = reset_mode
        logger.info(
            f"No monthly reset (mode={reset_mode}; "
            f"existing={reset_info['existing']}, "
            f"incoming={reset_info['incoming']})."
        )

    # Find last data row strictly below header (after any reset)
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

    # Get max column count for format copying
    max_col = max(col_index_map.values()) if col_index_map else ws.max_column

    def copy_row_format(source_row: int, target_row: int):
        """
        Copies cell format from source_row to target_row.
        Copies: number_format, font, fill, border, alignment, row height.
        This fixes the formatting mismatch between existing and new rows [2].
        """
        for col_idx in range(1, max_col + 1):
            src = ws.cell(row=source_row, column=col_idx)
            tgt = ws.cell(row=target_row, column=col_idx)
            tgt.number_format = src.number_format
            if src.font:
                tgt.font      = copy.copy(src.font)
            if src.fill:
                tgt.fill      = copy.copy(src.fill)
            if src.border:
                tgt.border    = copy.copy(src.border)
            if src.alignment:
                tgt.alignment = copy.copy(src.alignment)
        if ws.row_dimensions[source_row].height:
            ws.row_dimensions[target_row].height = (
                ws.row_dimensions[source_row].height
            )

    # Paste PTP rows with format copied from row above [2]
    logger.info(f"Pasting {len(ptp_df)} PTP rows at row {next_row}...")
    actual_row = next_row
    for row_data in ptp_df.itertuples(index=False):
        # Skip completely empty rows
        values = [
            str(v).strip() if v is not None else ""
            for v in row_data
        ]
        if not any(v != "" for v in values):
            logger.warning(f"Skipping blank row at {actual_row}")
            continue

        # Copy format from the row directly above
        # This ensures new rows match existing formatting [2]
        source_format_row = actual_row - 1
        if source_format_row >= header_row + 1:
            copy_row_format(source_format_row, actual_row)

        # Paste values by column name matching.
        # Date columns -> REAL Excel dates (pivot-friendly, dd/mm/yyyy).
        # PTP Amount   -> REAL numbers (#,##0.00, no green flag).
        row_dict = dict(zip(ptp_df.columns, row_data))
        for col_name, value in row_dict.items():
            if col_name in col_index_map:
                c_idx = col_index_map[col_name]
                cell = ws.cell(row=actual_row, column=c_idx)
                if col_name in PTP_DATE_COLS:
                    d = parse_date_value(value)
                    if d is not None:
                        cell.value = d
                        cell.number_format = EXCEL_DATE_FMT
                    else:
                        cell.value = ""
                elif col_name == "PTP Amount":
                    amt = parse_amount_value(value)
                    if amt is not None:
                        cell.value = amt
                        cell.number_format = EXCEL_AMOUNT_FMT
                    else:
                        if value is not None and str(value).strip() not in (
                            "", "nan", "None", "NaT", "NaN", "nat",
                        ):
                            logger.warning(
                                f"Unparseable PTP Amount "
                                f"'{value}' at row {actual_row} — "
                                f"left blank."
                            )
                        cell.value = ""
                else:
                    cleaned = strip_midnight_time(value)
                    if isinstance(cleaned, datetime):
                        cell.value = cleaned
                    else:
                        cell.value = cleaned

        actual_row += 1

    logger.info(
        f"Total PTP rows pasted: {actual_row - next_row} "
        f"(rows {next_row} to {actual_row - 1})"
    )

    last_written = actual_row - 1

    # ---------------------------------------------------------- #
    # Holding Date: WORKDAY(PTP Date, 2) pasted as VALUES
    # ---------------------------------------------------------- #
    _fill_holding_date(ws, col_index_map, next_row, last_written)

    # ---------------------------------------------------------- #
    # Expand Excel Table(s) to include the new rows so filters
    # and pivots keep working.
    # ---------------------------------------------------------- #
    if last_written >= next_row:
        _expand_tables(ws, header_row, last_written)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    logger.info("PTP Monitoring Report populated successfully.")
    return output, reset_info