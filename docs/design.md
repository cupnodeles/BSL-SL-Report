# Technical Design — BPI BL SL Automation

## 1. Architecture Overview

                                                     
┌─────────────────────────────────────────────────────┐
│                  Streamlit UI (ui.py)                │
│   Upload Files | Pivot Toggle | Run | Download       │
└────────────────────────┬────────────────────────────┘
                         │
              ┌──────────▼──────────┐
              │   app.py (Entry)    │
              └──────────┬──────────┘
                         │
         ┌───────────────┼───────────────┐
         │               │               │
         ▼               ▼               ▼
    dialer.py         drr.py         utils.py
         │               │
         │        ┌──────┴──────┐
         │        ▼             ▼
         │  map_remedial()  map_early()
         │        │             │
         ▼        ▼             ▼
    productivity.py ←───────────┘
         │
         ▼
      ptp.py
         │
         ▼
    pivot.py (optional — toggle ON only)
         │
         ▼
    encrypt.py
         │
         ▼
    Download Buttons (Streamlit)


## 2. Data Flow

### Step 1 — Dialer Report [1]

DIALER REPORT BPI BUSINESS LOAN SL *.xlsx
    Sheet: "Overall Combined Summary"
    ↓
extract_dialer_data()
    - Dynamically finds header row (looks for CYCLE + DATE)
    - Extracts 19 columns (DATE to CALL DROP RATE, excludes CYCLE)
    - Confirmed data row:
      DATE: 2026-07-21 | CLIENT: BPI BUSINESS LOAN SL
      ACCOUNTS: 487 | TOTAL DIALED: 2312
      PENETRATION RATE: 4.7474 | ...
    ↓
Returns: pd.DataFrame (19 columns)
    ↓
populate_productivity() → Appends to "Penetration Per Day" at last row


### Step 2 — DRR Cleaning [3]

Daily_Remark_Report*.xlsx
    ↓
clean_drr()
    - Delete rows: Remark By = TLEYVA or ZMONTANO
    - Delete rows: Remark starts with:
        * "Sub Special Status"
        * "System Auto Update"
        * "Updates when case reassign to another collector."
    ↓
split_drr()
    ├── Remedial: Batch No starts with "BPI BUSINESS LOAN REMEDIAL"
    └── Early SL: everything else
            e.g., BPI BL EARLY SL_01/07/2026


### Step 3 — Productivity Template Population [4]

SPM Productivity & Penetration Report_BSL-early&remedial_*.xlsx
    ↓
populate_productivity()
    ├── Sheet: "Penetration Per Day"
    │       → Append dialer data at last row (existing data preserved)
    │
    ├── Sheet: "Stat Result - BL Remedial SL"
    │       → Clear all rows except headers
    │       → Paste Remedial mapped data (Values & Number Formats only)
    │
    └── Sheet: "Stat Result - BL Early SL"
            → Clear all rows except headers
            → Paste Early SL mapped data (Values & Number Formats only)
            → Includes extra column: Bucket = Cycle
    ↓
Returns: io.BytesIO (populated workbook)


### Step 4 — PTP Monitoring Report [2]

Populated Productivity BytesIO
    ↓
extract_ptp_rows()
    - Read "Stat Result - BL Early SL"
    - Read "Stat Result - BL Remedial SL"
    - Filter: Status starts with "PTP"
      e.g., PTP NEW - PUSHBACK
            PTP NEW - UPFRONT PAYMENT FULL
            PTP NEW - ATU
            PTP OLD - PUSHBACK
    - Map columns to PTP Monitoring template
    ↓
SPM PTP Monitoring Report_BSL-early_*.xlsx
    ↓
populate_ptp()
    - Find last row dynamically
    - Append PTP rows at bottom of "PTP Monitoring" sheet
    ↓
Returns: io.BytesIO (populated workbook)


### Step 5 — Pivot Refresh (Toggle ON only)

IF Toggle ON (local + Excel installed):
    pivot.py
        - Save BytesIO to temp file
        - xw.App(visible=False)   ← Excel hidden in background
        - wb.api.RefreshAll()     ← Refreshes all pivots silently
        - wb.save() + wb.close()
        - app.quit()
        - Reload refreshed file back to BytesIO
        - Clean up temp files

IF Toggle OFF (web or no Excel):
    → Skip pivot refresh
    → Show reminder to manually refresh after download


### Step 6 — Encryption & Download

Both BytesIO outputs
    ↓
encrypt_file()
    - msoffcrypto applies file open password SPM*123
    ↓
Streamlit download_button()
    - Productivity: SPM Productivity & Penetration Report_BSL-early&remedial_{MMDDYYYY}.xlsx
    - PTP:          SPM PTP Monitoring Report_BSL-early_{MMDDYYYY}.xlsx


## 3. Module Design

### app.py — Entry Point

from src.ui import launch_app

if __name__ == "__main__":
    launch_app()

- Single responsibility: launch the Streamlit app


### src/ui.py — Streamlit UI

| Function                  | Responsibility                                          |
|---------------------------|---------------------------------------------------------|
| launch_app()              | Page config, layout, file uploaders, toggle, run button |
| run_automation()          | Orchestrates all 6 steps with progress bar              |
| _refresh_with_xlwings()   | Saves to temp, refreshes, reloads BytesIO               |

UI Layout:

┌─────────────────────────────────────────────┐
│        📊 BPI BL SL Automation              │
│─────────────────────────────────────────────│
│  📊 Dialer Report        📄 Prod Template   │
│  [ Upload / Drag & Drop ] [ Upload ]        │
│                                             │
│  📋 Daily Remark Report  📄 PTP Template    │
│  [ Upload / Drag & Drop ] [ Upload ]        │
│─────────────────────────────────────────────│
│  ⚙️ Settings                                │
│  🔄 Auto-refresh Pivots [ Toggle ON/OFF ]   │
│─────────────────────────────────────────────│
│  [ ▶ RUN AUTOMATION ]                       │
│                                             │
│  Step 1/6: Extracting Dialer data...        │
│  ████████████████░░ 85%                     │
│─────────────────────────────────────────────│
│  📥 Download Productivity Report            │
│  📥 Download PTP Monitoring Report          │
└─────────────────────────────────────────────┘


### src/dialer.py — Dialer Report Processing [1]

| Function                    | Input         | Output                   |
|-----------------------------|---------------|--------------------------|
| extract_dialer_data(file)   | Uploaded file | pd.DataFrame (19 cols)   |

Key Logic:
- Sheet: Overall Combined Summary
- Dynamically finds header row (searches for CYCLE + DATE)
- Strips and uppercases all column names
- Drops CYCLE, keeps 19 columns: DATE to CALL DROP RATE


### src/drr.py — DRR Cleaning & Mapping [3]

| Function          | Input        | Output                      |
|-------------------|--------------|-----------------------------|
| clean_drr(file)   | Uploaded file| Cleaned pd.DataFrame        |
| split_drr(df)     | Cleaned df   | (remedial_df, early_df)     |
| map_remedial(df)  | Remedial df  | Mapped pd.DataFrame         |
| map_early(df)     | Early df     | Mapped pd.DataFrame +Bucket |

Column Mapping — Remedial & Early SL:

| Template Column         | DRR Column         |
|-------------------------|--------------------|
| Date                    | Date               |
| Time                    | Time               |
| Name                    | Debtor             |
| LAN                     | Account No.        |
| Status                  | Status             |
| Remark                  | Remark             |
| Remark By               | Remark By          |
| PTP Amount              | PTP Amount         |
| PTP Date                | PTP Date           |
| Payment Amount          | Claim Paid Amount  |
| Payment Date            | Claim Paid Date    |
| Dialed Number           | Dialed Number      |
| Bucket (Early only)     | Cycle              |


### src/productivity.py — Productivity Template [4]

| Function                                              | Input                    | Output        |
|-------------------------------------------------------|--------------------------|---------------|
| populate_productivity(template,dialer,early,remedial) | Template + DataFrames    | io.BytesIO    |
| _clear_data_rows(ws)                                  | Worksheet                | None          |
| _paste_dataframe(ws, df)                              | Worksheet + DataFrame    | None          |

Sheet Actions:

| Sheet                        | Action                                      |
|------------------------------|---------------------------------------------|
| Penetration Per Day          | Find last row → append dialer data          |
| Stat Result - BL Remedial SL | Clear rows (keep header) → paste remedial   |
| Stat Result - BL Early SL    | Clear rows (keep header) → paste early      |


### src/ptp.py — PTP Monitoring Report [2]

| Function                       | Input               | Output           |
|--------------------------------|---------------------|------------------|
| extract_ptp_rows(prod_file)    | Productivity BytesIO| PTP pd.DataFrame |
| populate_ptp(template, ptp_df) | Template + DataFrame| io.BytesIO       |
| _sheet_to_df(ws)               | Worksheet           | pd.DataFrame     |

PTP Column Mapping:

| PTP Monitoring  | Source (Stat Result) |
|-----------------|----------------------|
| Date            | Date                 |
| Time            | Time                 |
| Name            | Name                 |
| LAN             | LAN                  |
| Status          | Status               |
| Remark          | Remark               |
| Remark By       | Remark By            |
| PTP Amount      | PTP Amount           |
| PTP Date        | PTP Date             |
| Dialed Number   | Dialed Number        |
| Bucket          | Bucket               |


### src/pivot.py — xlwings Pivot Refresh

| Function               | Input          | Output |
|------------------------|----------------|--------|
| refresh_pivots(filepath)| File path (str)| None   |

Behavior:
- xw.App(visible=False, add_book=False) → Excel completely hidden
- wb.api.RefreshAll() → refreshes ALL pivot tables in workbook
- Saves, closes, quits silently
- Only called when Toggle is ON


### src/encrypt.py — File Encryption

| Function              | Input        | Output              |
|-----------------------|--------------|---------------------|
| encrypt_file(file_bytes)| io.BytesIO | Encrypted io.BytesIO|

- Uses msoffcrypto-tool
- Applies file open password SPM*123


### src/utils.py — Shared Helpers

| Function                    | Returns   | Notes                        |
|-----------------------------|-----------|------------------------------|
| get_last_row(sheet)         | int       | Last row with data           |
| get_yesterday_date()        | datetime  | Mon → returns Fri            |
| get_output_filename(prefix) | str       | {prefix}_{MMDDYYYY}.xlsx     |
| setup_logger(log_dir)       | Logger    | Writes to /logs/             |


## 4. Output Files

| File                  | Naming Convention                                                              | Password |
|-----------------------|--------------------------------------------------------------------------------|----------|
| Productivity Report   | SPM Productivity & Penetration Report_BSL-early&remedial_{MMDDYYYY}.xlsx      | SPM*123  |
| PTP Monitoring Report | SPM PTP Monitoring Report_BSL-early_{MMDDYYYY}.xlsx                           | SPM*123  |

Date Logic:

| Today          | Output Date      |
|----------------|------------------|
| Tuesday–Friday | Yesterday        |
| Monday         | Previous Friday  |
| Saturday/Sunday| Previous Friday  |


## 5. Pivot Refresh Toggle

| Environment              | Toggle  | Method            | Result        |
|--------------------------|---------|-------------------|---------------|
| Local + Excel installed  | ON      | xlwings invisible | Auto-refresh  |
| Local no Excel           | OFF     | None              | Manual refresh|
| Streamlit Web            | OFF     | None              | Manual refresh|


## 6. Technology Stack

| Library          | Purpose                                          |
|------------------|--------------------------------------------------|
| streamlit        | Web UI, file upload, download buttons            |
| openpyxl         | Excel read/write, cell-level operations          |
| pandas           | Data filtering, cleaning, mapping                |
| msoffcrypto-tool | File open password encryption                    |
| xlwings          | Invisible Excel pivot refresh (local only)       |
| Python 3.8+      | Runtime                                          |


## 7. Error Handling

| Scenario                 | Handling                                          |
|--------------------------|---------------------------------------------------|
| Wrong sheet name         | ValueError with descriptive message               |
| Missing columns          | ValueError listing missing columns                |
| File not readable        | ValueError with file error details                |
| xlwings fails (no Excel) | Caught + user-friendly error message in UI        |
| Encryption fails         | Caught + logged + shown in UI                     |
| All errors               | Logged to /logs/automation_log_{MMDDYYYY}.txt     |