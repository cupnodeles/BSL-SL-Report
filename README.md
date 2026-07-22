# BPI BSL SL Automation

A Streamlit-based automation tool for processing BPI Business Loan SL
daily reports — cleaning, mapping, populating templates, and generating
password-protected Excel output files.

---

## Project Structure

```
BPI_BL_SL_Automation/
|
|-- app.py                          <- Entry point (streamlit run app.py)
|-- requirements.txt                <- All pip dependencies
|-- config.json                     <- User settings
|-- .gitignore
|-- README.md
|
|-- /src/
|   |-- __init__.py
|   |-- ui.py                       <- Streamlit UI
|   |-- dialer.py                   <- Dialer Report processing
|   |-- drr.py                      <- DRR cleaning & mapping
|   |-- productivity.py             <- Productivity template population
|   |-- ptp.py                      <- PTP Monitoring Report population
|   |-- pivot.py                    <- xlwings pivot refresh (local only)
|   |-- encrypt.py                  <- File open password encryption
|   |-- utils.py                    <- Shared helpers
|
|-- /docs/
|   |-- requirements.md
|   |-- design.md
|   |-- tasks.md
|   |-- future_improvements.md
|
|-- /logs/                          <- Auto-generated per run
|
|-- /tests/
    |-- __init__.py
    |-- test_dialer.py
    |-- test_drr.py
    |-- test_productivity.py
    |-- test_ptp.py
```

---

## Setup

### 1. Install Dependencies
```
pip install -r requirements.txt
```

### 2. Run Locally
```
streamlit run app.py
```

### 3. Deploy to Web (Streamlit Cloud)
1. Push project to GitHub
2. Go to https://share.streamlit.io
3. Connect your GitHub repo
4. Set app.py as the main file
5. Click Deploy

---

## How to Use

1. Open the app in your browser
2. Upload the 4 required files:

   - Dialer Report
     DIALER REPORT BPI BUSINESS LOAN SL *.xlsx

   - Daily Remark Report
     Daily_Remark_Report*.xlsx

   - Productivity Template
     SPM Productivity & Penetration Report_BSL-early&remedial_*.xlsx

   - PTP Monitoring Template
     SPM PTP Monitoring Report_BSL-early_*.xlsx

3. Toggle Auto-refresh Pivots if running locally with Excel installed
4. Click RUN AUTOMATION
5. Download the 2 output files:

   - SPM Productivity & Penetration Report_BSL-early&remedial_{MMDDYYYY}.xlsx
   - SPM PTP Monitoring Report_BSL-early_{MMDDYYYY}.xlsx

---

## Output Files

Both output files are:
- Password protected with file open password SPM*123
- Named with yesterday's date suffix
- If today is Monday, date = previous Friday
- If today is Saturday or Sunday, date = previous Friday

---

## Pivot Refresh Toggle

| Mode                      | Toggle | Pivot Refresh                   |
|---------------------------|--------|---------------------------------|
| Local + Excel installed   | ON     | Auto-refresh via xlwings        |
| Local without Excel       | OFF    | Manual refresh after download   |
| Streamlit Web             | OFF    | Manual refresh after download   |

Note: When Toggle is OFF, a reminder will appear in the UI to
manually refresh pivot tables after opening the downloaded files.

---

## Input Files Reference

| File                  | Sheet Used                   | Action                        |
|-----------------------|------------------------------|-------------------------------|
| Dialer Report         | Overall Combined Summary     | Extract 19 cols, append       |
| Daily Remark Report   | Main sheet                   | Clean, split, map             |
| Productivity Template | Penetration Per Day          | Append dialer data            |
| Productivity Template | Stat Result - BL Early SL    | Clear + paste Early data      |
| Productivity Template | Stat Result - BL Remedial SL | Clear + paste Remedial data   |
| PTP Template          | PTP Monitoring               | Append PTP filtered rows      |

---

## DRR Cleaning Rules

Rows deleted where Remark By equals:
- TLEYVA
- ZMONTANO

Rows deleted where Remark starts with:
- Sub Special Status
- System Auto Update
- Updates when case reassign to another collector.

Remedial split:
- Batch No starts with: BPI BUSINESS LOAN REMEDIAL

Early SL split:
- Batch No does NOT start with: BPI BUSINESS LOAN REMEDIAL

---

## Column Mappings

### DRR to Productivity Template (Remedial & Early SL)

| Template Column   | DRR Column        |
|-------------------|-------------------|
| Date              | Date              |
| Time              | Time              |
| Name              | Debtor            |
| LAN               | Account No.       |
| Status            | Status            |
| Remark            | Remark            |
| Remark By         | Remark By         |
| PTP Amount        | PTP Amount        |
| PTP Date          | PTP Date          |
| Payment Amount    | Claim Paid Amount |
| Payment Date      | Claim Paid Date   |
| Dialed Number     | Dialed Number     |
| Bucket (Early SL) | Cycle             |

### Stat Result to PTP Monitoring

| PTP Monitoring  | Stat Result     |
|-----------------|-----------------|
| Date            | Date            |
| Time            | Time            |
| Name            | Name            |
| LAN             | LAN             |
| Status          | Status          |
| Remark          | Remark          |
| Remark By       | Remark By       |
| PTP Amount      | PTP Amount      |
| PTP Date        | PTP Date        |
| Dialed Number   | Dialed Number   |
| Bucket          | Bucket          |

---

## Dependencies

| Library          | Purpose                                     |
|------------------|---------------------------------------------|
| streamlit        | Web UI, file upload, download buttons       |
| openpyxl         | Excel read/write, cell-level operations     |
| pandas           | Data filtering, cleaning, mapping           |
| msoffcrypto-tool | File open password encryption               |
| xlwings          | Invisible Excel pivot refresh (local only)  |

Install all:
```
pip install -r requirements.txt
```

---

## Logs

Auto-generated log files are saved in /logs/ per run:
- Location: /logs/automation_log_{MMDDYYYY}.txt
- Contains: timestamps, step results, row counts, errors

---

## Future Improvements

See docs/future_improvements.md for planned features:
- Auto-detect latest file from file server folder path
- Settings tab with UNC path inputs (\\FileServer\Reports\...)
- Save/load config from config.json
- Auto-populate file fields on startup
- Write output directly to file server path

---

## Notes

- Pivot tables must be manually refreshed when Toggle is OFF
- All processing is done locally — no external API calls
- Sensitive data never leaves your machine
- Logs auto-create /logs/ directory if it does not exist
- Web version: files uploaded via browser, outputs downloaded via browser