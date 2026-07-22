# Requirements — BPI BL SL Automation

## 1. Overview
A Streamlit-based automation tool that processes BPI Business Loan SL
daily reports, cleans data, populates templates, and outputs
password-protected Excel files.

---

## 2. User Roles
| Role | Description |
|------|-------------|
| **Operator** | Uploads files, runs automation, downloads outputs |

---

## 3. Input Files
| File | Pattern | Source Sheet |
|------|---------|-------------|
| Dialer Report [1] | `DIALER REPORT BPI BUSINESS LOAN SL *.xlsx` | `Overall Combined Summary` |
| Daily Remark Report [3] | `Daily_Remark_Report*.xlsx` | Main sheet |
| Productivity Template [4] | `SPM Productivity & Penetration Report_BSL-early&remedial_*.xlsx` | Multiple sheets |
| PTP Monitoring Template [2] | `SPM PTP Monitoring Report_BSL-early_*.xlsx` | `PTP Monitoring` |

---

## 4. Functional Requirements

### 4.1 Dialer Report Processing [1]
- WHEN the operator uploads a Dialer Report,
  THE SYSTEM SHALL read the `Overall Combined Summary` sheet
- THE SYSTEM SHALL extract exactly 19 columns:
  DATE, CLIENT, ACCOUNTS, TOTAL DIALED, PENETRATION RATE,
  CONNECTED NU, CONNECTED UNIQUE, TOTAL RPC, PTP,
  CONNECTED % NU, CONNECTED % UNIQUE, RPC %, PTP %,
  TOTAL TALK TIME, TALK TIME AVE, TOTAL BALANCE,
  NEG DROP, SYSTEM DROP, CALL DROP RATE
- THE SYSTEM SHALL exclude the CYCLE column
- THE SYSTEM SHALL append the data to `Penetration Per Day`
  at the last row dynamically

### 4.2 DRR Cleaning [3]
- WHEN the operator uploads a Daily Remark Report,
  THE SYSTEM SHALL delete all rows where `Remark By` = TLEYVA or ZMONTANO
- THE SYSTEM SHALL delete all rows where `Remark` starts with:
  - "Sub Special Status"
  - "System Auto Update"
  - "Updates when case reassign to another collector."

### 4.3 Remedial SL Processing [3][4]
- THE SYSTEM SHALL filter rows where `Batch No` starts with
  "BPI BUSINESS LOAN REMEDIAL"
- THE SYSTEM SHALL delete all rows except headers in
  `Stat Result - BL Remedial SL`
- THE SYSTEM SHALL paste mapped data as Values & Number Formats only
- Column mapping:
  | Template | DRR |
  |----------|-----|
  | Date | Date |
  | Time | Time |
  | Name | Debtor |
  | LAN | Account No. |
  | Status | Status |
  | Remark | Remark |
  | Remark By | Remark By |
  | PTP Amount | PTP Amount |
  | PTP Date | PTP Date |
  | Payment Amount | Claim Paid Amount |
  | Payment Date | Claim Paid Date |
  | Dialed Number | Dialed Number |

### 4.4 Early SL Processing [3][4]
- THE SYSTEM SHALL filter rows where `Batch No` does NOT start with
  "BPI BUSINESS LOAN REMEDIAL"
- THE SYSTEM SHALL delete all rows except headers in
  `Stat Result - BL Early SL`
- THE SYSTEM SHALL paste mapped data as Values & Number Formats only
- Column mapping: Same as Remedial + additional:
  | Template | DRR |
  |----------|-----|
  | Bucket | Cycle |

### 4.5 PTP Monitoring Report [2]
- THE SYSTEM SHALL filter rows where `Status` starts with "PTP"
  from both Stat Result sheets
- THE SYSTEM SHALL append filtered rows to the bottom of
  `PTP Monitoring` sheet
- Column mapping:
  | PTP Monitoring | Report |
  |----------------|--------|
  | Date | Date |
  | Time | Time |
  | Name | Name |
  | LAN | LAN |
  | Status | Status |
  | Remark | Remark |
  | Remark By | Remark By |
  | PTP Amount | PTP Amount |
  | PTP Date | PTP Date |
  | Dialed Number | Dialed Number |
  | Bucket | Bucket |

### 4.6 Pivot Refresh (Toggle)
- WHEN the operator enables the pivot refresh toggle,
  THE SYSTEM SHALL use xlwings invisibly to refresh all pivot tables
- WHEN the toggle is disabled,
  THE SYSTEM SHALL skip pivot refresh
  and display a reminder to refresh manually

### 4.7 File Encryption
- THE SYSTEM SHALL encrypt both output files with
  file open password: `SPM*123`
- THE SYSTEM SHALL use `msoffcrypto` for encryption

### 4.8 Output Filenames
- Productivity Report:
  `SPM Productivity & Penetration Report_BSL-early&remedial_{MMDDYYYY}.xlsx`
- PTP Monitoring Report:
  `SPM PTP Monitoring Report_BSL-early_{MMDDYYYY}.xlsx`
- Date = yesterday's date
- IF yesterday is Saturday or Sunday,
  THE SYSTEM SHALL use the most recent Friday

### 4.9 Download
- THE SYSTEM SHALL provide download buttons for both output files
  via the Streamlit UI

---

## 5. Non-Functional Requirements
| # | Requirement |
|---|-------------|
| 1 | App must run locally via `streamlit run app.py` |
| 2 | App must be deployable to Streamlit Cloud (web) |
| 3 | Pivot refresh requires Excel installed locally |
| 4 | All sensitive data stays local (no external API calls) |
| 5 | Logs auto-generated per run in `/logs/` folder |

---

## 6. Constraints
- Python 3.8+
- Windows OS required for xlwings pivot refresh (toggle ON)
- `msoffcrypto-tool`, `openpyxl`, `pandas`, `streamlit`, `xlwings`