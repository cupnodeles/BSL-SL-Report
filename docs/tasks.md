# Implementation Tasks — BPI BL SL Automation

## ✅ Completed Tasks

### Phase 1 — Project Setup
- [x] Define project folder structure
- [x] Define technology stack
- [x] Create `requirements.txt`
- [x] Create `config.json`
- [x] Create `app.py` entry point
- [x] Create `src/__init__.py`

### Phase 2 — Core Modules
- [x] Write `src/utils.py`
  - [x] `get_last_row()`
  - [x] `get_yesterday_date()` (skips weekends)
  - [x] `get_output_filename()`
  - [x] `setup_logger()`

- [x] Write `src/dialer.py`
  - [x] Read `Overall Combined Summary` sheet [1]
  - [x] Dynamically find header row
  - [x] Extract 19 columns (exclude CYCLE)

- [x] Write `src/drr.py`
  - [x] Delete rows by Remark By filter [3]
  - [x] Delete rows by Remark starts with filter [3]
  - [x] Split Remedial vs Early SL
  - [x] Map Remedial columns
  - [x] Map Early SL columns (+ Bucket)

- [x] Write `src/productivity.py`
  - [x] Append dialer to `Penetration Per Day` [4]
  - [x] Clear + paste Remedial to `Stat Result - BL Remedial SL` [4]
  - [x] Clear + paste Early to `Stat Result - BL Early SL` [4]
  - [x] Paste as Values & Number Formats only

- [x] Write `src/ptp.py`
  - [x] Read both Stat Result sheets
  - [x] Filter Status starts with "PTP"
  - [x] Map to PTP Monitoring columns [2]
  - [x] Append to PTP Monitoring template [2]

- [x] Write `src/pivot.py`
  - [x] xlwings invisible Excel open
  - [x] RefreshAll() pivot tables
  - [x] Save and close silently

- [x] Write `src/encrypt.py`
  - [x] msoffcrypto file open [PASSWORD CONTEXT_REDACTED]

### Phase 3 — UI
- [x] Write `src/ui.py`
  - [x] Streamlit page config
  - [x] 4 file uploaders with drag & drop
  - [x] Pivot refresh toggle
  - [x] Run button with disabled state
  - [x] Progress bar (6 steps)
  - [x] Download buttons for both outputs
  - [x] Error handling with user-friendly messages

### Phase 4 — Documentation
- [x] Write `docs/requirements.md`
- [x] Write `docs/design.md`
- [x] Write `docs/tasks.md`
- [x] Write `docs/future_improvements.md`
- [x] Write `README.md`
- [x] Write `.gitignore`

---

## 🔲 Pending Tasks

### Phase 5 — Testing
- [ ] Write `tests/test_dialer.py`
  - [ ] Test valid Dialer Report extraction
  - [ ] Test missing sheet error
  - [ ] Test missing columns error
- [ ] Write `tests/test_drr.py`
  - [ ] Test TLEYVA/ZMONTANO row deletion
  - [ ] Test Remark starts with deletion
  - [ ] Test Remedial vs Early split
- [ ] Write `tests/test_productivity.py`
  - [ ] Test last row detection
  - [ ] Test data append
  - [ ] Test clear + paste
- [ ] Write `tests/test_ptp.py`
  - [ ] Test PTP filter
  - [ ] Test column mapping
  - [ ] Test append to template

### Phase 6 — Future Improvements
- [ ] Add Settings tab in UI
- [ ] Add file server path inputs in config
- [ ] Auto-detect latest file from file server path
- [ ] Save/load config from `config.json`
- [ ] Show auto-detected file timestamp in UI