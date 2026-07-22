# Future Improvements — BPI BL SL Automation

## 1. Auto-Find Latest File from File Server Path

### Overview
Instead of manually uploading files each run, the automation will
automatically detect and load the **latest Excel file** from
pre-configured file server folder paths.

---

## 2. File Server Path Configuration

### `config.json` (Future)
```json
{
  "password": "SPM*123",
  "file_server": {
    "dialer_report_path":         "\\\\FileServer\\Reports\\Dialer\\",
    "drr_path":                   "\\\\FileServer\\Reports\\DRR\\",
    "productivity_template_path": "\\\\FileServer\\Templates\\Productivity\\",
    "ptp_template_path":          "\\\\FileServer\\Templates\\PTP\\",
    "output_path":                "\\\\FileServer\\Output\\"
  }
}