# ▚ BSL SL Report

[![Streamlit](https://img.shields.io/badge/Streamlit-live-57e389?style=flat-square)](https://bsl-sl-report.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.x-0b101d?style=flat-square)](./app.py)

**Live app:** [bsl-sl-report.streamlit.app](https://bsl-sl-report.streamlit.app/)

Dialer + DRR → Productivity & PTP Monitoring reports. Upload the Daily Remark Report, Dialer Report (optional), and the Productivity/PTP templates — download finished, pivot-ready workbooks.

## ▚ Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Deployed via [share.streamlit.io](https://share.streamlit.io) — push to `main` to update the live app.

## ▚ Stack

Python · Streamlit · pandas/openpyxl · `src/` pipeline (`launch_app` in `src/ui.py`)

## ▚ Notes

- Never commit real DRR data — fixtures only.
- See also: [XDAYS-SL-Export-Direct](https://github.com/cupnodeles/XDAYS-SL-Export-Direct) (sibling standardizer) and the in-browser port (private `bpi-automation-sl` repo).
