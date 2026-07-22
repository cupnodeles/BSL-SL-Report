# tests/test_ptp.py
# BPI BL SL Automation — PTP Monitoring Tests

import pytest
import pandas as pd
from src.ptp import PTP_STATUS_PREFIX


def test_ptp_filter_starts_with():
    """Only rows where Status starts with 'PTP' should be kept."""
    df = pd.DataFrame({
        "Status": [
            "PTP NEW - PUSHBACK",
            "PTP OLD - ATU",
            "RNA",
            "BUSY",
            "PTP NEW - UPFRONT PAYMENT FULL"
        ]
    })
    result = df[
        df["Status"].str.strip().str.upper()
        .str.startswith(PTP_STATUS_PREFIX.upper())
    ]
    assert len(result) == 3


def test_ptp_column_mapping():
    """Mapped DataFrame should have correct PTP Monitoring columns."""
    # TODO: full integration test with fixture
    pass


def test_populate_ptp_appends_at_bottom():
    """PTP rows should be appended after last existing row."""
    pass