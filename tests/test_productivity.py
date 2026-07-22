# tests/test_productivity.py
# BPI BL SL Automation — Productivity Template Tests

import pytest
import pandas as pd
from src.utils import get_last_row
from unittest.mock import MagicMock


def test_get_last_row_empty_sheet():
    """Empty sheet should return 0."""
    ws = MagicMock()
    ws.iter_rows.return_value = []
    result = get_last_row(ws)
    assert result == 0


def test_get_last_row_with_data():
    """Should return correct last row index."""
    # TODO: test with actual openpyxl worksheet
    pass


def test_populate_productivity_appends_dialer():
    """Dialer data should be appended after last row."""
    # TODO: full integration test with fixture templates
    pass


def test_clear_data_rows_keeps_header():
    """All rows except row 1 (header) should be cleared."""
    pass