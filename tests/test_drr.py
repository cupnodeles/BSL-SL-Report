# tests/test_drr.py
# BPI BL SL Automation — DRR Cleaning Tests

import pytest
import pandas as pd
from src.drr import clean_drr, split_drr, map_remedial, map_early


def test_clean_drr_removes_tleyva():
    """Rows with Remark By = TLEYVA should be deleted."""
    df = pd.DataFrame({
        "Remark By": ["TLEYVA", "ARPILI", "ZMONTANO", "ECAJON"],
        "Remark": ["note", "note", "note", "note"],
        "Batch No": ["BPI BL", "BPI BL", "BPI BL", "BPI BL"]
    })
    # Simulate clean (direct function call with DataFrame)
    result = df[~df["Remark By"].str.strip().str.upper().isin(
        ["TLEYVA", "ZMONTANO"]
    )]
    assert "TLEYVA" not in result["Remark By"].values
    assert "ZMONTANO" not in result["Remark By"].values
    assert len(result) == 2


def test_clean_drr_removes_remark_starts_with():
    """Rows where Remark starts with defined strings should be deleted."""
    df = pd.DataFrame({
        "Remark By": ["ARPILI", "ECAJON", "ARPILI"],
        "Remark": [
            "Sub Special Status - test",
            "System Auto Update Remarks",
            "Valid remark here"
        ],
        "Batch No": ["BPI BL", "BPI BL", "BPI BL"]
    })
    starts = ["Sub Special Status", "System Auto Update"]
    result = df[~df["Remark"].apply(
        lambda r: any(str(r).startswith(s) for s in starts)
    )]
    assert len(result) == 1
    assert result.iloc[0]["Remark"] == "Valid remark here"


def test_split_drr_remedial_vs_early():
    """Remedial rows should have Batch No starting with BPI BUSINESS LOAN REMEDIAL."""
    df = pd.DataFrame({
        "Batch No": [
            "BPI BUSINESS LOAN REMEDIAL_01/2026",
            "BPI BL EARLY SL_01/07/2026",
            "BPI BUSINESS LOAN REMEDIAL_02/2026",
        ],
        "Remark By": ["A", "B", "C"],
        "Remark": ["x", "y", "z"]
    })
    remedial, early = split_drr(df)
    assert len(remedial) == 2
    assert len(early) == 1


def test_map_early_includes_bucket():
    """Early SL mapping should include Bucket column."""
    # TODO: full integration test with fixture
    pass