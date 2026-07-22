# src/pivot.py
# BPI BL SL Automation — xlwings Pivot Refresh (Local Only)
# Only called when Toggle is ON in the UI
# Requires: Microsoft Excel installed on local machine

import logging

logger = logging.getLogger("BPI_BSL_SL")


def refresh_pivots(filepath: str) -> None:
    """
    Opens Excel file invisibly via xlwings.
    Refreshes ALL pivot tables silently.
    Saves and closes Excel without showing any window.
    Only works when Excel is installed locally.
    """
    try:
        import xlwings as xw
    except ImportError:
        logger.error("xlwings is not installed. Run: pip install xlwings")
        raise

    logger.info(f"Refreshing pivots for: {filepath}")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(filepath)
        wb.api.RefreshAll()
        wb.save()
        wb.close()
        logger.info("Pivot refresh completed successfully.")
    except Exception as e:
        logger.error(f"Pivot refresh failed: {e}")
        raise
    finally:
        app.quit()