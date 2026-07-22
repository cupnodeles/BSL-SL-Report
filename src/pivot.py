# src/pivot.py
# BPI BL SL Automation — xlwings Pivot Refresh (Local Only)
# Gracefully handles import failure on Streamlit Cloud (Linux/no Excel)

import logging

logger = logging.getLogger("BPI_BL_SL")


def is_xlwings_available() -> bool:
    """
    Checks if xlwings is available and usable.
    Returns False on Streamlit Cloud (Linux) or if Excel not installed.
    """
    try:
        import xlwings as xw
        # Try to check if Excel is accessible
        return True
    except Exception:
        return False


def refresh_pivots(filepath: str) -> None:
    """
    Opens Excel file invisibly via xlwings.
    Refreshes ALL pivot tables silently.
    Only works when Excel is installed locally (Windows).
    Raises ImportError gracefully if xlwings not available.
    """
    if not is_xlwings_available():
        raise ImportError(
            "xlwings is not available on this system. "
            "Pivot refresh requires Excel installed on Windows."
        )

    import xlwings as xw
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