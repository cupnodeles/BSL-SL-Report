# src/ui.py
# BSL SL Automation — Streamlit UI
# Fix: ModuleNotFoundError on Streamlit Cloud
# Fix: xlwings disabled on non-Windows
# Fix: Download buttons persist via session_state

import os
import io
import platform
import streamlit as st
import logging
from src.utils import setup_logger, get_output_filename
from src.dialer import extract_dialer_data
from src.drr import clean_drr, split_drr, map_remedial, map_early
from src.productivity import populate_productivity
from src.ptp import extract_ptp_rows, populate_ptp
from src.encrypt import encrypt_file, decrypt_file, is_encrypted

logger = setup_logger()

PROD_PREFIX  = "SPM Productivity & Penetration Report_BSL-early&remedial"
PTP_PREFIX   = "SPM PTP Monitoring Report_BSL-early"
IS_WINDOWS   = platform.system() == "Windows"


def _init_session_state():
    """Initialize all session state variables."""
    defaults = {
        "prod_encrypted":   None,
        "ptp_encrypted":    None,
        "prod_filename":    None,
        "ptp_filename":     None,
        "automation_done":  False,
        "automation_error": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def launch_app():
    st.set_page_config(
        page_title="BSL SL Automation",
        page_icon="📊",
        layout="centered"
    )

    # Initialize session state FIRST
    _init_session_state()

    st.title("📊 BSL SL Automation")
    st.markdown("---")

    # ------------------------------------------------------------------ #
    # FILE UPLOADS
    # ------------------------------------------------------------------ #
    st.subheader("📁 Upload Input Files")

    col1, col2 = st.columns(2)
    with col1:
        dialer_file = st.file_uploader(
            "📊 Dialer Report",
            type=["xlsx"],
            help="DIALER REPORT BUSINESS LOAN SL *.xlsx",
            key="uploader_dialer"
        )
        drr_file = st.file_uploader(
            "📋 Daily Remark Report",
            type=["xlsx"],
            help="Daily_Remark_Report*.xlsx",
            key="uploader_drr"
        )
    with col2:
        prod_template = st.file_uploader(
            "📄 Productivity Template",
            type=["xlsx"],
            help="SPM Productivity & Penetration Report_BSL-early&remedial_*.xlsx",
            key="uploader_prod"
        )
        ptp_template = st.file_uploader(
            "📄 PTP Monitoring Template",
            type=["xlsx"],
            help="SPM PTP Monitoring Report_BSL-early_*.xlsx",
            key="uploader_ptp"
        )

    st.markdown("---")

    # ------------------------------------------------------------------ #
    # SETTINGS
    # ------------------------------------------------------------------ #
    st.subheader("⚙️ Settings")

    # Only enable pivot refresh toggle on Windows with Excel
    if not IS_WINDOWS:
        st.warning(
            f"⚠️ Auto-refresh Pivot Tables is only available on Windows "
            f"with Excel installed. Running on: **{platform.system()}**"
        )
        auto_refresh = False
        st.toggle(
            "🔄 Auto-refresh Pivot Tables",
            value=False,
            disabled=True,
            help="Not available on this platform.",
            key="toggle_refresh"
        )
    else:
        auto_refresh = st.toggle(
            "🔄 Auto-refresh Pivot Tables",
            value=False,
            help="Requires Microsoft Excel installed locally on Windows.",
            key="toggle_refresh"
        )
        if auto_refresh:
            st.info(
                "✅ Pivot tables will be refreshed automatically via xlwings."
            )
        else:
            st.warning(
                "⚠️ Pivot tables will NOT be refreshed. "
                "Please refresh manually after opening the files."
            )

    st.markdown("---")

    # ------------------------------------------------------------------ #
    # RUN BUTTON
    # ------------------------------------------------------------------ #
    all_uploaded = all([dialer_file, drr_file, prod_template, ptp_template])

    if not all_uploaded:
        st.info("👆 Please upload all 4 files to proceed.")

    if st.button(
        "▶ RUN AUTOMATION",
        disabled=not all_uploaded,
        type="primary",
        key="btn_run"
    ):
        # Reset outputs before new run
        st.session_state.prod_encrypted  = None
        st.session_state.ptp_encrypted   = None
        st.session_state.prod_filename   = None
        st.session_state.ptp_filename    = None
        st.session_state.automation_done  = False
        st.session_state.automation_error = None

        run_automation(
            dialer_file, drr_file,
            prod_template, ptp_template,
            auto_refresh
        )

    # ------------------------------------------------------------------ #
    # DOWNLOAD SECTION — persists via session_state
    # Rendered OUTSIDE run_automation so it survives reruns
    # from download button clicks
    # ------------------------------------------------------------------ #
    if st.session_state.automation_done:
        st.markdown("---")

        if st.session_state.automation_error:
            st.error(
                f"❌ Automation failed: "
                f"{st.session_state.automation_error}"
            )
        elif (
            st.session_state.prod_encrypted is not None and
            st.session_state.ptp_encrypted  is not None
        ):
            st.subheader("📥 Download Output Files")
            st.success(
                "✅ Automation complete! Download your files below."
            )

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Productivity Report",
                    data=st.session_state.prod_encrypted,
                    file_name=st.session_state.prod_filename,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    key="dl_prod"
                )
            with col2:
                st.download_button(
                    label="📥 Download PTP Monitoring Report",
                    data=st.session_state.ptp_encrypted,
                    file_name=st.session_state.ptp_filename,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    key="dl_ptp"
                )

            if not auto_refresh:
                st.info(
                    "💡 Remember to manually refresh pivot tables "
                    "after opening the files in Excel."
                )


def run_automation(
    dialer_file, drr_file,
    prod_template, ptp_template,
    auto_refresh: bool
):
    """Main automation runner with progress tracking."""

    progress = st.progress(0)
    status   = st.empty()

    try:
        # ---------------------------------------------------------- #
        # STEP 1: Extract Dialer Data [1]
        # ---------------------------------------------------------- #
        status.info("📊 Step 1/6: Extracting Dialer Report data...")
        dialer_bytes = io.BytesIO(dialer_file.read())
        if is_encrypted(dialer_bytes):
            logger.info("Dialer file encrypted. Decrypting...")
            dialer_bytes = decrypt_file(dialer_bytes)
        dialer_df = extract_dialer_data(dialer_bytes)
        logger.info(f"Dialer rows extracted: {len(dialer_df)}")
        progress.progress(15)

        # ---------------------------------------------------------- #
        # STEP 2: Clean DRR [3]
        # ---------------------------------------------------------- #
        status.info("🧹 Step 2/6: Cleaning Daily Remark Report...")
        drr_bytes = io.BytesIO(drr_file.read())
        if is_encrypted(drr_bytes):
            logger.info("DRR file encrypted. Decrypting...")
            drr_bytes = decrypt_file(drr_bytes)
        clean_df                = clean_drr(drr_bytes)
        remedial_raw, early_raw = split_drr(clean_df)
        remedial_df             = map_remedial(remedial_raw)
        early_df                = map_early(early_raw)
        logger.info(
            f"Remedial rows: {len(remedial_df)} | "
            f"Early rows: {len(early_df)}"
        )
        progress.progress(35)

        # ---------------------------------------------------------- #
        # STEP 3: Populate Productivity Template [4]
        # ---------------------------------------------------------- #
        status.info("📄 Step 3/6: Populating Productivity Report...")
        prod_bytes = io.BytesIO(prod_template.read())
        if is_encrypted(prod_bytes):
            logger.info("Productivity template encrypted. Decrypting...")
            prod_bytes = decrypt_file(prod_bytes)
        prod_out = populate_productivity(
            prod_bytes, dialer_df, early_df, remedial_df
        )
        logger.info("Productivity template populated.")
        progress.progress(55)

        # ---------------------------------------------------------- #
        # STEP 4: Populate PTP Monitoring [2]
        # ---------------------------------------------------------- #
        status.info("📋 Step 4/6: Populating PTP Monitoring Report...")
        ptp_bytes = io.BytesIO(ptp_template.read())
        if is_encrypted(ptp_bytes):
            logger.info("PTP template encrypted. Decrypting...")
            ptp_bytes = decrypt_file(ptp_bytes)
        prod_out.seek(0)
        ptp_df  = extract_ptp_rows(prod_out)
        prod_out.seek(0)
        ptp_out = populate_ptp(ptp_bytes, ptp_df)
        logger.info(f"PTP rows pasted: {len(ptp_df)}")
        progress.progress(70)

        # ---------------------------------------------------------- #
        # STEP 5: Pivot Refresh (Windows + Toggle ON only)
        # ---------------------------------------------------------- #
        if auto_refresh and IS_WINDOWS:
            status.info(
                "🔄 Step 5/6: Refreshing pivot tables via xlwings..."
            )
            prod_out.seek(0)
            ptp_out.seek(0)
            _refresh_with_xlwings(prod_out, ptp_out)
            logger.info("Pivot refresh completed.")
        else:
            status.info(
                "⏭️ Step 5/6: Skipping pivot refresh (manual mode)..."
            )
        progress.progress(85)

        # ---------------------------------------------------------- #
        # STEP 6: Encrypt & Store in session_state
        # ---------------------------------------------------------- #
        status.info("🔐 Step 6/6: Encrypting output files...")
        prod_out.seek(0)
        ptp_out.seek(0)
        prod_encrypted = encrypt_file(prod_out)
        ptp_encrypted  = encrypt_file(ptp_out)
        progress.progress(100)

        # Read as bytes — required for download_button persistence
        prod_data = prod_encrypted.read()
        ptp_data  = ptp_encrypted.read()

        if not prod_data:
            raise ValueError(
                "Productivity file is empty after encryption!"
            )
        if not ptp_data:
            raise ValueError(
                "PTP file is empty after encryption!"
            )

        # Store in session_state so downloads persist after reruns
        st.session_state.prod_encrypted  = prod_data
        st.session_state.ptp_encrypted   = ptp_data
        st.session_state.prod_filename   = get_output_filename(PROD_PREFIX)
        st.session_state.ptp_filename    = get_output_filename(PTP_PREFIX)
        st.session_state.automation_done  = True
        st.session_state.automation_error = None

        logger.info(
            f"Prod file: {len(prod_data)} bytes | "
            f"PTP file: {len(ptp_data)} bytes"
        )
        status.success(
            "✅ Automation complete! Scroll down to download your files."
        )
        st.balloons()

    except Exception as e:
        progress.empty()
        st.session_state.automation_done  = True
        st.session_state.automation_error = str(e)
        st.error(f"❌ Automation failed: {e}")
        logger.error(f"Automation error: {e}", exc_info=True)


def _refresh_with_xlwings(
    prod_bytes: io.BytesIO,
    ptp_bytes: io.BytesIO
):
    """
    Saves BytesIO to temp files, refreshes pivots via xlwings, reloads.
    Only called on Windows with Excel installed.
    """
    import tempfile
    from src.pivot import refresh_pivots

    with tempfile.NamedTemporaryFile(
        suffix=".xlsx", delete=False
    ) as tmp_prod:
        tmp_prod.write(prod_bytes.read())
        tmp_prod_path = tmp_prod.name

    with tempfile.NamedTemporaryFile(
        suffix=".xlsx", delete=False
    ) as tmp_ptp:
        tmp_ptp.write(ptp_bytes.read())
        tmp_ptp_path = tmp_ptp.name

    try:
        refresh_pivots(tmp_prod_path)
        refresh_pivots(tmp_ptp_path)

        with open(tmp_prod_path, "rb") as f:
            prod_bytes.seek(0)
            prod_bytes.write(f.read())
            prod_bytes.truncate()
            prod_bytes.seek(0)

        with open(tmp_ptp_path, "rb") as f:
            ptp_bytes.seek(0)
            ptp_bytes.write(f.read())
            ptp_bytes.truncate()
            ptp_bytes.seek(0)

    finally:
        if os.path.exists(tmp_prod_path):
            os.remove(tmp_prod_path)
        if os.path.exists(tmp_ptp_path):
            os.remove(tmp_ptp_path)