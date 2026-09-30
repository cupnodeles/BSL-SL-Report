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
from src.theme import inject_hub_css
from src.chaos import is_chaos, toggle_chaos, inject_chaos_css, celebrate
from src.chaos_key import chaos_keypress
from src.dialer import extract_dialer_data
from src.drr import (
    clean_drr, split_drr, map_remedial, map_early, drr_latest_date,
)
from src.productivity import populate_productivity
from src.ptp import extract_ptp_rows, populate_ptp
from src.encrypt import encrypt_file, decrypt_file, is_encrypted

logger = setup_logger()

PROD_PREFIX  = "SPM Productivity & Penetration Report_BSL-early&remedial"
PTP_PREFIX   = "SPM PTP Monitoring Report_BSL-early"
IS_WINDOWS   = platform.system() == "Windows"

# Repo root (ui.py lives in src/) — for bundled assets like chaos music.
_APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CHAOS_MP3 = os.path.join(_APP_ROOT, "assets", "dwa.mp3")

# Bulk-upload slots + filename patterns (case-insensitive substring).
# Each slot's variants share one common prefix, so matching is
# order-independent.
UPLOAD_SLOTS = ("drr", "dialer", "prod_template", "ptp_template")
SLOT_LABELS = {
    "drr":           "📋 Daily Remark Report",
    "dialer":        "📊 Dialer Report (Optional)",
    "prod_template": "📄 Productivity Template",
    "ptp_template":  "📄 PTP Monitoring Template",
}
SLOT_PATTERNS = {
    # Daily_Remark_Report BSL... / Daily_Remark_Report...
    "drr":           ["daily_remark_report"],
    # DIALER REPORT BUSINESS LOAN SL ...
    "dialer":        ["dialer report"],
    # SPM Productivity & Penetration Report_BSL-early&remedial_... /
    # SPM Productivity & Penetration Report...
    "prod_template": ["spm productivity & penetration report"],
    # SPM PTP Monitoring Report_BSL-early_... /
    # SPM PTP Monitoring Report...
    "ptp_template":  ["spm ptp monitoring report"],
}


def classify_upload(filename: str):
    """
    Classifies an uploaded filename into one of UPLOAD_SLOTS.
    Case-insensitive substring match. Returns the slot key or None
    when the file is not recognized.
    """
    name = (filename or "").upper()
    for slot in UPLOAD_SLOTS:
        for pattern in SLOT_PATTERNS[slot]:
            if pattern.upper() in name:
                return slot
    return None


def classify_uploads(files):
    """
    Sorts uploaded files into slots. First file per slot wins;
    extras and unrecognized files are reported (not silently dropped).
    Returns (slots_dict, warnings) where slots_dict maps each slot to
    its file (or None) and warnings is a list of message strings.
    """
    slots = {slot: None for slot in UPLOAD_SLOTS}
    warnings = []
    for f in files or []:
        slot = classify_upload(getattr(f, "name", ""))
        if slot is None:
            warnings.append(f"⚠️ Unrecognized file ignored: {f.name}")
        elif slots[slot] is None:
            slots[slot] = f
        else:
            warnings.append(
                f"⚠️ Duplicate {SLOT_LABELS[slot]} ignored: {f.name} "
                f"(using {slots[slot].name})."
            )
    return slots, warnings


def _init_session_state():
    """Initialize all session state variables."""
    defaults = {
        "prod_encrypted":   None,
        "ptp_encrypted":    None,
        "prod_filename":    None,
        "ptp_filename":     None,
        "removed_rows":     None,
        "reset_info":       None,
        "pen_reset_info":   None,
        "automation_done":  False,
        "automation_error": None,
        "hub_theme": "light",
        "chaos": False,
        "chaos_dock": False,
        "chaos_key_last": 0,
        "celebrate_once": False,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def launch_app():
    st.set_page_config(
        page_title="BPI Automation SL — BSL Report",
        page_icon="▚",
        layout="wide"
    )

    # Initialize session state FIRST
    _init_session_state()

    # Hub theme (light default, night override) — CSS vars only, no reboot.
    st.markdown(
        inject_hub_css(st.session_state.get("hub_theme", "light")),
        unsafe_allow_html=True,
    )
    # Chaos overlay — session-only visual, silent, pipeline untouched.
    inject_chaos_css()

    # ---- Floating chaos dock (bottom-right) ---- #
    # "`" toggles the chaos button; the button toggles chaos mode.
    # The keydown component is best-effort (iframe focus); the "`"
    # button always works. Render order matters: each anchor div must
    # immediately precede its control for the dock CSS selectors.
    st.markdown('<div class="chaos-dock-tick"></div>',
                unsafe_allow_html=True)
    if st.button("`", key="btn_chaos_dock",
                 help="Show/hide the chaos button"):
        st.session_state.chaos_dock = not st.session_state.get(
            "chaos_dock", False
        )
        st.rerun()
    if st.session_state.get("chaos_dock", False):
        st.markdown('<div class="chaos-dock-btn"></div>',
                    unsafe_allow_html=True)
        chaos_label = "✕ Calm" if is_chaos() else "🌈 Chaos"
        if st.button(
            chaos_label,
            key="btn_chaos",
            help="Warning: flashing lights. Session-only visual.",
        ):
            toggle_chaos()
            if is_chaos():
                # One-shot celebration: survives the rerun below,
                # renders once, popped so later reruns stay clean.
                st.session_state.celebrate_once = True
            st.rerun()
    # One-shot overlay render (set by the handler above).
    if st.session_state.pop("celebrate_once", False):
        celebrate()
    try:
        _presses = chaos_keypress()
    except Exception as e:
        logger.warning(f"Chaos key component unavailable: {e}")
        _presses = 0
    if _presses != st.session_state.get("chaos_key_last", 0):
        st.session_state.chaos_key_last = _presses
        st.session_state.chaos_dock = not st.session_state.get(
            "chaos_dock", False
        )
        st.rerun()
    if is_chaos() and os.path.exists(_CHAOS_MP3):
        st.markdown('<div class="chaos-dock-audio"></div>',
                    unsafe_allow_html=True)
        # Autoplay works once the browser allows it (the Chaos click
        # counts as interaction); otherwise the visible player remains.
        # Guarded: older Streamlit without autoplay/loop falls back.
        try:
            st.audio(_CHAOS_MP3, format="audio/mp3",
                     autoplay=True, loop=True)
        except TypeError:
            logger.warning(
                "st.audio autoplay/loop unsupported here — plain player."
            )
            st.audio(_CHAOS_MP3, format="audio/mp3")

    # ---- Sidebar: brand + nav + theme + chaos (mirrors hub Sidebar) ---- #
    with st.sidebar:
        st.markdown(
            """
            <div class="px-panel">
              <div class="px-brand">BPI AUTOMATION</div>
              <div style="font-size:10px;text-transform:uppercase;letter-spacing:.15em;opacity:.75;">
                Hub v1 · Pixel — BSL SL
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption("> BSL SL · active")
        st.caption("> XDAYS SL · sibling app")
        theme_label = (
            "☀ Light mode" if st.session_state.hub_theme == "night"
            else "☾ Night mode"
        )
        if st.button(theme_label, key="btn_hub_theme", use_container_width=True):
            st.session_state.hub_theme = (
                "light" if st.session_state.hub_theme == "night" else "night"
            )
            st.rerun()
        st.divider()
        st.caption("Sys.ok · 8-bit UI · Streamlit Cloud safe")

    # ---- Hero (mirrors BslSlPage title + ①②③ intro) ---- #
    st.markdown(
        '<h1 class="px-title">▚ BSL SL AUTOMATION</h1>',
        unsafe_allow_html=True,
    )
    st.markdown(
        "Pick the card below, then follow the steps: "
        "**① Drop files → ② Check what was detected → ③ Run & download.**"
    )
    with st.container():
        st.markdown(
            """
            <div class="px-panel">
              <span class="badge-pixel">Selected</span>
              <strong>&nbsp; BSL Report</strong><br/>
              <span>Dialer + DRR → Productivity &amp; PTP Monitoring reports.
              Upload the Daily Remark Report, Dialer Report (optional), and the
              Productivity/PTP templates — download finished, pivot-ready workbooks.</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("---")

    # ------------------------------------------------------------------ #
    # FILE UPLOADS — all files at once, auto-classified by filename
    # ------------------------------------------------------------------ #
    st.markdown(
        '<h2><span class="step-badge">1</span>Drop files — select all input files at once (auto-sorted)</h2>',
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "📁 Select all input files at once",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Select all files together — they are classified automatically: "
             "Daily_Remark_Report* → Daily Remark Report, "
             "DIALER REPORT* → Dialer (optional), "
             "SPM Productivity & Penetration Report* → Productivity Template, "
             "SPM PTP Monitoring Report* → PTP Monitoring Template.",
        key="uploader_all",
    )

    slots, upload_warnings = classify_uploads(uploaded_files)
    dialer_file   = slots["dialer"]
    drr_file      = slots["drr"]
    prod_template = slots["prod_template"]
    ptp_template  = slots["ptp_template"]

    # Classification summary: one row per slot
    for slot in UPLOAD_SLOTS:
        f = slots[slot]
        if f is not None:
            st.success(f"✅ {SLOT_LABELS[slot]}: {f.name}")
        else:
            msg = f"⬜ {SLOT_LABELS[slot]}: not uploaded"
            if slot == "dialer":
                st.info(msg + " (optional — Penetration sheet stays as-is).")
            else:
                st.info(msg)
    for w in upload_warnings:
        st.warning(w)

    st.markdown("---")

    # ------------------------------------------------------------------ #
    # SETTINGS
    # ------------------------------------------------------------------ #
    st.markdown(
        '<h2><span class="step-badge">2</span>Check — every slot must show ✅ before you can run</h2>',
        unsafe_allow_html=True,
    )

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

    new_month_run = st.toggle(
        "🗓 New Month Run",
        value=False,
        help="ON = wipe all existing PTP List AND Penetration records "
             "before pasting. OFF = never wipe — every run appends, "
             "even across months.",
        key="toggle_newmonth",
    )
    if new_month_run:
        st.info(
            "🗓 New Month Run is ON — existing PTP and Penetration "
            "records will be cleared and replaced."
        )

    st.markdown("---")

    # ------------------------------------------------------------------ #
    # RUN BUTTON — Dialer is OPTIONAL, other 3 files are required
    # ------------------------------------------------------------------ #
    st.markdown(
        '<h2><span class="step-badge">3</span>Run &amp; download</h2>',
        unsafe_allow_html=True,
    )
    required_uploaded = all([drr_file, prod_template, ptp_template])

    if not required_uploaded:
        st.info("👆 Please upload Daily Remark Report + both templates to proceed.")

    if dialer_file is None and required_uploaded:
        st.warning(
            "⚠️ No Dialer Report uploaded — 'Penetration Per Day' "
            "will be left as-is. PTP extraction only uses Early/Remedial data."
        )

    if st.button(
        "▶ RUN AUTOMATION",
        disabled=not required_uploaded,
        type="primary",
        key="btn_run"
    ):
        # Reset outputs before new run
        st.session_state.prod_encrypted  = None
        st.session_state.ptp_encrypted   = None
        st.session_state.prod_filename   = None
        st.session_state.ptp_filename    = None
        st.session_state.removed_rows    = None
        st.session_state.reset_info      = None
        st.session_state.pen_reset_info  = None
        st.session_state.automation_done  = False
        st.session_state.automation_error = None

        run_automation(
            dialer_file, drr_file,
            prod_template, ptp_template,
            auto_refresh, new_month_run,
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

            # ------------------------------------------------------ #
            # Removed PTP rows report — persists via session_state
            # ------------------------------------------------------ #
            removed = st.session_state.removed_rows
            if removed is not None and len(removed):
                st.markdown("---")
                st.subheader("⚠️ Removed PTP Rows")
                st.warning(
                    f"{len(removed)} PTP row(s) were excluded from the "
                    f"PTP Monitoring Report (blank PTP Date + zero "
                    f"PTP Amount)."
                )
                show_cols = [
                    c for c in [
                        "Source Row #", "Name", "LAN", "Status",
                        "PTP Date", "PTP Amount", "Reason",
                    ]
                    if c in removed.columns
                ]
                st.dataframe(
                    removed[show_cols] if show_cols else removed,
                    use_container_width=True,
                    key="tbl_removed",
                )
            elif removed is not None:
                st.success("✅ No PTP rows removed — all rows valid.")

            # ------------------------------------------------------ #
            # Monthly reset report — persists via session_state
            # ------------------------------------------------------ #
            reset = st.session_state.reset_info
            if reset is not None and reset.get("reset"):
                st.warning(
                    f"🗓 Monthly reset (manual New Month Run): cleared "
                    f"{reset.get('cleared', 0)} existing PTP record(s) "
                    f"before pasting."
                )
            elif reset is not None:
                st.info(
                    f"📅 Appended — kept "
                    f"{reset.get('existing_rows', 0)} existing PTP "
                    f"record(s)."
                )

            # ------------------------------------------------------ #
            # Penetration monthly reset report — via session_state
            # ------------------------------------------------------ #
            pen_reset = st.session_state.pen_reset_info
            if pen_reset is not None and pen_reset.get("reset"):
                st.warning(
                    f"🗓 Penetration monthly reset "
                    f"(manual New Month Run): cleared "
                    f"{pen_reset.get('cleared', 0)} existing record(s) "
                    f"before pasting."
                )
            elif pen_reset is not None and pen_reset.get("mode") == (
                "skipped-no-dialer"
            ):
                st.info("📅 Penetration skipped — no Dialer Report.")
            elif pen_reset is not None:
                st.info(
                    f"📅 Appended — kept "
                    f"{pen_reset.get('existing_rows', 0)} existing "
                    f"Penetration record(s)."
                )


def run_automation(
    dialer_file, drr_file,
    prod_template, ptp_template,
    auto_refresh: bool,
    new_month: bool = False,
):
    """Main automation runner with progress tracking."""

    progress = st.progress(0)
    status   = st.empty()

    try:
        # ---------------------------------------------------------- #
        # STEP 1: Extract Dialer Data [1] — OPTIONAL
        # PTP flow does not use dialer; only Penetration sheet does.
        # ---------------------------------------------------------- #
        if dialer_file is not None:
            status.info("📊 Step 1/6: Extracting Dialer Report data...")
            dialer_bytes = io.BytesIO(dialer_file.read())
            if is_encrypted(dialer_bytes):
                logger.info("Dialer file encrypted. Decrypting...")
                dialer_bytes = decrypt_file(dialer_bytes)
            dialer_df = extract_dialer_data(dialer_bytes)
            logger.info(f"Dialer rows extracted: {len(dialer_df)}")
        else:
            status.info(
                "⏭️ Step 1/6: No Dialer Report — "
                "skipping Penetration update..."
            )
            logger.warning(
                "Dialer file not provided — "
                "'Penetration Per Day' will be left untouched."
            )
            dialer_df = None
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
        # Filename date = latest DRR Date (falls back to yesterday).
        try:
            report_date = drr_latest_date(clean_df)
        except Exception as e:
            logger.warning(f"Could not derive report date: {e}")
            report_date = None
        if report_date is not None:
            status.info(
                f"📅 Report date from DRR: "
                f"{report_date.strftime('%m/%d/%Y')}"
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
        prod_out, pen_reset_info = populate_productivity(
            prod_bytes, dialer_df, early_df, remedial_df,
            new_month=new_month,
        )
        if pen_reset_info.get("reset"):
            status.warning(
                f"🗓 Penetration monthly reset (manual New Month Run): "
                f"cleared {pen_reset_info.get('cleared', 0)} "
                f"existing record(s)."
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

        # CRITICAL: Reset prod_out position before passing to extract_ptp_rows
        prod_out.seek(0)
        ptp_df, removed_df = extract_ptp_rows(prod_out)
        if len(removed_df):
            status.warning(
                f"⚠️ {len(removed_df)} PTP row(s) removed "
                f"(blank PTP Date + zero PTP Amount). "
                f"See details below."
            )
            logger.warning(
                f"Removed PTP rows: {len(removed_df)} "
                f"(blank PTP Date + zero PTP Amount)."
            )
        else:
            logger.info("Removed PTP rows: 0.")

        # CRITICAL: Reset again before populating productivity output
        prod_out.seek(0)
        ptp_out, reset_info = populate_ptp(
            ptp_bytes, ptp_df, new_month=new_month
        )
        if reset_info.get("reset"):
            status.warning(
                f"🗓 Monthly reset (manual New Month Run): cleared "
                f"{reset_info.get('cleared', 0)} existing PTP record(s)."
            )
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
        st.session_state.prod_filename   = get_output_filename(
            PROD_PREFIX, report_date
        )
        st.session_state.ptp_filename    = get_output_filename(
            PTP_PREFIX, report_date
        )
        st.session_state.removed_rows    = removed_df
        st.session_state.reset_info      = reset_info
        st.session_state.pen_reset_info  = pen_reset_info
        st.session_state.automation_done  = True
        st.session_state.automation_error = None

        logger.info(
            f"Prod file: {len(prod_data)} bytes | "
            f"PTP file: {len(ptp_data)} bytes"
        )
        status.success(
            "✅ Automation complete! Scroll down to download your files."
        )
        celebrate()

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