# src/theme.py
# BPI Automation Hub theme port — Cloud-safe CSS only (no JS).
# Tokens mirrored 1:1 from bpi-automation-sl/src/index.css.
# Light is the default (matches .streamlit/config.toml); night is a
# runtime CSS-var override driven by st.session_state (no reboot).

LIGHT_VARS = {
    "--app-bg": "#e8e4d8",
    "--panel-bg": "#f4f1e6",
    "--card-bg": "#faf8ef",
    "--border-color": "#a89e83",
    "--accent": "#166b44",
    "--accent-ink": "#ffffff",
    "--accent-dim": "#cfe3d4",
    "--text-main": "#23241f",
    "--text-dim": "#4f4e43",
    "--text-faint": "#6f6e61",
    "--edge": "#23241f",
}

NIGHT_VARS = {
    "--app-bg": "#0f1420",
    "--panel-bg": "#161d2e",
    "--card-bg": "#1a2338",
    "--border-color": "#2b3857",
    "--accent": "#57e389",
    "--accent-ink": "#0b101d",
    "--accent-dim": "#244b38",
    "--text-main": "#e4eafb",
    "--text-dim": "#a7b4d4",
    "--text-faint": "#7e8cab",
    "--edge": "#0a0e18",
}


def inject_hub_css(mode: str = "light") -> str:
    """Returns the full hub stylesheet for the given mode."""
    palette = NIGHT_VARS if mode == "night" else LIGHT_VARS
    var_block = "\n".join(f"  {k}: {v};" for k, v in palette.items())
    return f"""
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Silkscreen:wght@400;700&display=swap');
<style>
:root {{
{var_block}
  --px: 4px;
  --font-pixel: "Silkscreen", ui-monospace, "Cascadia Mono", Menlo, Consolas, monospace;
  --font-sans: "Inter", ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
}}
[data-testid="stAppViewContainer"] {{
  background-color: var(--app-bg);
  background-image:
    repeating-linear-gradient(0deg, rgba(127,127,127,0.08) 0 2px, transparent 2px 4px),
    repeating-linear-gradient(90deg, rgba(127,127,127,0.06) 0 2px, transparent 2px 4px),
    linear-gradient(180deg, rgba(87,227,137,0.06) 0 120px, transparent 320px);
  color: var(--text-main);
  font-family: var(--font-sans);
}}
[data-testid="stSidebar"] {{
  background-color: var(--panel-bg);
  border-right: 4px solid var(--edge);
}}
[data-testid="stSidebar"] * {{
  color: var(--text-main);
  font-family: var(--font-sans);
}}
.px-panel {{
  background: var(--panel-bg);
  border: 4px solid var(--edge);
  box-shadow: 0 -4px 0 0 var(--border-color), 0 4px 0 0 var(--border-color),
    -4px 0 0 0 var(--border-color), 4px 0 0 0 var(--border-color);
  clip-path: polygon(0 12px, 12px 12px, 12px 0, calc(100% - 12px) 0,
    calc(100% - 12px) 12px, 100% 12px, 100% calc(100% - 12px),
    calc(100% - 12px) calc(100% - 12px), calc(100% - 12px) 100%,
    12px 100%, 12px calc(100% - 12px), 0 calc(100% - 12px));
  padding: 0.9rem 1rem;
}}
.px-title {{
  font-family: var(--font-pixel) !important;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  text-shadow: 2px 2px 0 var(--edge);
  color: var(--text-main);
}}
.px-brand {{
  font-family: var(--font-pixel) !important;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  color: var(--accent);
  font-weight: 700;
}}
.step-badge {{
  display: inline-flex; align-items: center; justify-content: center;
  width: 1.4rem; height: 1.4rem; border-radius: 9999px;
  background: var(--accent); color: var(--accent-ink);
  font-weight: 800; font-size: 0.8rem; margin-right: 0.5rem;
}}
.badge-pixel {{
  display: inline-block; font-family: var(--font-pixel);
  font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.08em; background: var(--accent); color: var(--accent-ink);
  padding: 0.25rem 0.625rem; border: 2px solid var(--edge);
}}
.status-ok {{ color: #1a7a3c; font-weight: 600; }}
.status-warn {{ color: #92580a; font-weight: 600; }}
.status-err {{ color: #b3261e; font-weight: 600; }}
[data-testid="stAppViewContainer"] .status-ok {{ color: #1a7a3c; }}
section[data-testid="stFileUploader"] {{
  border: 2px dashed var(--accent) !important;
  border-radius: 0.75rem;
  padding: 1rem;
  background: var(--card-bg);
}}
.stButton > button[kind="primary"], .stDownloadButton > button {{
  background: var(--accent) !important; color: var(--accent-ink) !important;
  font-weight: 700 !important; border: 3px solid var(--edge) !important;
  clip-path: polygon(8px 0, 100% 0, 100% calc(100% - 8px),
    calc(100% - 8px) 100%, 0 100%, 0 8px);
}}
.stButton > button[kind="secondary"] {{
  background: var(--card-bg) !important; color: var(--text-main) !important;
  font-weight: 700 !important; border: 3px solid var(--edge) !important;
  clip-path: polygon(8px 0, 100% 0, 100% calc(100% - 8px),
    calc(100% - 8px) 100%, 0 100%, 0 8px);
}}
.stSubheader {{ color: var(--text-main); font-weight: 700; }}
.stAlert, [data-testid="stStatusWidget"] {{ font-size: 0.95rem; }}
.block-container {{ max-width: 64rem; }}
code {{ color: var(--text-main); }}
</style>
"""
