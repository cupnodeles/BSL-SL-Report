# src/chaos_key/__init__.py
# Backtick-key reporter as a dependency-free Streamlit custom component
# (vanilla JS speaks the component protocol — no npm build step).
#
# HONEST LIMITATION: browsers deliver key events only to the focused
# frame. This iframe is 0-height and unfocusable in practice, so presses
# are captured best-effort only (e.g. right after interacting with it,
# which rarely happens). The floating "`" button in the main UI is the
# guaranteed toggle; treat this component's signal as a bonus.
#
# Returns an ever-increasing press counter (0 = no press yet).

import os

import streamlit.components.v1 as components

_FRONTEND_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "frontend"
)

_chaos_key = components.declare_component("chaos_key", path=_FRONTEND_DIR)


def chaos_keypress(default: int = 0) -> int:
    """Renders the hidden key listener; returns total ` presses."""
    try:
        value = _chaos_key(default=default, height=0)
    except Exception:
        return default
    try:
        return int(value)
    except Exception:
        return default
