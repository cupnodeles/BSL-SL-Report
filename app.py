# app.py
# BSL SL Automation — Main Entry Point
# Run locally:  streamlit run app.py
# Deploy:       push to GitHub → connect to share.streamlit.io

import sys
import os

# Ensure src/ is importable
sys.path.insert(0, os.path.dirname(__file__))

from src.ui import launch_app

if __name__ == "__main__":
    launch_app()