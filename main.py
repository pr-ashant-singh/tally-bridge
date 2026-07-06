"""
TallyBridge - Zerodha Tax P&L to Tally Converter

Main entry point for the application.
"""

import sys
import os

# Ensure the project root is in path (needed for PyInstaller)
if getattr(sys, 'frozen', False):
    # Running as PyInstaller bundle
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, base_path)

from src.gui import run_app


if __name__ == "__main__":
    run_app()
