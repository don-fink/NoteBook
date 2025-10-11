"""
ui_loader.py
Loads the main window from the .ui file using PyQt5's uic module.
"""
from PyQt5 import uic, QtWidgets
import os

def load_main_window():
    """Load the main window UI.
    Prefer the new two-column layout (main_window_2_column.ui) if present,
    otherwise fall back to the legacy tabbed layout (main_window.ui).
    """
    base = os.path.dirname(__file__)
    two_col = os.path.join(base, 'main_window_2_column.ui')
    legacy = os.path.join(base, 'main_window.ui')
    ui_path = two_col if os.path.exists(two_col) else legacy
    return uic.loadUi(ui_path)
