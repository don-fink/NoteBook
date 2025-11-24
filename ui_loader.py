"""
ui_loader.py
Centralized helpers for loading .ui files with PyQt5's uic module.

Benefits:
- Single place to resolve UI file paths (dev vs packaged)
- Consistent error handling and future portability (e.g., PyQt6/PySide6)
"""

import os
import sys
from typing import Optional

from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import Qt


def _base_dir() -> str:
    """Return the directory where UI files are located.

    In development this is the package directory. When packaged (PyInstaller one-folder),
    .ui files are collected alongside the application in the same directory as this module.
    If a frozen one-file variant is ever used, sys._MEIPASS can be added here.
    """
    try:
        # Prefer the module directory
        return os.path.dirname(__file__)
    except Exception:
        return os.path.abspath(os.getcwd())


def get_ui_path(name: str) -> str:
    """Resolve a UI filename to an absolute path.

    Tries the package directory first; falls back to current working directory.
    """
    base = _base_dir()
    p = os.path.join(base, name)
    if os.path.exists(p):
        return p
    try:
        # Fallback for atypical layouts
        alt = os.path.abspath(name)
        if os.path.exists(alt):
            return alt
    except Exception:
        pass
    # Return the intended primary path even if missing; caller may decide fallback
    return p


def load_ui(name: str, base_instance: Optional[QtWidgets.QWidget] = None):
    """Load a .ui by name and return the constructed widget/dialog.

    If base_instance is provided, it will be populated by uic.loadUi.
    """
    path = get_ui_path(name)
    return uic.loadUi(path, base_instance)


def load_dialog(name: str, parent: Optional[QtWidgets.QWidget] = None) -> QtWidgets.QDialog:
    """Construct a QDialog and populate it from the given .ui file name.

    Ensures normal window chrome and movement behavior while honoring parent modality.
    """
    path = get_ui_path(name)
    dlg = QtWidgets.QDialog(parent)
    uic.loadUi(path, dlg)
    try:
        dlg.setWindowFlag(Qt.Window, True)
    except Exception:
        pass
    return dlg


def load_settings_dialog(parent: Optional[QtWidgets.QWidget] = None) -> QtWidgets.QDialog:
    """Load and return the Settings dialog from settings_dialog.ui.

    Construct a QDialog with the given parent and populate it via uic.loadUi
    so it remains a movable, top-level dialog with normal window chrome.
    """
    dlg = load_dialog("settings_dialog.ui", parent)
    try:
        dlg.setWindowModality(Qt.ApplicationModal)
    except Exception:
        pass
    return dlg


def load_main_window():
    """Load the two-column main window UI."""
    base = _base_dir()
    ui_path = os.path.join(base, "main_window_2_column.ui")
    if not os.path.exists(ui_path):
        raise FileNotFoundError(f"Main window UI file not found: {ui_path}")
    return uic.loadUi(ui_path)
