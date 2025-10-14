"""
page_editor.py

Thin wrapper around the existing two‑pane editor functions so we can gradually
extract logic out of ui_tabs.py without changing behavior. For now, these
functions delegate to ui_tabs and keep names stable for future imports.

Contract (inputs/outputs):
- window: QMainWindow with child widgets pageEdit (QTextEdit) and pageTitleEdit (QLineEdit)
- page_id: int or None
- html: str or None

All functions must be safe to call when widgets are missing; they no‑op.
"""

from typing import Optional

# Delegate to existing implementation in ui_tabs
from ui_tabs import (
    _cancel_autosave as cancel_autosave,  # noqa: F401
    _is_two_column_ui as is_two_column_ui,  # noqa: F401
    _load_first_page_two_column as _ui_load_first_page_two_column,  # internal alias
    _load_page_two_column as _ui_load_page_two_column,  # noqa: F401
    _save_title_two_column as save_title_two_column,  # noqa: F401
    _set_page_edit_html as _ui_set_page_edit_html,  # noqa: F401
    save_current_page_two_column,  # noqa: F401
    save_current_page as _save_current_page_generic,
)

from PyQt5 import QtWidgets
from PyQt5.QtCore import QUrl
import os
try:
    from ui_planning_register import refresh_planning_register_styles
except Exception:  # pragma: no cover - optional import guard
    def refresh_planning_register_styles(_te):
        return


def save_current_page(window) -> None:
    """Public save that respects current UI mode (two‑pane vs legacy).

    Currently forwards to ui_tabs.save_current_page which handles both modes.
    """

    _save_current_page_generic(window)


def load_page(window, page_id: Optional[int] = None, html: Optional[str] = None) -> None:
    """Load a page into the editor in two‑pane mode. If page_id is None, clears the editor.

    After loading, re-apply Planning Register styles (header/totals shading, numeric alignment).
    """
    # Ensure relative media paths will resolve during the ensuing HTML load
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            media_root = getattr(window, "_media_root", None)
            if isinstance(media_root, str) and media_root:
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                te.document().setBaseUrl(QUrl.fromLocalFile(base))
    except Exception:
        pass
    _ui_load_page_two_column(window, page_id=page_id, html=html)
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            # Ensure relative media paths resolve against the DB media root
            media_root = getattr(window, "_media_root", None)
            if isinstance(media_root, str) and media_root:
                # Trailing separator ensures base url is treated as a folder
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                try:
                    te.document().setBaseUrl(QUrl.fromLocalFile(base))
                except Exception:
                    pass
            refresh_planning_register_styles(te)
    except Exception:
        pass


def set_page_edit_html(window, html: Optional[str]) -> None:
    """Set the editor HTML and apply Planning Register styles immediately after.

    This function replaces direct calls into ui_tabs to avoid modifying ui_tabs further.
    """
    # Set baseUrl before and after applying HTML to ensure resolution
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            media_root = getattr(window, "_media_root", None)
            if isinstance(media_root, str) and media_root:
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                te.document().setBaseUrl(QUrl.fromLocalFile(base))
    except Exception:
        pass
    _ui_set_page_edit_html(window, html)
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            media_root = getattr(window, "_media_root", None)
            if isinstance(media_root, str) and media_root:
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                te.document().setBaseUrl(QUrl.fromLocalFile(base))
            refresh_planning_register_styles(te)
    except Exception:
        pass


def load_first_page_two_column(window) -> None:
    """Load the first page for the current section in two‑pane mode and restyle PR tables."""
    # Ensure base is ready before first page load
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            media_root = getattr(window, "_media_root", None)
            if isinstance(media_root, str) and media_root:
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                te.document().setBaseUrl(QUrl.fromLocalFile(base))
    except Exception:
        pass
    _ui_load_first_page_two_column(window)
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            media_root = getattr(window, "_media_root", None)
            if isinstance(media_root, str) and media_root:
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                try:
                    te.document().setBaseUrl(QUrl.fromLocalFile(base))
                except Exception:
                    pass
            refresh_planning_register_styles(te)
    except Exception:
        pass
