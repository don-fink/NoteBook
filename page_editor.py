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
    _load_first_page_two_column as load_first_page_two_column,  # noqa: F401
    _load_page_two_column as load_page_two_column,  # noqa: F401
    _save_title_two_column as save_title_two_column,  # noqa: F401
    _set_page_edit_html as set_page_edit_html,  # noqa: F401
    save_current_page_two_column,  # noqa: F401
    save_current_page as _save_current_page_generic,
)


def save_current_page(window) -> None:
    """Public save that respects current UI mode (two‑pane vs legacy).

    Currently forwards to ui_tabs.save_current_page which handles both modes.
    """

    _save_current_page_generic(window)


def load_page(window, page_id: Optional[int] = None, html: Optional[str] = None) -> None:
    """Load a page into the editor in two‑pane mode. If page_id is None, clears the editor.

    This is a convenience wrapper matching the future API we want to expose.
    """

    load_page_two_column(window, page_id=page_id, html=html)
