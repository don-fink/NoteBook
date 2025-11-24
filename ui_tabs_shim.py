"""
ui_tabs_shim.py
Transitional shim to gradually move away from ui_tabs.py.
Routes two-pane calls into two_pane_core, falls back to ui_tabs for legacy/tabbed paths.
"""

from two_pane_core import (
    is_two_column_ui,
    setup_two_pane,
    load_first_page,
    load_page as load_page_two_pane,
    save_current_page as save_current_page_two_pane,
    select_section as select_section_two_pane,
)


def setup_tab_sync(window):
    if is_two_column_ui(window):
        return setup_two_pane(window)
    # Legacy fallback
    from ui_tabs import setup_tab_sync as _legacy_setup

    return _legacy_setup(window)


def restore_last_position(window):
    # Delegates to existing implementation (works for both modes)
    from ui_tabs import restore_last_position as _restore

    return _restore(window)


def load_first_page_for_current_tab(window):
    if is_two_column_ui(window):
        return load_first_page(window)
    from ui_tabs import load_first_page_for_current_tab as _legacy_load

    return _legacy_load(window)


def save_current_page(window):
    if is_two_column_ui(window):
        return save_current_page_two_pane(window)
    from ui_tabs import save_current_page as _legacy_save

    return _legacy_save(window)


def select_tab_for_section(window, section_id):
    if is_two_column_ui(window):
        return select_section_two_pane(window, section_id)
    from ui_tabs import select_tab_for_section as _legacy_select

    return _legacy_select(window, section_id)
