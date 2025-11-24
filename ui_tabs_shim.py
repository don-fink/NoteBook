"""
ui_tabs_shim.py
Compatibility shim that forwards to two_pane_core.
Kept for backward compatibility with existing imports.
"""

from two_pane_core import (
    setup_two_pane as setup_tab_sync,
    load_first_page as load_first_page_for_current_tab,
    save_current_page,
    select_section as select_tab_for_section,
    restore_last_position,
)
