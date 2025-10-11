"""
ui_sections.py
Handles logic for displaying sections (and now pages) in the left binder tree.
This module is used by other parts of the UI to populate the left panel tree.

Change: add pages under each section so the left tree shows
Binder -> Section -> Pages.
In two-column mode, page items are selectable and clicking them
loads the page content into the center editor.
"""

from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt

from db_pages import get_pages_by_section_id
from db_sections import get_sections_by_notebook_id
from page_editor import is_two_column_ui as _is_two_col
from page_editor import load_page as _load_page_2col
from ui_tabs import load_first_page_for_current_tab, select_tab_for_section


def add_sections_as_children(tree_widget, notebook_id, parent_item, db_path):
    """Populate the given binder item with its sections and pages.

    Creates one child per Section under the provided parent binder item, and
    for each section, adds its Pages as children of that section item.

    Roles used:
      - column 0, role 1000: id (section_id or page_id)
      - column 0, role 1001: kind ('section' or 'page')
    """
    sections = get_sections_by_notebook_id(notebook_id, db_path)
    for section in sections:
        # section: (id, notebook_id, title, ...)
        section_id = section[0]
        section_title = str(section[2])
        sec_item = QtWidgets.QTreeWidgetItem([section_title])
        sec_item.setData(0, 1000, section_id)  # Store section_id in UserRole
        try:
            sec_item.setData(0, 1001, "section")
        except Exception:
            pass
        # Sections: enabled + selectable + draggable; DO NOT accept drops
        # Reordering within a binder uses the binder as the drop target so the
        # drop indicator appears between section rows. Allowing drops directly
        # on a section causes Qt to treat it as a child-drop and expand instead
        # of reordering.
        try:
            flags = sec_item.flags()
            flags = (
                flags
                | Qt.ItemIsEnabled
                | Qt.ItemIsSelectable
                | Qt.ItemIsDragEnabled
            )
            sec_item.setFlags(flags)
        except Exception:
            pass
        parent_item.addChild(sec_item)

        # Add pages under this section
        try:
            pages = get_pages_by_section_id(section_id, db_path)
        except Exception:
            pages = []
        # Prefer order_index (index 6) then id (index 0) when available
        try:
            pages_sorted = sorted(pages, key=lambda p: (p[6], p[0]))
        except Exception:
            pages_sorted = pages
        for p in pages_sorted:
            page_id = p[0]
            page_title = str(p[2])
            page_item = QtWidgets.QTreeWidgetItem([page_title])
            page_item.setData(0, 1000, page_id)
            try:
                page_item.setData(0, 1001, "page")
                # Also store parent section id for convenience (1002 consistent with ui_tabs)
                page_item.setData(0, 1002, section_id)
            except Exception:
                pass
            # Two-column mode: pages are selectable and draggable for reordering within a section.
            # Legacy tab mode: pages are enabled but not selectable or draggable here.
            try:
                pflags = page_item.flags()
                if _is_two_col(tree_widget.window()):
                    pflags = (pflags | Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled) & ~(
                        Qt.ItemIsDropEnabled
                    )
                else:
                    pflags = (pflags | Qt.ItemIsEnabled) & ~(
                        Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled
                    )
                page_item.setFlags(pflags)
            except Exception:
                pass
            sec_item.addChild(page_item)


def on_notebook_clicked(item, column):
    tree_widget = item.treeWidget()
    # Remove all children from all top-level items
    for i in range(tree_widget.topLevelItemCount()):
        tree_widget.topLevelItem(i).takeChildren()
    # Find notebook id by item index (assuming order matches DB)
    notebook_id = item.data(0, 1000)  # Custom role for id
    if notebook_id is not None:
        add_sections_as_children(tree_widget, notebook_id, item)
    # Connect section click handler
    tree_widget.itemClicked.disconnect()
    tree_widget.itemClicked.connect(on_section_clicked)


def on_section_clicked(item, column, db_path):
    # Only handle if this is a section (not a notebook)
    if item.parent() is None:
        # It's a notebook, reconnect notebook handler
        tree_widget = item.treeWidget()
        tree_widget.itemClicked.disconnect()
        tree_widget.itemClicked.connect(lambda i, c: on_notebook_clicked(i, c, db_path))
        return
    kind = item.data(0, 1001)
    window = item.treeWidget().window()
    if kind == "section":
        section_id = item.data(0, 1000)
        if section_id is not None:
            select_tab_for_section(window, section_id)
            # In two-column mode, do not auto-load; legacy tabs will load via helper
            if not _is_two_col(window):
                load_first_page_for_current_tab(window)
    elif kind == "page" and _is_two_col(window):
        page_id = item.data(0, 1000)
        section_id = item.data(0, 1002)
        if page_id is not None and section_id is not None:
            try:
                # Update current context map
                if not hasattr(window, "_current_page_by_section"):
                    window._current_page_by_section = {}
                window._current_page_by_section[int(section_id)] = int(page_id)
                window._current_section_id = int(section_id)
            except Exception:
                pass
            try:
                _load_page_2col(window, int(page_id))
            except Exception:
                pass
            try:
                # Persist last state
                from settings_manager import set_last_state

                set_last_state(section_id=int(section_id), page_id=int(page_id))
            except Exception:
                pass
