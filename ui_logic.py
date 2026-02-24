"""
ui_logic.py
Contains logic for populating the notebook tree widget and handling notebook expand/collapse events.
"""

from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush

from db_access import get_notebooks


# Color constants for deleted items
DELETED_ITEM_COLOR = QColor(128, 128, 128)  # Grey


def populate_notebook_names(window, db_path):
    try:
        from settings_manager import get_show_deleted
        include_deleted = get_show_deleted()
    except Exception:
        include_deleted = False
    
    notebooks = get_notebooks(db_path, include_deleted=include_deleted)
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    tree_widget.clear()
    try:
        # Show expand/collapse arrows on top-level items
        tree_widget.setRootIsDecorated(True)
        tree_widget.setItemsExpandable(True)
        # Ensure tree is configured for internal DnD moves
        tree_widget.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        tree_widget.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        tree_widget.setDefaultDropAction(Qt.MoveAction)
        tree_widget.setAcceptDrops(True)
        tree_widget.setDragEnabled(True)
        tree_widget.setDropIndicatorShown(True)
    except Exception:
        pass
    for notebook in notebooks:
        # notebook[0] = id, notebook[1] = name, ..., notebook[5] = deleted_at
        item = QtWidgets.QTreeWidgetItem([str(notebook[1])])
        item.setData(0, 1000, notebook[0])  # Store notebook_id in UserRole
        
        # Check if this notebook is deleted (column 5 = deleted_at)
        is_deleted = False
        try:
            is_deleted = notebook[5] is not None
        except (IndexError, TypeError):
            pass
        
        # Store deleted status for context menu logic
        item.setData(0, 1003, is_deleted)  # 1003 = is_deleted flag
        
        # Grey out deleted items
        if is_deleted:
            item.setForeground(0, QBrush(DELETED_ITEM_COLOR))
        
        # Always show an expander so users know it can be expanded to sections
        try:
            item.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.ShowIndicator)
        except Exception:
            pass
        # Enable dragging binders and allow dropping onto a binder to reorder sections beneath it
        # But disable DnD for deleted items
        try:
            flags = item.flags()
            if is_deleted:
                flags = (flags | Qt.ItemIsEnabled | Qt.ItemIsSelectable) & ~(Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
            else:
                flags = (
                    flags
                    | Qt.ItemIsEnabled
                    | Qt.ItemIsSelectable
                    | Qt.ItemIsDragEnabled
                    | Qt.ItemIsDropEnabled
                )
            item.setFlags(flags)
        except Exception:
            pass
        tree_widget.addTopLevelItem(item)
        # Add a hidden placeholder child to ensure the expander arrow is always visible
        # without introducing visible spacing when collapsed.
        try:
            placeholder = QtWidgets.QTreeWidgetItem([""])
            placeholder.setDisabled(True)
            placeholder.setHidden(True)
            item.addChild(placeholder)
        except Exception:
            pass
    # Do not connect click handlers here; ui_tabs.setup_tab_sync manages clicks/expansion


def toggle_notebook_expand(item, column, db_path):
    # Deprecated: left-click handling moved to ui_tabs. Keep as no-op for compatibility.
    return
