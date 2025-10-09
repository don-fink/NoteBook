"""
ui_logic.py
Contains logic for populating the notebook tree widget and handling notebook expand/collapse events.
"""
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from db_access import get_notebooks

def populate_notebook_names(window, db_path):
    notebooks = get_notebooks(db_path)
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    tree_widget.clear()
    try:
        # Show expand/collapse arrows on top-level items
        tree_widget.setRootIsDecorated(True)
        tree_widget.setItemsExpandable(True)
    except Exception:
        pass
    for notebook in notebooks:
        # notebook[0] = id, notebook[1] = name
        item = QtWidgets.QTreeWidgetItem([str(notebook[1])])
        item.setData(0, 1000, notebook[0])  # Store notebook_id in UserRole
        # Always show an expander so users know it can be expanded to sections
        try:
            item.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.ShowIndicator)
        except Exception:
            pass
        # Enable dragging binders but do not allow dropping onto a binder item (prevents nesting)
        try:
            flags = item.flags()
            flags = (flags | Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled) & ~Qt.ItemIsDropEnabled
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
