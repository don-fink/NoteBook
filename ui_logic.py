"""
ui_logic.py
Contains logic for populating the notebook tree widget and handling notebook expand/collapse events.
"""
from PyQt5 import QtWidgets
from db_access import get_notebooks

def populate_notebook_names(window, db_path):
    notebooks = get_notebooks(db_path)
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    tree_widget.clear()
    for notebook in notebooks:
        # notebook[0] = id, notebook[1] = name
        item = QtWidgets.QTreeWidgetItem([str(notebook[1])])
        item.setData(0, 1000, notebook[0])  # Store notebook_id in UserRole
        tree_widget.addTopLevelItem(item)
    # Do not connect click handlers here; ui_tabs.setup_tab_sync manages clicks/expansion

def toggle_notebook_expand(item, column, db_path):
    # Deprecated: left-click handling moved to ui_tabs. Keep as no-op for compatibility.
    return
