"""
ui_logic.py
Contains logic for populating the notebook tree widget and handling notebook expand/collapse events.
"""
from PyQt5 import QtWidgets
from db_access import get_notebooks
from ui_sections import on_notebook_clicked

def populate_notebook_names(window, db_path):
    notebooks = get_notebooks(db_path)
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    tree_widget.clear()
    for notebook in notebooks:
        # notebook[0] = id, notebook[1] = name
        item = QtWidgets.QTreeWidgetItem([str(notebook[1])])
        item.setData(0, 1000, notebook[0])  # Store notebook_id in UserRole
        tree_widget.addTopLevelItem(item)
    tree_widget.itemClicked.connect(lambda item, column: toggle_notebook_expand(item, column, db_path))

def toggle_notebook_expand(item, column, db_path):
    # Only handle top-level (notebook) items
    if item.parent() is not None:
        from ui_sections import on_section_clicked
        on_section_clicked(item, column, db_path)
        return
    tree_widget = item.treeWidget()
    if item.childCount() == 0:
        from ui_sections import add_sections_as_children
        notebook_id = item.data(0, 1000)
        if notebook_id is not None:
            add_sections_as_children(tree_widget, notebook_id, item, db_path)
        item.setExpanded(True)
    else:
        if item.isExpanded():
            item.setExpanded(False)
        else:
            item.setExpanded(True)
