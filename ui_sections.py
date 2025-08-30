
"""
ui_sections.py
Handles logic for displaying sections and pages in the UI, including populating sections as children and handling section clicks.
"""
from PyQt5 import QtWidgets
from db_sections import get_sections_by_notebook_id
from db_pages import get_pages_by_section_id

def add_sections_as_children(tree_widget, notebook_id, parent_item, db_path):
    sections = get_sections_by_notebook_id(notebook_id, db_path)
    for section in sections:
        # section[2] is the title field
        child = QtWidgets.QTreeWidgetItem([str(section[2])])
        child.setData(0, 1000, section[0])  # Store section_id in UserRole
        parent_item.addChild(child)

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
    section_id = item.data(0, 1000)
    window = item.treeWidget().window()
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    tab_widget.clear()
    if section_id is not None:
        pages = get_pages_by_section_id(section_id, db_path)
        for page in pages:
            # page[2] = title, page[3] = content_html
            tab = QtWidgets.QWidget()
            layout = QtWidgets.QVBoxLayout(tab)
            text_edit = QtWidgets.QTextEdit(tab)
            text_edit.setObjectName('textEdit')
            text_edit.setHtml(page[3])
            layout.addWidget(text_edit)
            tab_widget.addTab(tab, page[2])
