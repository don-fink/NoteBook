"""
main.py
Entry point for the NoteBook application. Handles main window setup, menu actions, database creation/opening, and application startup.
"""
import sys
from PyQt5 import QtWidgets
from ui_loader import load_main_window
from ui_logic import populate_notebook_names
from settings_manager import get_last_db, set_last_db

def create_new_database(window):
    options = QtWidgets.QFileDialog.Options()
    file_name, _ = QtWidgets.QFileDialog.getSaveFileName(window, "Create New Database", "", "SQLite DB Files (*.db);;All Files (*)", options=options)
    if file_name:
        import sqlite3
        conn = sqlite3.connect(file_name)
        cursor = conn.cursor()
        cursor.executescript('''
            CREATE TABLE IF NOT EXISTS notebooks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                modified_at TEXT NOT NULL DEFAULT (datetime('now')),
                order_index INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS sections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                notebook_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                modified_at TEXT NOT NULL DEFAULT (datetime('now')),
                order_index INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (notebook_id) REFERENCES notebooks(id)
            );
            CREATE TABLE IF NOT EXISTS pages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                section_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                content_html TEXT,
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                modified_at TEXT NOT NULL DEFAULT (datetime('now')),
                order_index INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (section_id) REFERENCES sections(id)
            );
        ''')
        conn.commit()
        # Set version to 1
        cursor.execute('PRAGMA user_version = 1')
        conn.commit()
        conn.close()
        set_last_db(file_name)
        # Clear widgets
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
        if tree_widget:
            tree_widget.clear()
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if tab_widget:
            tab_widget.clear()
        populate_notebook_names(window, file_name)

def migrate_database_if_needed(db_path):
    from db_version import get_db_version, set_db_version
    current_version = get_db_version(db_path)
    target_version = 1  # Update as needed for future migrations
    if current_version < target_version:
        # Placeholder for migration logic
        # Example: if current_version == 0: ...
        set_db_version(target_version, db_path)

def open_database(window):
    options = QtWidgets.QFileDialog.Options()
    file_name, _ = QtWidgets.QFileDialog.getOpenFileName(window, "Open Database", "", "SQLite DB Files (*.db);;All Files (*)", options=options)
    if file_name:
        migrate_database_if_needed(file_name)
        set_last_db(file_name)
        # Clear widgets
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
        if tree_widget:
            tree_widget.clear()
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if tab_widget:
            tab_widget.clear()
        populate_notebook_names(window, file_name)

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = load_main_window()
    db_path = get_last_db() or 'notes.db'
    populate_notebook_names(window, db_path)

    # Connect menu actions
    action_newdb = window.findChild(QtWidgets.QAction, 'actionNewDB')
    if action_newdb:
        action_newdb.triggered.connect(lambda: create_new_database(window))
    action_open = window.findChild(QtWidgets.QAction, 'actionOpen')
    if action_open:
        action_open.triggered.connect(lambda: open_database(window))

    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()