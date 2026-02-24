"""
main.py
Entry point for the NoteBook application. Handles main window setup, menu actions, database creation/opening, and application startup.
"""
import os

import sys
import warnings

from PyQt5 import QtWidgets, uic
from ui_toast import show_toast
from PyQt5.QtCore import QProcess, Qt, QTimer, QUrl

from db_access import create_notebook as db_create_notebook
from db_access import delete_notebook as db_delete_notebook
from db_access import rename_notebook as db_rename_notebook
from db_pages import create_page as db_create_page
from db_pages import delete_page as db_delete_page
from db_pages import set_pages_order as db_set_pages_order
from db_pages import update_page_title as db_update_page_title
from db_sections import create_section as db_create_section
from db_sections import delete_section as db_delete_section
from db_sections import get_sections_by_notebook_id as db_get_sections_by_notebook_id
from db_sections import move_section_down as db_move_section_down
from db_sections import move_section_up as db_move_section_up
from db_sections import rename_section as db_rename_section
from settings_manager import (
    clear_last_state,
    get_last_db,
    get_window_geometry,
    get_window_maximized,
    set_last_db,
    set_last_state,
    set_window_geometry,
    set_window_maximized,
)
from ui_loader import load_main_window
from ui_logic import populate_notebook_names
from left_tree import ensure_left_tree_sections, refresh_for_notebook
from two_pane_core import restore_last_position, setup_two_pane
from left_tree import select_left_tree_page, update_left_tree_page_title
from page_editor import (
    is_two_column_ui as _is_two_column_ui,
    load_page as _load_page_two_column,
    load_first_page_two_column as _load_first_page_two_column,
    save_current_page,
)
from ui_planning_register import insert_planning_register
from ui_richtext import insert_table_from_preset
from ui_richtext import install_image_support
from ui_planning_register import ensure_planning_register_watcher


def _install_global_excepthook():
    """Install a sys.excepthook that shows a critical dialog and prints the traceback.

    This helps diagnose unexpected crashes that may otherwise close the app silently.
    """
    import sys as _sys
    import traceback as _traceback

    def _handler(exctype, value, tb):
        msg = "".join(_traceback.format_exception(exctype, value, tb))
        # Write to a local crash log for diagnostics
        try:
            import os as _os
            log_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "crash.log")
            with open(log_path, "a", encoding="utf-8") as _f:
                _f.write("\n=== Unhandled exception ===\n")
                _f.write(msg)
        except Exception:
            pass
        try:
            if QtWidgets.QApplication.instance() is not None:
                QtWidgets.QMessageBox.critical(None, "Unexpected Error", msg)
        except Exception:
            pass
        try:
            print(msg)
        except Exception:
            pass

    try:
        _sys.excepthook = _handler
    except Exception:
        pass

# --- Database initialization / migration helpers (reintroduced after rollback) ---
def create_new_database_file(db_path: str):
    """Create a brand new SQLite database at db_path using schema.sql and set initial version.

    Safe to call if the file already exists (will no-op)."""
    import os, sqlite3
    if os.path.isfile(db_path):
        # Do not overwrite existing DB; assume it's valid or will be initialized later.
        return
    schema_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql")
    schema_sql = ""
    try:
        with open(schema_file, "r", encoding="utf-8") as f:
            schema_sql = f.read()
    except Exception:
        pass
    conn = sqlite3.connect(db_path)
    try:
        if schema_sql.strip():
            conn.executescript(schema_sql)
        # Set initial user_version
        try:
            conn.execute("PRAGMA user_version = 5")
        except Exception:
            pass
        conn.commit()
    finally:
        conn.close()

def ensure_database_initialized(db_path: str):
    """Ensure required tables exist; if missing, apply schema.

    Idempotent: running on an existing, fully initialized DB is safe."""
    import os, sqlite3
    # If file missing, create directly
    if not os.path.isfile(db_path):
        create_new_database_file(db_path)
        return
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='notebooks'")
        row = cur.fetchone()
        if not row:
            # Tables absent -> apply schema
            schema_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql")
            try:
                with open(schema_file, "r", encoding="utf-8") as f:
                    conn.executescript(f.read())
                try:
                    conn.execute("PRAGMA user_version = 5")
                except Exception:
                    pass
                conn.commit()
            except Exception:
                pass
    finally:
        conn.close()

def migrate_database_if_needed(db_path: str, parent_window=None) -> bool:
    """Apply in-place migrations based on PRAGMA user_version.

    Current target version = 5. Returns True if migration succeeded or wasn't needed,
    False if user cancelled/chose to load a different database.
    
    Version history:
    - Version 4: Base schema (notebooks, sections, pages with order_index, parent_page_id)
    - Version 5: Added deleted_at column to notebooks, sections, pages for soft-delete
    """
    import sqlite3
    TARGET = 5
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA user_version")
        version = cur.fetchone()[0]
        
        if version >= TARGET:
            return True  # Already up to date
        
        # Check if deleted_at column exists (handles case where version wasn't bumped)
        def column_exists(table, column):
            cur.execute(f"PRAGMA table_info({table})")
            columns = [row[1] for row in cur.fetchall()]
            return column in columns
        
        needs_soft_delete = (
            not column_exists('notebooks', 'deleted_at') or
            not column_exists('sections', 'deleted_at') or
            not column_exists('pages', 'deleted_at')
        )
        
        if needs_soft_delete:
            # Prompt user before upgrading
            if parent_window is not None:
                msg_box = QtWidgets.QMessageBox(parent_window)
                msg_box.setWindowTitle("Database Upgrade Required")
                msg_box.setIcon(QtWidgets.QMessageBox.Question)
                msg_box.setText(
                    f"The database '{os.path.basename(db_path)}' uses an older schema.\n\n"
                    "This version of NoteBook requires a schema upgrade to support "
                    "the new soft-delete/restore feature.\n\n"
                    "It's recommended to back up your database before proceeding.\n\n"
                    "Would you like to upgrade now?"
                )
                upgrade_btn = msg_box.addButton("Upgrade", QtWidgets.QMessageBox.AcceptRole)
                different_btn = msg_box.addButton("Open Different Database", QtWidgets.QMessageBox.RejectRole)
                cancel_btn = msg_box.addButton(QtWidgets.QMessageBox.Cancel)
                msg_box.exec_()
                clicked = msg_box.clickedButton()
                
                if clicked == cancel_btn or clicked == different_btn:
                    conn.close()
                    return False  # Caller should handle - either exit or prompt for different DB
            
            # Apply migration: add deleted_at columns
            try:
                if not column_exists('notebooks', 'deleted_at'):
                    cur.execute("ALTER TABLE notebooks ADD COLUMN deleted_at TEXT NULL")
                if not column_exists('sections', 'deleted_at'):
                    cur.execute("ALTER TABLE sections ADD COLUMN deleted_at TEXT NULL")
                if not column_exists('pages', 'deleted_at'):
                    cur.execute("ALTER TABLE pages ADD COLUMN deleted_at TEXT NULL")
                
                # Create indexes for soft-delete queries
                cur.execute("CREATE INDEX IF NOT EXISTS idx_notebooks_deleted ON notebooks(deleted_at)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_sections_deleted ON sections(deleted_at)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_pages_deleted ON pages(deleted_at)")
                
                conn.commit()
            except Exception as e:
                if parent_window is not None:
                    QtWidgets.QMessageBox.critical(
                        parent_window,
                        "Migration Failed",
                        f"Failed to upgrade database schema:\n{e}"
                    )
                conn.close()
                return False
        
        # Bump version to target
        try:
            cur.execute(f"PRAGMA user_version = {int(TARGET)}")
            conn.commit()
        except Exception:
            pass
        
        return True
    finally:
        conn.close()

def open_database(window):
    """Prompt user to open an existing database file and switch context."""
    dlg_path, _ = QtWidgets.QFileDialog.getOpenFileName(
        window,
        "Open Database",
        get_last_db() or "notes.db",
        "SQLite DB (*.db);;All Files (*)",
    )
    if not dlg_path:
        return
    try:
        ensure_database_initialized(dlg_path)
        if not migrate_database_if_needed(dlg_path, parent_window=window):
            # User cancelled upgrade or chose to open different database
            return
    except Exception as e:
        QtWidgets.QMessageBox.critical(window, "Open Database", f"Failed to open DB: {e}")
        return
    try:
        set_last_db(dlg_path)
        clear_last_state()
    except Exception:
        pass
    window._db_path = dlg_path
    populate_notebook_names(window, dlg_path)
    setup_two_pane(window)
    restore_last_position(window)
    try:
        window.setWindowTitle(f"NoteBook — {dlg_path}")
    except Exception:
        pass


def _enable_faulthandler(log_path: str):
    """Enable Python faulthandler to dump tracebacks on fatal errors (e.g., segfaults).

    Writes native crash backtraces for all threads to the given log file.
    """
    try:
        import faulthandler as _faulthandler
        try:
            f = open(log_path, 'a', encoding='utf-8')
        except Exception:
            f = None
        if f is not None:
            _faulthandler.enable(file=f, all_threads=True)
    except Exception:
        pass


def _install_qt_message_handler(log_path: str):
    """Install a Qt message handler that appends warnings/errors to a log file."""
    try:
        from PyQt5.QtCore import qInstallMessageHandler
    except Exception:
        return
    def _handler(mode, context, message):  # pragma: no cover
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"[QT] {message}\n")
        except Exception:
            pass
    try:
        qInstallMessageHandler(_handler)
    except Exception:
        pass


def _select_left_tree_notebook(window, notebook_id: int):
    """Select a top-level notebook item in the left tree by id."""
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if not tree_widget:
            return
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            nid = top.data(0, 1000)
            try:
                if int(nid) == int(notebook_id):
                    tree_widget.setCurrentItem(top)
                    window._current_notebook_id = int(notebook_id)
                    return
            except Exception:
                pass
    except Exception:
        pass


def add_binder(window):
    """Create a new notebook (binder) and refresh the left tree."""
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    title, ok = QtWidgets.QInputDialog.getText(window, "Add Binder", "Title:", text="Untitled Binder")
    if not ok:
        return
    title = (title or "Untitled Binder").strip() or "Untitled Binder"
    try:
        nid = db_create_notebook(title, db_path)
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, "Add Binder", f"Failed: {e}")
        return
    populate_notebook_names(window, db_path)
    _select_left_tree_notebook(window, nid)
    try:
        set_last_state(notebook_id=int(nid))
    except Exception:
        pass


def create_new_database(window):
    """Create a brand new database file and switch application context to it."""
    dlg_path, _ = QtWidgets.QFileDialog.getSaveFileName(window, "Create New Database", "notes.db", "SQLite DB (*.db);;All Files (*)")
    if not dlg_path:
        return
    try:
        create_new_database_file(dlg_path)
        set_last_db(dlg_path)
        clear_last_state()
        window._db_path = dlg_path
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, "New Database", f"Failed: {e}")
        return
    populate_notebook_names(window, dlg_path)
    try:
        window.setWindowTitle(f"NoteBook — {dlg_path}")
    except Exception:
        pass


def rename_binder(window):
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if not tree_widget:
        return
    item = tree_widget.currentItem()
    if item is None:
        # fallback to first notebook
        item = tree_widget.topLevelItem(0) if tree_widget.topLevelItemCount() > 0 else None
    elif item.parent() is not None:
        # If a section is selected, rename its parent binder
        item = item.parent()
    if item is None:
        QtWidgets.QMessageBox.information(window, "Rename Binder", "No binder selected.")
        return
    nid = item.data(0, 1000)
    current = item.text(0) or ""
    new_title, ok = QtWidgets.QInputDialog.getText(
        window, "Rename Binder", "New title:", text=current
    )
    if not ok or not new_title.strip():
        return
    # Capture and persist expanded state before renaming
    try:
        expanded_ids = set()
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            try:
                if top.isExpanded():
                    tid = top.data(0, 1000)
                    if tid is not None:
                        expanded_ids.add(int(tid))
            except Exception:
                pass
        from settings_manager import set_expanded_notebooks

        set_expanded_notebooks(expanded_ids)
    except Exception:
        pass
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    db_rename_notebook(int(nid), new_title.strip(), db_path)
    populate_notebook_names(window, db_path)
    # Restore expansion from persisted state
    try:
        from settings_manager import get_expanded_notebooks
        from left_tree import ensure_left_tree_sections

        persisted_ids = get_expanded_notebooks()
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            tid = top.data(0, 1000)
            try:
                tid_int = int(tid)
            except Exception:
                tid_int = None
            if tid_int is not None and tid_int in persisted_ids:
                ensure_left_tree_sections(window, tid_int)
    except Exception:
        pass
    _select_left_tree_notebook(window, int(nid))
    restore_last_position(window)


def delete_binder(window):
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if not tree_widget:
        return
    item = tree_widget.currentItem()
    if item is None:
        # fallback to first binder
        item = tree_widget.topLevelItem(0) if tree_widget.topLevelItemCount() > 0 else None
    elif item.parent() is not None:
        # If a section is selected, delete its parent binder
        item = item.parent()
    if item is None:
        QtWidgets.QMessageBox.information(window, "Delete Binder", "No binder selected.")
        return
    # Capture index of the binder being deleted to select an adjacent one afterwards
    try:
        deleted_index = tree_widget.indexOfTopLevelItem(item)
        if deleted_index is None or deleted_index < 0:
            deleted_index = 0
    except Exception:
        deleted_index = 0
    nid = int(item.data(0, 1000))
    title_text = item.text(0) or "(untitled)"
    confirm = QtWidgets.QMessageBox.question(
        window,
        "Delete Binder",
        f'Are you sure you want to delete the binder "{title_text}" and all its sections and pages?',
        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
    )
    if confirm != QtWidgets.QMessageBox.Yes:
        return
    # Capture current expanded state of top-level binders to restore after refresh and persist across restarts
    expanded_ids = set()
    try:
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            try:
                if top.isExpanded():
                    tid = top.data(0, 1000)
                    if tid is not None:
                        expanded_ids.add(int(tid))
            except Exception:
                pass
        # Persist expanded set excluding the one being deleted
        try:
            from settings_manager import set_expanded_notebooks

            persisted = {eid for eid in expanded_ids if eid != int(nid)}
            set_expanded_notebooks(persisted)
        except Exception:
            pass
    except Exception:
        expanded_ids = set()
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    db_delete_notebook(nid, db_path)
    # Clear any remembered state that points to this notebook
    clear_last_state()
    # Refresh UI: repopulate binders (selection will change shortly)
    populate_notebook_names(window, db_path)
    # Restore previously expanded binders (excluding the one we just deleted), based on persisted state
    try:
        from settings_manager import get_expanded_notebooks
        from left_tree import ensure_left_tree_sections

        persisted_ids = get_expanded_notebooks()
        if persisted_ids:
            for i in range(tree_widget.topLevelItemCount()):
                top = tree_widget.topLevelItem(i)
                tid = top.data(0, 1000)
                try:
                    tid_int = int(tid)
                except Exception:
                    tid_int = None
                if tid_int is not None and tid_int in persisted_ids and tid_int != nid:
                    ensure_left_tree_sections(window, tid_int)
    except Exception:
        pass
    # Attempt to select an adjacent remaining binder (same index if possible, else previous)
    remaining = tree_widget.topLevelItemCount()
    if remaining > 0:
        target_index = deleted_index if deleted_index < remaining else remaining - 1
        target_item = tree_widget.topLevelItem(target_index)
        if target_item is not None:
            nb_id = int(target_item.data(0, 1000))
            # Persist and set current notebook context eagerly
            set_last_state(notebook_id=nb_id)
            try:
                window._current_notebook_id = nb_id
            except Exception:
                pass
            _select_left_tree_notebook(window, nb_id)
            # Only expand/populate the selected binder if it was previously expanded (persisted)
            try:
                from settings_manager import get_expanded_notebooks

                if nb_id in get_expanded_notebooks():
                    from left_tree import ensure_left_tree_sections

                    ensure_left_tree_sections(window, nb_id)
            except Exception:
                pass
            # Single unified refresh
            refresh_for_notebook(window, nb_id)
            # Fallback: if binder has sections but tabs are empty, force full UI refresh once
            try:
                tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                sections = db_get_sections_by_notebook_id(nb_id, db_path)
                if sections and (not tab_widget or tab_widget.count() == 0):
                    _full_ui_refresh(window)
                    refresh_for_notebook(window, nb_id)
            except Exception:
                pass
    else:
        # No binders left: clear tabs and right pane explicitly
        tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
        if tab_widget:
            tab_widget.clear()
        right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
        if right_tw:
            right_tw.clear()
        right_tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
        if right_tv and right_tv.model() is not None:
            right_tv.setModel(None)


def add_section(window):
    # Determine target notebook: current selection in left tree
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if not tree_widget or tree_widget.topLevelItemCount() == 0:
        QtWidgets.QMessageBox.information(window, "Add Section", "Please add a binder first.")
        return
    item = tree_widget.currentItem() or tree_widget.topLevelItem(0)
    # If a section is selected, use its parent notebook
    if item and item.parent() is not None:
        item = item.parent()
    nb_id = item.data(0, 1000) if item else None
    if nb_id is None:
        QtWidgets.QMessageBox.information(window, "Add Section", "Please select a binder.")
        return
    title, ok = QtWidgets.QInputDialog.getText(
        window, "Add Section", "Section title:", text="Untitled Section"
    )
    if not ok:
        return
    title = (title or "").strip() or "Untitled Section"
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    sid = db_create_section(int(nb_id), title, db_path)
    # Preserve left-tree state: avoid full repopulate; refresh only the target binder children
    try:
        from left_tree import ensure_left_tree_sections, refresh_for_notebook

        set_last_state(notebook_id=int(nb_id), section_id=sid, page_id=None)
        # Keep current selection but ensure the binder’s children reflect the new section
        ensure_left_tree_sections(window, int(nb_id), select_section_id=sid)
        refresh_for_notebook(window, int(nb_id), select_section_id=sid)
    except Exception:
        # Fallback minimal refresh if helper not available
        set_last_state(notebook_id=int(nb_id), section_id=sid, page_id=None)
        _select_left_tree_notebook(window, int(nb_id))
        refresh_for_notebook(window, int(nb_id), select_section_id=sid)


def _full_ui_refresh(window):
    """Two-pane only: clear left tree, repopulate binders, restore last position."""
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if tree_widget:
        tree_widget.clear()
    populate_notebook_names(window, db_path)
    setup_two_pane(window)
    restore_last_position(window)
    try:
        splitter = window.findChild(QtWidgets.QSplitter, "mainSplitter")
        if splitter is not None:
            splitter.setStretchFactor(0, 0)
            splitter.setStretchFactor(1, 2)
            # Ignore third pane (legacy right panel) if present
    except Exception:
        pass


def add_page(window):
    """Add a new page under the currently selected Section in the left tree.

    Two-pane simplification: ignore legacy tab/right panel paths; rely solely on
    left binder tree selection. If a page is selected, use its parent section.
    """
    section_id = None
    try:
        tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if tree is not None:
            cur = tree.currentItem()
            if cur is not None:
                kind = cur.data(0, 1001)
                if kind == "section":
                    section_id = cur.data(0, 1000)
                elif kind == "page":
                    section_id = cur.data(0, 1002)  # stored parent section id
                elif kind is None and cur.parent() is None:
                    # Binder selected: choose first child section if any
                    for i in range(cur.childCount()):
                        ch = cur.child(i)
                        if ch.data(0, 1001) == "section":
                            section_id = ch.data(0, 1000)
                            break
    except Exception:
        section_id = None
    if section_id is None:
        QtWidgets.QMessageBox.information(window, "Add Page", "Please select a Section first.")
        return
    try:
        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
        title, ok = QtWidgets.QInputDialog.getText(
            window, "Add Page", "Page title:", text="Untitled Page"
        )
        if not ok:
            return
        title = (title or "").strip() or "Untitled Page"
        pid = db_create_page(int(section_id), title, db_path)
        # Refresh children for the section's binder and select the new page
        # Look up the notebook_id for this section instead of using the current active notebook
        nb_id = None
        try:
            import sqlite3
            con = sqlite3.connect(db_path)
            cur = con.cursor()
            cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
            row = cur.fetchone()
            con.close()
            nb_id = int(row[0]) if row else None
        except Exception:
            nb_id = getattr(window, "_current_notebook_id", None)
        if nb_id is not None:
            ensure_left_tree_sections(window, int(nb_id), select_section_id=int(section_id))
        try:
            select_left_tree_page(window, int(section_id), int(pid))
        except Exception:
            pass
        try:
            window._current_section_id = int(section_id)
            if not hasattr(window, "_current_page_by_section"):
                window._current_page_by_section = {}
            window._current_page_by_section[int(section_id)] = int(pid)
        except Exception:
            pass
        try:
            _load_page_two_column(window, int(pid))
        except Exception:
            pass
        try:
            set_last_state(section_id=int(section_id), page_id=int(pid))
        except Exception:
            pass
    except Exception:
        pass


def _current_page_context(window):
    """Return (section_id, page_id) for the currently active context.

    - In two-column UI: use window._current_section_id and window._current_page_by_section
      and fall back to the left tree current selection if needed.
    - In tabbed UI: use the current tab's section and tracked page id.
    """
    try:
        if _is_two_column_ui(window):
            sid = getattr(window, "_current_section_id", None)
            pid = None
            try:
                if sid is not None:
                    pid = getattr(window, "_current_page_by_section", {}).get(int(sid))
            except Exception:
                pid = getattr(window, "_current_page_by_section", {}).get(sid)
            if pid is None:
                # Try left tree selection
                tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
                cur = tree.currentItem() if tree is not None else None
                if cur is not None:
                    kind = cur.data(0, 1001)
                    if kind == "page":
                        pid = cur.data(0, 1000)
                        sid = cur.data(0, 1002)
                    elif kind == "section" and sid is None:
                        sid = cur.data(0, 1000)
            return sid, pid
    except Exception:
        pass

    # Legacy tabs path
    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
    if not tab_widget or tab_widget.count() == 0:
        return None, None
    tab_bar = tab_widget.tabBar()
    if tab_bar is None:
        return None, None
    section_id = tab_bar.tabData(tab_widget.currentIndex())
    if section_id is None:
        return None, None
    # See if a page is currently tracked for this section
    page_id = getattr(window, "_current_page_by_section", {}).get(section_id)
    return section_id, page_id


def insert_attachment(window):
    """Prompt for a file and attach it to the current page via media store.

    If the selected file is an image, also insert it inline at the current caret position.
    Non-image files are just attached (reference saved) without inline HTML.
    """
    section_id, page_id = _current_page_context(window)
    if page_id is None:
        QtWidgets.QMessageBox.information(
            window, "Insert Attachment", "Please open or create a page first."
        )
        return
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    options = QtWidgets.QFileDialog.Options()
    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
        window,
        "Select Attachment",
        "",
        "All Files (*);;Images (*.png *.jpg *.jpeg *.gif *.bmp);;PDF (*.pdf)",
        options=options,
    )
    if not file_path:
        return
    try:
        from media_store import add_media_ref, save_file_into_store, guess_mime_and_ext

        # Save into store and record attachment ref
        media_id, rel_path = save_file_into_store(db_path, file_path)
        add_media_ref(db_path, media_id, page_id=page_id, role="attachment")

        # If it's an image, also insert inline at the caret using a relative src
        mime, ext = guess_mime_and_ext(file_path)
        is_image_mime = isinstance(mime, str) and mime.lower().startswith("image/")
        raw_exts = {
            "dng", "nef", "cr2", "cr3", "arw", "orf", "rw2", "raf", "srw", "pef",
            "rw1", "3fr", "erf", "kdc", "mrw", "nrw", "ptx", "r3d", "sr2", "x3f"
        }
        did_inline = False
        if is_image_mime and (ext or "").lower() not in raw_exts:
            te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
            if te is not None:
                # Use default width from settings; fallback to 400px
                try:
                    from settings_manager import get_image_insert_long_side

                    long_side = int(get_image_insert_long_side())
                except Exception:
                    long_side = 400
                # Insert HTML img tag; baseUrl is already set to media root
                name_attr = rel_path.replace("\\", "/")
                html = f'<img src="{name_attr}" width="{int(max(16, long_side))}" />'
                try:
                    cur = te.textCursor()
                    before = cur.position()
                    cur.insertHtml(html)
                    after = cur.position()
                    te.setTextCursor(cur)
                    did_inline = True
                except Exception:
                    pass
        # If we didn't insert an inline image, insert a clickable link to the attachment instead
        if not did_inline:
            te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
            if te is not None:
                try:
                    # Use original filename as link text
                    link_text = os.path.basename(file_path)
                except Exception:
                    link_text = rel_path.replace("\\", "/")
                href = rel_path.replace("\\", "/")
                # Insert an anchor; baseUrl is set so relative href opens the local file
                link_html = f'<a href="{href}">📎 {link_text}</a>'
                try:
                    cur = te.textCursor()
                    cur.insertHtml(link_html)
                    te.setTextCursor(cur)
                except Exception:
                    pass
        # Non-intrusive confirmation
        if did_inline:
            show_toast(window, f"Inserted image + attached: {rel_path}", 2500)
        else:
            show_toast(window, f"Attached: {rel_path}", 2500)
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, "Insert Attachment", f"Failed to attach file: {e}")


def backup_database_now(window):
    """Create a manual backup using current settings (destination and retention).

    - Uses the same bundle format as on-exit backups (DB at root, media/ folder included).
    - If no destination folder is configured, prompt the user to choose one for this run.
    """
    try:
        # Flush any unsaved edits before snapshotting
        try:
            save_current_page(window)
        except Exception:
            pass

        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
        from settings_manager import (
            get_exit_backup_dir,
            get_backups_to_keep,
            set_exit_backup_dir,
        )

        dest = (get_exit_backup_dir() or "").strip()
        if not dest:
            options = QtWidgets.QFileDialog.Options()
            chosen = QtWidgets.QFileDialog.getExistingDirectory(
                window, "Choose Backup Folder", os.path.dirname(db_path) or "", options=options
            )
            if not chosen:
                return
            dest = chosen
            try:
                remember = QtWidgets.QMessageBox.question(
                    window,
                    "Remember Backup Folder",
                    f"Use this folder for future backups?\n{dest}",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                )
                if remember == QtWidgets.QMessageBox.Yes:
                    set_exit_backup_dir(dest)
            except Exception:
                pass

        QtWidgets.QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            from backup import make_exit_backup
            bundle = make_exit_backup(
                db_path, dest, keep=int(get_backups_to_keep()), include_media=True
            )
        finally:
            try:
                QtWidgets.QApplication.restoreOverrideCursor()
            except Exception:
                pass

        if bundle:
            QtWidgets.QMessageBox.information(
                window, "Backup Complete", f"Backup created:\n{bundle}"
            )
        else:
            QtWidgets.QMessageBox.warning(
                window,
                "Backup Failed",
                "Backup did not create a bundle."
            )
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, "Backup Failed", f"Error: {e}")

def save_database_as(window):
    """Save the current database and its media folder under a new name (copy) and switch to it."""
    # Ensure any unsaved edits are flushed first
    try:
        save_current_page(window)
    except Exception:
        pass
    cur_db = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    # Propose a default new name in the Databases root
    try:
        import os

        from settings_manager import get_databases_root

        base = os.path.basename(cur_db)
        name, ext = os.path.splitext(base)
        proposed = name + "_copy" + (ext if ext else ".db")
        initial = os.path.join(get_databases_root(), proposed)
    except Exception:
        initial = cur_db
    options = QtWidgets.QFileDialog.Options()
    new_path, _ = QtWidgets.QFileDialog.getSaveFileName(
        window,
        "Save Database As",
        initial,
        "SQLite DB Files (*.db);;All Files (*)",
        options=options,
    )
    if not new_path:
        return
    if not str(new_path).lower().endswith(".db"):
        new_path = new_path + ".db"
    try:
        import os
        import shutil

        # Confirm overwrite if target exists
        if os.path.exists(new_path):
            resp = QtWidgets.QMessageBox.question(
                window,
                "Overwrite File?",
                f"{new_path}\nAlready exists. Overwrite?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if resp != QtWidgets.QMessageBox.Yes:
                return
        # Copy DB file
        shutil.copy2(cur_db, new_path)
        # Copy media folder tree if present
        from media_store import media_root_for_db

        src_media = media_root_for_db(cur_db)
        dst_media = media_root_for_db(new_path)
        if os.path.isdir(src_media):
            # If destination exists and we are overwriting, remove it first to avoid nested copies
            if os.path.exists(dst_media):
                shutil.rmtree(dst_media, ignore_errors=True)
            shutil.copytree(src_media, dst_media)
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, "Save As", f"Failed to copy database or media: {e}")
        return
    # Switch to the new database
    try:
        set_last_db(new_path)
        clear_last_state()
    except Exception:
        pass
    restart_application()


def restart_application():
    """Restart the application process with the same interpreter and script."""
    import os

    python = sys.executable
    script = os.path.abspath(__file__)
    # Start a new detached process and quit current app
    QProcess.startDetached(python, [script])
    QtWidgets.QApplication.quit()


def print_current_selection(window):
    """Print the currently selected item (page, section, or binder)."""
    try:
        from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
        from PyQt5.QtGui import QTextDocument
        
        # Get HTML content for selected item
        html_content = _get_print_html_content(window)
        if not html_content:
            return
        
        # Create printer and dialog
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, window)
        dialog.setWindowTitle("Print")
        
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            # Create document and print
            document = QTextDocument()
            document.setHtml(html_content)
            document.print_(printer)
            
    except Exception as e:
        QtWidgets.QMessageBox.critical(window, "Print Error", f"Failed to print: {str(e)}")


def print_preview_current_selection(window):
    """Show print preview for the currently selected item (page, section, or binder)."""
    try:
        from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog
        from PyQt5.QtGui import QTextDocument
        
        # Get HTML content for selected item
        html_content = _get_print_html_content(window)
        if not html_content:
            return
        
        # Create document
        document = QTextDocument()
        document.setHtml(html_content)
        
        # Create printer and preview dialog
        printer = QPrinter(QPrinter.HighResolution)
        preview = QPrintPreviewDialog(printer, window)
        preview.setWindowTitle("Print Preview")
        
        # Connect the paintRequested signal to render the document
        preview.paintRequested.connect(lambda p: document.print_(p))
        
        # Show preview dialog (which includes print button)
        preview.exec_()
            
    except Exception as e:
        QtWidgets.QMessageBox.critical(window, "Print Error", f"Failed to print: {str(e)}")


def _get_print_html_content(window):
    """Build HTML content for printing based on current selection."""
    try:
        # Determine what's selected in the left tree
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if tree_widget is None:
            QtWidgets.QMessageBox.information(window, "Print", "Please select a page, section, or binder to print.")
            return None
        
        current_item = tree_widget.currentItem()
        if current_item is None:
            QtWidgets.QMessageBox.information(window, "Print", "Please select a page, section, or binder to print.")
            return None
        
        # Get item type and ID
        item_kind = current_item.data(0, 1001)  # 'page', 'section', or None (binder)
        item_id = current_item.data(0, 1000)
        
        db_path = getattr(window, "_db_path", None) or "notes.db"
        
        # Build HTML content based on selection
        html_content = "<html><head><style>"
        html_content += "body { font-family: Arial, sans-serif; margin: 20px; }"
        html_content += ".page-header { text-align: center; font-size: 10pt; color: #666; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 15px; }"
        html_content += ".page-title { text-align: center; font-size: 14pt; font-weight: bold; margin: 15px 0; }"
        html_content += ".page-content { margin-top: 10px; }"
        html_content += ".page-break { page-break-before: always; }"
        html_content += "</style></head><body>"
        
        if item_kind == "page":
            # Print single page - get binder and section context
            binder_name, section_name = _get_page_context(current_item)
            html_content += _get_page_html(item_id, db_path, binder_name, section_name)
        elif item_kind == "section":
            # Print all pages in section - get binder context
            section_title = current_item.text(0)
            binder_name = current_item.parent().text(0) if current_item.parent() else "Unknown Binder"
            html_content += _get_section_pages_html(item_id, db_path, binder_name, section_title)
        else:
            # Print entire binder (notebook)
            binder_title = current_item.text(0)
            html_content += _get_binder_html(item_id, db_path, binder_title)
        
        html_content += "</body></html>"
        return html_content
        
    except Exception as e:
        QtWidgets.QMessageBox.critical(window, "Print Error", f"Failed to build print content: {str(e)}")
        return None


def _get_page_context(tree_item):
    """Get binder and section names for a page tree item."""
    binder_name = "Unknown Binder"
    section_name = "Unknown Section"
    
    # Walk up the tree to find section and binder
    parent = tree_item.parent()
    if parent:
        if parent.data(0, 1001) == "section":
            section_name = parent.text(0)
            grandparent = parent.parent()
            if grandparent:
                binder_name = grandparent.text(0)
        elif parent.data(0, 1001) == "page":
            # This is a subpage - find section through parent page
            while parent and parent.data(0, 1001) == "page":
                parent = parent.parent()
            if parent and parent.data(0, 1001) == "section":
                section_name = parent.text(0)
                grandparent = parent.parent()
                if grandparent:
                    binder_name = grandparent.text(0)
    
    return binder_name, section_name


def _get_page_html(page_id, db_path, binder_name=None, section_name=None):
    """Get HTML for a single page with hierarchical header."""
    try:
        from db_pages import get_page_by_id
        import sqlite3
        
        page = get_page_by_id(int(page_id), db_path)
        if not page:
            return "<p>Page not found.</p>"
        
        title = page[2]  # title is at index 2
        content = page[3]  # content is at index 3
        
        html = ""
        # Add hierarchical header if we have context
        if binder_name and section_name:
            html += f'<div class="page-header">{binder_name} &gt; {section_name} &gt; {title}</div>'
        html += f'<div class="page-title">{title}</div>'
        html += f'<div class="page-content">{content}</div>'
        return html
    except Exception:
        return "<p>Error loading page.</p>"


def _get_section_pages_html(section_id, db_path, binder_name, section_name):
    """Get HTML for all pages in a section with hierarchical headers."""
    try:
        from db_pages import get_root_pages_by_section_id
        import sqlite3
        
        html = ""
        pages = get_root_pages_by_section_id(int(section_id), db_path)
        
        for idx, page in enumerate(pages):
            page_id = page[0]
            if idx > 0:
                html += '<div class="page-break"></div>'
            html += _get_page_html(page_id, db_path, binder_name, section_name)
            # Recursively get subpages
            html += _get_subpages_html(page_id, db_path, binder_name, section_name)
        
        return html if html else "<p>No pages in this section.</p>"
    except Exception:
        return "<p>Error loading section pages.</p>"


def _get_subpages_html(parent_page_id, db_path, binder_name, section_name):
    """Recursively get HTML for subpages with hierarchical headers."""
    try:
        import sqlite3
        
        html = ""
        con = sqlite3.connect(db_path)
        cur = con.cursor()
        cur.execute(
            "SELECT id FROM pages WHERE parent_page_id = ? ORDER BY order_index, id",
            (int(parent_page_id),)
        )
        subpages = cur.fetchall()
        con.close()
        
        for (page_id,) in subpages:
            html += '<div class="page-break"></div>'
            html += _get_page_html(page_id, db_path, binder_name, section_name)
            # Recursively get children of this subpage
            html += _get_subpages_html(page_id, db_path, binder_name, section_name)
        
        return html
    except Exception:
        return ""


def _get_binder_html(notebook_id, db_path, binder_name):
    """Get HTML for all sections and pages in a binder with hierarchical headers."""
    try:
        from db_sections import get_sections_by_notebook_id
        
        html = ""
        sections = get_sections_by_notebook_id(int(notebook_id), db_path)
        first_page = True
        
        for section in sections:
            section_id = section[0]
            section_title = section[2]
            
            # Get pages for this section - pass context and first_page flag
            section_html = _get_section_pages_html_for_binder(section_id, db_path, binder_name, section_title, first_page)
            if section_html:
                html += section_html
                first_page = False
        
        return html if html else "<p>No sections in this binder.</p>"
    except Exception:
        return "<p>Error loading binder content.</p>"


def _get_section_pages_html_for_binder(section_id, db_path, binder_name, section_name, first_page):
    """Get HTML for section pages within a binder print, handling first page specially."""
    try:
        from db_pages import get_root_pages_by_section_id
        
        html = ""
        pages = get_root_pages_by_section_id(int(section_id), db_path)
        
        for idx, page in enumerate(pages):
            page_id = page[0]
            # Add page break before every page except the very first one
            if not (first_page and idx == 0):
                html += '<div class="page-break"></div>'
            html += _get_page_html(page_id, db_path, binder_name, section_name)
            # Recursively get subpages
            html += _get_subpages_html(page_id, db_path, binder_name, section_name)
        
        return html
    except Exception:
        return ""


def main():
    # Suppress noisy SIP deprecation warning from PyQt5 about sipPyTypeDict
    warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*sipPyTypeDict.*")
    # Ensure proper High DPI behavior so images render at the requested logical size on scaled displays
    try:
        # These must be set BEFORE creating the QApplication instance
        from PyQt5.QtCore import Qt as _Qt
        try:
            QtWidgets.QApplication.setAttribute(_Qt.AA_EnableHighDpiScaling, True)
        except Exception:
            pass
        try:
            QtWidgets.QApplication.setAttribute(_Qt.AA_UseHighDpiPixmaps, True)
        except Exception:
            pass
    except Exception:
        pass
    app = QtWidgets.QApplication(sys.argv)
    
    # Install a global exception hook so unexpected errors surface in a dialog instead of closing silently
    try:
        _install_global_excepthook()
    except Exception:
        pass
    # Prepare crash/diagnostic logs
    try:
        _here = os.path.dirname(os.path.abspath(__file__))
        _crash_log = os.path.join(_here, "crash.log")
        _native_log = os.path.join(_here, "native_crash.log")
        _enable_faulthandler(_native_log)
        _install_qt_message_handler(_crash_log)
    except Exception:
        pass

    # Safe mode: disable risky UI hooks to isolate crashes quickly
    SAFE_MODE = os.environ.get("NOTEBOOK_SAFE_MODE", "0").strip() in {"1", "true", "yes"}
    window = load_main_window()
    # Ensure the main window explicitly carries the app icon (improves taskbar behavior on Windows)
    try:
        window.setWindowIcon(app.windowIcon())
    except Exception:
        pass
    # Apply saved theme QSS early
    try:
        from settings_manager import get_theme_name

        theme = get_theme_name()
        themes_dir = os.path.join(os.path.dirname(__file__), "themes")
        name_to_file = {
            "Default": "default.qss",
            "High Contrast": "high-contrast.qss",
        }
        qss_file = name_to_file.get(theme)
        if qss_file:
            path = os.path.join(themes_dir, qss_file)
            if os.path.isfile(path):
                with open(path, "r", encoding="utf-8") as f:
                    app.setStyleSheet(f.read())
    except Exception:
        pass
    # Restore window geometry and maximized state
    geom = get_window_geometry()
    if geom and all(k in geom for k in ("x", "y", "w", "h")):
        window.setGeometry(int(geom["x"]), int(geom["y"]), int(geom["w"]), int(geom["h"]))
    if get_window_maximized():
        window.showMaximized()
    db_path = get_last_db() or "notes.db"
    # Ensure database exists and is initialized before any queries
    try:
        ensure_database_initialized(db_path)
        if not migrate_database_if_needed(db_path, parent_window=window):
            # User cancelled upgrade - offer to create new or open different database
            msg_box = QtWidgets.QMessageBox(window)
            msg_box.setWindowTitle("Database Not Upgraded")
            msg_box.setIcon(QtWidgets.QMessageBox.Question)
            msg_box.setText("The database was not upgraded. What would you like to do?")
            create_btn = msg_box.addButton("Create New Database", QtWidgets.QMessageBox.AcceptRole)
            open_btn = msg_box.addButton("Open Different Database", QtWidgets.QMessageBox.ActionRole)
            exit_btn = msg_box.addButton("Exit", QtWidgets.QMessageBox.RejectRole)
            msg_box.exec_()
            clicked = msg_box.clickedButton()
            if clicked == create_btn:
                create_new_database(window)
            elif clicked == open_btn:
                open_database(window)
            return
    except Exception as e:
        # If database initialization fails, show error and create new database dialog
        QtWidgets.QMessageBox.critical(
            window,
            "Database Error",
            f"Failed to initialize database '{db_path}':\n{str(e)}\n\nPlease create a new database or select an existing one."
        )
        create_new_database(window)
        return
    window._db_path = db_path
    # Show current DB in the window title (avoid duplicating in the status bar)
    try:
        window.setWindowTitle(f"NoteBook — {db_path}")
    except Exception:
        pass

    # (Binder menu removed from UI definition; no runtime removal needed)
    # Prepare media root path for this database (not yet used by UI)
    try:
        from media_store import ensure_dir, media_root_for_db

        window._media_root = media_root_for_db(db_path)
        ensure_dir(window._media_root)
    except Exception:
        window._media_root = None
    # Defensive: initialize the rich-text document baseUrl early so relative media src resolves on first load
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        media_root = getattr(window, "_media_root", None)
        if te is not None and media_root:
            # Ensure trailing separator so relative paths resolve as children of this directory
            if not media_root.endswith(os.sep) and not media_root.endswith("/"):
                media_root = media_root + os.sep
            te.document().setBaseUrl(QUrl.fromLocalFile(media_root))
            # Also optionally disable image resize overlay in safe mode
            if SAFE_MODE:
                os.environ["NOTEBOOK_DISABLE_IMAGE_RESIZE"] = "1"
            # Install image context menu and keyboard shortcuts regardless of toolbar wiring
            try:
                install_image_support(te)
            except Exception:
                pass
            # Install spell checking
            try:
                from spell_check import install_spell_check, is_spell_check_available
                from settings_manager import get_spell_check_enabled, get_spell_check_language
                if is_spell_check_available():
                    spell_enabled = get_spell_check_enabled()
                    spell_lang = get_spell_check_language()
                    spell_checker = install_spell_check(te, enabled=spell_enabled, language=spell_lang)
                    if spell_checker:
                        window._spell_checker = spell_checker
            except Exception:
                pass
    except Exception:
        pass
    populate_notebook_names(window, db_path)
    setup_two_pane(window)
    restore_last_position(window)
    # Apply saved list scheme (ordered/unordered) to rich text
    try:
        from settings_manager import get_list_schemes_settings
        from ui_richtext import set_list_schemes

        ord_s, unord_s = get_list_schemes_settings()
        set_list_schemes(ordered=ord_s, unordered=unord_s)
    except Exception:
        pass
    # Apply default image insert size from settings
    try:
        from settings_manager import get_image_insert_long_side
        import ui_richtext as rt

        rt.DEFAULT_IMAGE_LONG_SIDE = int(get_image_insert_long_side())
    except Exception:
        pass
    # Apply default paste mode to override Ctrl+V behavior
    try:
        from settings_manager import get_default_paste_mode

        window._default_paste_mode = get_default_paste_mode()
    except Exception:
        window._default_paste_mode = "rich"
    # Restore left-panel expanded binders from settings after initial build
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        from settings_manager import get_expanded_notebooks
        expanded_ids = get_expanded_notebooks()
        if tree_widget is not None and expanded_ids:
            for i in range(tree_widget.topLevelItemCount()):
                top = tree_widget.topLevelItem(i)
                tid = top.data(0, 1000)
                try:
                    tid_int = int(tid)
                except Exception:
                    tid_int = None
                if tid_int is not None and tid_int in expanded_ids:
                    ensure_left_tree_sections(window, tid_int)
    except Exception:
        pass

    # Ensure planning register watcher is active on the editor to support
    # formatting and totals for existing content as well.
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None and not SAFE_MODE:
            ensure_planning_register_watcher(te)
            try:
                from ui_richtext import ensure_currency_columns_watcher
                ensure_currency_columns_watcher(te)
            except Exception:
                pass
    except Exception:
        pass

    # Left binder tree: unified context menu (New/Rename/Delete Binder; New Binder on blank space)
    try:
        tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if tree is not None:
            tree.setContextMenuPolicy(Qt.CustomContextMenu)

            def _tree_ctx_menu(pos):
                try:
                    item = tree.itemAt(pos)
                    global_pos = tree.viewport().mapToGlobal(pos)
                    
                    # Helper to get show_deleted setting
                    def _get_show_deleted_setting():
                        try:
                            from settings_manager import get_show_deleted
                            return get_show_deleted()
                        except Exception:
                            return False
                    
                    # Helper to toggle show_deleted and refresh tree
                    def _toggle_show_deleted():
                        try:
                            from settings_manager import get_show_deleted, set_show_deleted
                            current = get_show_deleted()
                            set_show_deleted(not current)
                            # Sync the File menu's Show Deleted Items action
                            if hasattr(window, "_show_deleted_action"):
                                window._show_deleted_action.setChecked(not current)
                        except Exception:
                            pass
                        # Refresh the tree
                        try:
                            db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                            populate_notebook_names(window, db_path)
                            nb_id = getattr(window, "_current_notebook_id", None)
                            if nb_id is not None:
                                ensure_left_tree_sections(window, int(nb_id))
                        except Exception:
                            pass
                    
                    # Blank area: offer New Binder
                    if item is None:
                        m = QtWidgets.QMenu(tree)
                        act_new = m.addAction("New Binder")
                        m.addSeparator()
                        act_collapse_all = m.addAction("Collapse All Binders")
                        m.addSeparator()
                        # Show Deleted Items toggle
                        act_show_deleted = m.addAction("Show Deleted Items")
                        act_show_deleted.setCheckable(True)
                        act_show_deleted.setChecked(_get_show_deleted_setting())
                        chosen = m.exec_(global_pos)
                        if chosen == act_new:
                            add_binder(window)
                        elif chosen == act_collapse_all:
                            try:
                                # Collapse all top-level items and clear persisted expanded state
                                for i in range(tree.topLevelItemCount()):
                                    top = tree.topLevelItem(i)
                                    top.setExpanded(False)
                                try:
                                    from settings_manager import set_expanded_notebooks

                                    set_expanded_notebooks(set())
                                except Exception:
                                    pass
                            except Exception:
                                pass
                        elif chosen == act_show_deleted:
                            _toggle_show_deleted()
                        return
                    
                    # Check if item is deleted
                    is_item_deleted = bool(item.data(0, 1003))
                    
                    # Top-level binder item
                    if item.parent() is None:
                        tree.setCurrentItem(item)
                        m = QtWidgets.QMenu(tree)
                        
                        if is_item_deleted:
                            # Deleted binder: show restore/permanent delete options
                            act_restore = m.addAction("Restore Binder")
                            act_perm_delete = m.addAction("Delete Permanently")
                            m.addSeparator()
                            act_show_deleted = m.addAction("Show Deleted Items")
                            act_show_deleted.setCheckable(True)
                            act_show_deleted.setChecked(_get_show_deleted_setting())
                            chosen = m.exec_(global_pos)
                            if chosen == act_restore:
                                try:
                                    from db_access import restore_notebook
                                    nb_id = item.data(0, 1000)
                                    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                                    restore_notebook(int(nb_id), db_path)
                                    populate_notebook_names(window, db_path)
                                    ensure_left_tree_sections(window, int(nb_id))
                                except Exception:
                                    pass
                            elif chosen == act_perm_delete:
                                nb_name = item.text(0) or "(untitled)"
                                confirm = QtWidgets.QMessageBox.warning(
                                    tree,
                                    "Delete Permanently",
                                    f'Permanently delete binder "{nb_name}" and all its contents?\n\nThis cannot be undone.',
                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                )
                                if confirm == QtWidgets.QMessageBox.Yes:
                                    try:
                                        from db_access import permanently_delete_notebook
                                        nb_id = item.data(0, 1000)
                                        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                                        permanently_delete_notebook(int(nb_id), db_path)
                                        populate_notebook_names(window, db_path)
                                    except Exception:
                                        pass
                            elif chosen == act_show_deleted:
                                _toggle_show_deleted()
                            return
                        
                        # Normal binder menu
                        # Place 'New Section' at the very top, followed by a separator
                        act_new_section = m.addAction("New Section")
                        m.addSeparator()
                        # Binder operations
                        act_new = m.addAction("New Binder")
                        act_rename = m.addAction("Rename Binder")
                        act_delete = m.addAction("Delete Binder")
                        m.addSeparator()
                        act_collapse_all = m.addAction("Collapse All Binders")
                        m.addSeparator()
                        act_show_deleted = m.addAction("Show Deleted Items")
                        act_show_deleted.setCheckable(True)
                        act_show_deleted.setChecked(_get_show_deleted_setting())
                        m.addSeparator()
                        act_collapse_all = m.addAction("Collapse All Binders")
                        chosen = m.exec_(global_pos)
                        if chosen == act_new:
                            add_binder(window)
                        elif chosen == act_rename:
                            rename_binder(window)
                        elif chosen == act_delete:
                            delete_binder(window)
                        elif chosen == act_new_section:
                            add_section(window)
                        elif chosen == act_collapse_all:
                            try:
                                for i in range(tree.topLevelItemCount()):
                                    top = tree.topLevelItem(i)
                                    top.setExpanded(False)
                                try:
                                    from settings_manager import set_expanded_notebooks

                                    set_expanded_notebooks(set())
                                except Exception:
                                    pass
                            except Exception:
                                pass
                        elif chosen == act_show_deleted:
                            _toggle_show_deleted()
                        return
                    # Non top-level (section or page)
                    tree.setCurrentItem(item)
                    m = QtWidgets.QMenu(tree)
                    kind = item.data(0, 1001)
                    if kind == "section":
                        if is_item_deleted:
                            # Deleted section: show restore/permanent delete options
                            act_restore = m.addAction("Restore Section")
                            act_perm_delete = m.addAction("Delete Permanently")
                            m.addSeparator()
                            act_show_deleted = m.addAction("Show Deleted Items")
                            act_show_deleted.setCheckable(True)
                            act_show_deleted.setChecked(_get_show_deleted_setting())
                            chosen = m.exec_(global_pos)
                            if chosen == act_restore:
                                try:
                                    # Check if parent binder is deleted - can't restore into a deleted binder
                                    parent = item.parent()
                                    if parent is not None and bool(parent.data(0, 1003)):
                                        QtWidgets.QMessageBox.warning(
                                            tree,
                                            "Cannot Restore",
                                            "Cannot restore this section because its parent binder is deleted.\n\nPlease restore the binder first.",
                                        )
                                        return
                                    from db_sections import restore_section
                                    section_id = item.data(0, 1000)
                                    nb_id = parent.data(0, 1000) if parent is not None else None
                                    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                                    restore_section(int(section_id), db_path)
                                    if nb_id is not None:
                                        ensure_left_tree_sections(window, int(nb_id))
                                except Exception:
                                    pass
                            elif chosen == act_perm_delete:
                                sec_name = item.text(0) or "(untitled)"
                                confirm = QtWidgets.QMessageBox.warning(
                                    tree,
                                    "Delete Permanently",
                                    f'Permanently delete section "{sec_name}" and all its pages?\n\nThis cannot be undone.',
                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                )
                                if confirm == QtWidgets.QMessageBox.Yes:
                                    try:
                                        from db_sections import permanently_delete_section
                                        section_id = item.data(0, 1000)
                                        parent = item.parent()
                                        nb_id = parent.data(0, 1000) if parent is not None else None
                                        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                                        permanently_delete_section(int(section_id), db_path)
                                        if nb_id is not None:
                                            ensure_left_tree_sections(window, int(nb_id))
                                    except Exception:
                                        pass
                            elif chosen == act_show_deleted:
                                _toggle_show_deleted()
                            return
                        
                        # Normal section menu
                        act_add_page = m.addAction("Add Page")
                        m.addSeparator()
                        act_new_section = m.addAction("New Section")
                        act_rename_section = m.addAction("Rename Section")
                        act_delete_section = m.addAction("Delete Section")
                        m.addSeparator()
                        act_show_deleted = m.addAction("Show Deleted Items")
                        act_show_deleted.setCheckable(True)
                        act_show_deleted.setChecked(_get_show_deleted_setting())
                        chosen = m.exec_(global_pos)
                        if chosen is None:
                            return
                        if chosen == act_add_page:
                            add_page(window)
                            return
                        if chosen == act_new_section:
                            add_section(window)
                            return
                        if chosen == act_show_deleted:
                            _toggle_show_deleted()
                            return
                        # Get ids/context
                        section_id = item.data(0, 1000)
                        parent = item.parent()
                        nb_id = parent.data(0, 1000) if parent is not None else None
                        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                        if chosen == act_rename_section and section_id is not None:
                            current_text = item.text(0) or ""
                            new_title, ok = QtWidgets.QInputDialog.getText(
                                tree, "Rename Section", "New title:", text=current_text
                            )
                            if ok and new_title.strip():
                                try:
                                    db_rename_section(int(section_id), new_title.strip(), db_path)
                                except Exception:
                                    pass
                                # Update UI bits
                                try:
                                    item.setText(0, new_title.strip())
                                except Exception:
                                    pass
                                try:
                                    if nb_id is not None:
                                        refresh_for_notebook(
                                            window, int(nb_id), select_section_id=int(section_id)
                                        )
                                        ensure_left_tree_sections(
                                            window, int(nb_id), select_section_id=int(section_id)
                                        )
                                except Exception:
                                    pass
                            return
                        if chosen == act_delete_section and section_id is not None:
                            sec_name = item.text(0) or "(untitled)"
                            confirm = QtWidgets.QMessageBox.question(
                                tree,
                                "Delete Section",
                                f'Are you sure you want to delete the section "{sec_name}" and all its pages?',
                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                            )
                            if confirm != QtWidgets.QMessageBox.Yes:
                                return
                            try:
                                save_current_page(window)
                            except Exception:
                                pass
                            try:
                                db_delete_section(int(section_id), db_path)
                            except Exception:
                                pass
                            if nb_id is not None:
                                try:
                                    refresh_for_notebook(window, int(nb_id))
                                    ensure_left_tree_sections(window, int(nb_id))
                                except Exception:
                                    pass
                            return
                    elif kind == "page":
                        if is_item_deleted:
                            # Deleted page: show restore/permanent delete options
                            act_restore = m.addAction("Restore Page")
                            act_perm_delete = m.addAction("Delete Permanently")
                            m.addSeparator()
                            act_show_deleted = m.addAction("Show Deleted Items")
                            act_show_deleted.setCheckable(True)
                            act_show_deleted.setChecked(_get_show_deleted_setting())
                            chosen = m.exec_(global_pos)
                            if chosen == act_restore:
                                try:
                                    # Check if parent section (or binder) is deleted - can't restore into deleted parent
                                    parent = item.parent()
                                    # Find the section by traversing up (parent could be another page or a section)
                                    section_item = parent
                                    while section_item is not None and section_item.data(0, 1001) == "page":
                                        section_item = section_item.parent()
                                    # Check if section is deleted
                                    if section_item is not None and bool(section_item.data(0, 1003)):
                                        QtWidgets.QMessageBox.warning(
                                            tree,
                                            "Cannot Restore",
                                            "Cannot restore this page because its parent section is deleted.\n\nPlease restore the section first.",
                                        )
                                        return
                                    # Check if binder (grandparent) is deleted
                                    binder_item = section_item.parent() if section_item is not None else None
                                    if binder_item is not None and bool(binder_item.data(0, 1003)):
                                        QtWidgets.QMessageBox.warning(
                                            tree,
                                            "Cannot Restore",
                                            "Cannot restore this page because its binder is deleted.\n\nPlease restore the binder first.",
                                        )
                                        return
                                    from db_pages import restore_page
                                    page_id = item.data(0, 1000)
                                    section_id = item.data(0, 1002)
                                    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                                    restore_page(int(page_id), db_path)
                                    # Refresh tree
                                    import sqlite3
                                    con = sqlite3.connect(db_path)
                                    cur = con.cursor()
                                    cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
                                    row = cur.fetchone()
                                    con.close()
                                    nb_id = int(row[0]) if row else getattr(window, "_current_notebook_id", None)
                                    if nb_id is not None:
                                        ensure_left_tree_sections(window, int(nb_id), select_section_id=int(section_id))
                                except Exception:
                                    pass
                            elif chosen == act_perm_delete:
                                page_name = item.text(0) or "(untitled)"
                                confirm = QtWidgets.QMessageBox.warning(
                                    tree,
                                    "Delete Permanently",
                                    f'Permanently delete page "{page_name}" and all its subpages?\n\nThis cannot be undone.',
                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                )
                                if confirm == QtWidgets.QMessageBox.Yes:
                                    try:
                                        from db_pages import permanently_delete_page
                                        page_id = item.data(0, 1000)
                                        section_id = item.data(0, 1002)
                                        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                                        permanently_delete_page(int(page_id), db_path)
                                        # Refresh tree
                                        import sqlite3
                                        con = sqlite3.connect(db_path)
                                        cur = con.cursor()
                                        cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
                                        row = cur.fetchone()
                                        con.close()
                                        nb_id = int(row[0]) if row else getattr(window, "_current_notebook_id", None)
                                        if nb_id is not None:
                                            ensure_left_tree_sections(window, int(nb_id), select_section_id=int(section_id))
                                    except Exception:
                                        pass
                            elif chosen == act_show_deleted:
                                _toggle_show_deleted()
                            return
                        
                        # Normal page menu
                        act_add_page = m.addAction("Add Page")
                        act_add_subpage = m.addAction("Add Subpage")
                        act_rename_page = m.addAction("Rename Page")
                        act_delete_page = m.addAction("Delete Page")
                        m.addSeparator()
                        act_show_deleted = m.addAction("Show Deleted Items")
                        act_show_deleted.setCheckable(True)
                        act_show_deleted.setChecked(_get_show_deleted_setting())
                        chosen = m.exec_(global_pos)
                        if chosen is None:
                            return
                        if chosen == act_show_deleted:
                            _toggle_show_deleted()
                            return
                        # Context: ids
                        page_id = item.data(0, 1000)
                        section_id = item.data(0, 1002)
                        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                        if chosen == act_add_page:
                            add_page(window)
                            return
                        if chosen == act_add_subpage and page_id is not None and section_id is not None:
                            # Prompt for title and create a child page under this page
                            new_title, ok = QtWidgets.QInputDialog.getText(
                                tree, "New Subpage", "Subpage title:", text="Untitled Page"
                            )
                            if not ok:
                                return
                            new_title = (new_title or "").strip() or "Untitled Page"
                            # Before creating, persist current root page order to prevent shuffle on refresh
                            try:
                                ordered_root_ids = []
                                # Find the section item in the left tree to read current visual order
                                sec_item = item
                                while sec_item is not None and sec_item.data(0, 1001) != "section":
                                    sec_item = sec_item.parent()
                                if sec_item is not None:
                                    for j in range(sec_item.childCount()):
                                        ch = sec_item.child(j)
                                        try:
                                            if ch.data(0, 1001) == "page" and ch.parent() is sec_item:
                                                pid = ch.data(0, 1000)
                                                if pid is not None:
                                                    ordered_root_ids.append(int(pid))
                                        except Exception:
                                            pass
                                if ordered_root_ids:
                                    try:
                                        db_set_pages_order(int(section_id), ordered_root_ids, db_path)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                            try:
                                from db_pages import create_page as db_create_page

                                new_pid = db_create_page(int(section_id), new_title, db_path, parent_page_id=int(page_id))
                            except Exception:
                                new_pid = None
                            # Refresh left tree for this binder and select the new subpage
                            # Look up the notebook_id for this section instead of traversing the tree
                            nb_id = None
                            try:
                                import sqlite3
                                con = sqlite3.connect(db_path)
                                cur = con.cursor()
                                cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
                                row = cur.fetchone()
                                con.close()
                                nb_id = int(row[0]) if row else None
                            except Exception:
                                nb_id = getattr(window, "_current_notebook_id", None)
                            if nb_id is not None:
                                try:
                                    # Pass parent page_id to expand_page_id so the parent opens to show the new subpage
                                    # Don't pass select_section_id - let select_left_tree_page handle the selection below
                                    ensure_left_tree_sections(window, int(nb_id), expand_page_id=int(page_id))
                                except Exception:
                                    pass
                            # Select the newly created subpage
                            try:
                                from left_tree import select_left_tree_page as _select_left_tree_page

                                if new_pid is not None:
                                    _select_left_tree_page(window, int(section_id), int(new_pid))
                            except Exception:
                                pass
                            # If in two-column UI, load the new page into the editor
                            try:
                                from page_editor import load_page as _load_page_two_column

                                if new_pid is not None and _is_two_column_ui(window):
                                    _load_page_two_column(window, int(new_pid))
                                    try:
                                        from settings_manager import set_last_state

                                        set_last_state(section_id=int(section_id), page_id=int(new_pid))
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                            return
                        if chosen == act_rename_page and page_id is not None:
                            # Prefill current title
                            try:
                                from db_pages import get_page_by_id as db_get_page_by_id

                                row = db_get_page_by_id(int(page_id), db_path)
                                current_title = str(row[2]) if row else ""
                            except Exception:
                                current_title = ""
                            new_title, ok = QtWidgets.QInputDialog.getText(
                                tree, "Rename Page", "New title:", text=current_title
                            )
                            if not ok or not new_title.strip():
                                return
                            try:
                                db_update_page_title(int(page_id), new_title.strip(), db_path)
                            except Exception:
                                pass
                            # Update editor title if this page is active and left-tree label
                            try:
                                if _is_two_column_ui(window):
                                    # Update title field if currently viewing this page
                                    sid_ctx, pid_ctx = _current_page_context(window)
                                    if pid_ctx is not None and int(pid_ctx) == int(page_id):
                                        title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
                                        if title_le is not None:
                                            title_le.blockSignals(True)
                                            title_le.setText(new_title.strip())
                                            title_le.blockSignals(False)
                                    # Update left tree label directly
                                    if section_id is not None:
                                        update_left_tree_page_title(
                                            window, int(section_id), int(page_id), new_title.strip()
                                        )
                            except Exception:
                                pass
                            # Optionally update right pane labels in-place without rebuilding trees
                            try:
                                right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
                                if right_tw is not None and section_id is not None:
                                    for i in range(right_tw.topLevelItemCount()):
                                        sec_item = right_tw.topLevelItem(i)
                                        try:
                                            if int(sec_item.data(0, 1000)) == int(section_id):
                                                for j in range(sec_item.childCount()):
                                                    ch = sec_item.child(j)
                                                    if ch.data(0, 1001) == "page" and int(ch.data(0, 1000)) == int(page_id):
                                                        ch.setText(0, new_title.strip())
                                                        raise StopIteration
                                        except Exception:
                                            pass
                                else:
                                    # Model view path
                                    right_tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
                                    if right_tv is not None and right_tv.model() is not None and section_id is not None:
                                        model = right_tv.model()
                                        from two_pane_core import USER_ROLE_ID, USER_ROLE_KIND
                                        for row in range(model.rowCount()):
                                            idx = model.index(row, 0)
                                            try:
                                                if idx.data(USER_ROLE_KIND) == "section" and int(idx.data(USER_ROLE_ID)) == int(section_id):
                                                    for crow in range(model.rowCount(idx)):
                                                        cidx = model.index(crow, 0, idx)
                                                        if cidx.data(USER_ROLE_KIND) == "page" and int(cidx.data(USER_ROLE_ID)) == int(page_id):
                                                            model.setData(cidx, new_title.strip())
                                                            raise StopIteration
                                            except Exception:
                                                pass
                            except StopIteration:
                                pass
                            except Exception:
                                pass
                            return
                        if chosen == act_delete_page and page_id is not None:
                            # Confirm and delete
                            confirm = QtWidgets.QMessageBox.question(
                                tree,
                                "Delete Page",
                                "Are you sure you want to delete this page?",
                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                            )
                            if confirm != QtWidgets.QMessageBox.Yes:
                                return
                            try:
                                save_current_page(window)
                            except Exception:
                                pass
                            try:
                                db_delete_page(int(page_id), db_path)
                            except Exception:
                                pass
                            # Determine if this was a subpage by checking if parent is a page
                            parent_page_id = None
                            try:
                                parent_item = item.parent()
                                if parent_item is not None and parent_item.data(0, 1001) == "page":
                                    parent_page_id = parent_item.data(0, 1000)
                            except Exception:
                                pass
                            # Two-column: refresh section's children and select parent page or section
                            try:
                                if _is_two_column_ui(window):
                                    # Determine notebook id for this section (always look up from section, not current notebook)
                                    nb_id = None
                                    if section_id is not None:
                                        try:
                                            import sqlite3

                                            con = sqlite3.connect(db_path)
                                            cur = con.cursor()
                                            cur.execute(
                                                "SELECT notebook_id FROM sections WHERE id = ?",
                                                (int(section_id),),
                                            )
                                            row = cur.fetchone()
                                            con.close()
                                            nb_id = int(row[0]) if row else None
                                        except Exception:
                                            nb_id = getattr(window, "_current_notebook_id", None)
                                    if nb_id is not None:
                                        # If this was a subpage, expand the parent page instead of selecting the section
                                        if parent_page_id is not None:
                                            ensure_left_tree_sections(window, int(nb_id), expand_page_id=int(parent_page_id))
                                        else:
                                            ensure_left_tree_sections(
                                                window, int(nb_id), select_section_id=int(section_id) if section_id is not None else None
                                            )
                                    # Select the parent page or load first page
                                    try:
                                        if parent_page_id is not None:
                                            # Select and load the parent page
                                            from left_tree import select_left_tree_page as _select_left_tree_page
                                            from page_editor import load_page as _load_page_two_column
                                            _select_left_tree_page(window, int(section_id), int(parent_page_id))
                                            _load_page_two_column(window, int(parent_page_id))
                                        else:
                                            # Clear current if we deleted the active page, then load first page
                                            if section_id is not None:
                                                sid_int = int(section_id)
                                                cur_pid = getattr(window, "_current_page_by_section", {}).get(sid_int)
                                                if cur_pid is not None and int(cur_pid) == int(page_id):
                                                    window._current_page_by_section[sid_int] = None
                                            _load_first_page_two_column(window)
                                    except Exception:
                                        pass
                                else:
                                    # Legacy: rebuild panes for current notebook
                                    nb_id = getattr(window, "_current_notebook_id", None)
                                    if nb_id is not None:
                                        refresh_for_notebook(
                                            window, int(nb_id), select_section_id=int(section_id) if section_id is not None else None
                                        )
                            except Exception:
                                pass
                            return
                        return
                    else:
                        # Fallback: treat as section
                        chosen = None
                    # Get ids/context
                    section_id = item.data(0, 1000)
                    parent = item.parent()
                    nb_id = parent.data(0, 1000) if parent is not None else None
                    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                    if chosen == act_rename_section and section_id is not None:
                        current_text = item.text(0) or ""
                        new_title, ok = QtWidgets.QInputDialog.getText(
                            tree, "Rename Section", "New title:", text=current_text
                        )
                        if ok and new_title.strip():
                            try:
                                db_rename_section(int(section_id), new_title.strip(), db_path)
                            except Exception:
                                pass
                            # Update UI bits
                            try:
                                item.setText(0, new_title.strip())
                                # Update tab label if visible
                                tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                                if tab_widget:
                                    tab_bar = tab_widget.tabBar()
                                    for i in range(tab_widget.count()):
                                        sid = tab_bar.tabData(i)
                                        if sid == section_id:
                                            tab_widget.setTabText(i, new_title.strip())
                                            break
                            except Exception:
                                pass
                            # Rebuild right pane and keep selection
                            try:
                                if nb_id is not None:
                                    refresh_for_notebook(
                                        window, int(nb_id), select_section_id=int(section_id)
                                    )
                                    ensure_left_tree_sections(
                                        window, int(nb_id), select_section_id=int(section_id)
                                    )
                            except Exception:
                                pass
                        return
                    if chosen == act_delete_section and section_id is not None:
                        sec_name = item.text(0) or "(untitled)"
                        confirm = QtWidgets.QMessageBox.question(
                            tree,
                            "Delete Section",
                            f'Are you sure you want to delete the section "{sec_name}" and all its pages?',
                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                        )
                        if confirm != QtWidgets.QMessageBox.Yes:
                            return
                        # Save any dirty page before delete
                        try:
                            save_current_page(window)
                        except Exception:
                            pass
                        try:
                            db_delete_section(int(section_id), db_path)
                        except Exception:
                            pass
                        # Refresh UI after deletion
                        try:
                            if nb_id is not None:
                                refresh_for_notebook(window, int(nb_id))
                                ensure_left_tree_sections(window, int(nb_id))
                        except Exception:
                            pass
                        return
                except Exception:
                    pass

            # Ensure single connection
            try:
                tree.customContextMenuRequested.disconnect()
            except Exception:
                pass
            tree.customContextMenuRequested.connect(_tree_ctx_menu)

            # Enable drag-and-drop reordering for top-level binders only
            try:
                # Configure tree-wide DnD behavior
                tree.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
                tree.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
                tree.setDefaultDropAction(Qt.MoveAction)
                tree.setAcceptDrops(True)
                tree.setDragEnabled(True)
                tree.setDropIndicatorShown(True)
                # Make drop indicator more pronounced (thicker blue line + subtle fill)
                try:
                    existing = tree.styleSheet() or ""
                    indicator_style = "QTreeView::drop-indicator { border: 2px solid #0078D7; background: rgba(0,120,215,0.25); }"
                    if indicator_style not in existing:
                        tree.setStyleSheet(existing + "\n" + indicator_style)
                except Exception:
                    pass
                # Ensure event filter constrains DnD to top-level and persists order
                if not hasattr(window, "_left_tree_dnd_filter"):
                    from left_tree import LeftTreeDnDFilter

                    window._left_tree_dnd_filter = LeftTreeDnDFilter(window)
                tree.installEventFilter(window._left_tree_dnd_filter)
                if hasattr(tree, "viewport") and tree.viewport() is not None:
                    tree.viewport().installEventFilter(window._left_tree_dnd_filter)
            except Exception:
                pass
    except Exception:
        pass

    # Connect menu actions
    # Updated QAction name from UI: actionNew_Database
    action_newdb = window.findChild(QtWidgets.QAction, "actionNew_Database")
    if action_newdb:
        action_newdb.triggered.connect(lambda: create_new_database(window))
    # Binder (notebook) actions
    act_add_wb_variants = [
        window.findChild(QtWidgets.QAction, "actionAdd_WorkBook"),
        window.findChild(QtWidgets.QAction, "actionAdd_Workbook"),
    ]
    for act in act_add_wb_variants:
        if act:
            act.triggered.connect(lambda: add_binder(window))
    act_rename_wb = window.findChild(QtWidgets.QAction, "actionRename_WorkBook")
    if act_rename_wb:
        act_rename_wb.triggered.connect(lambda: rename_binder(window))
    act_delete_wb = window.findChild(QtWidgets.QAction, "actionDelete_Workbook")
    if act_delete_wb:
        act_delete_wb.triggered.connect(lambda: delete_binder(window))
    # File menu: Open
    action_open = window.findChild(QtWidgets.QAction, "actionOpen")
    if action_open:
        action_open.triggered.connect(lambda: open_database(window))
        # Add Ctrl+O shortcut
        from PyQt5.QtGui import QKeySequence
        action_open.setShortcut(QKeySequence.Open)  # Ctrl+O
    
    # File menu: Save (saves current page)
    action_save = window.findChild(QtWidgets.QAction, "actionSave")
    if action_save:
        from PyQt5.QtGui import QKeySequence
        action_save.setShortcut(QKeySequence.Save)  # Ctrl+S
        action_save.triggered.connect(lambda: save_current_page(window))
    
    # File menu: Save As (copy database to new location)
    action_save_as = window.findChild(QtWidgets.QAction, "actionSave_As")
    if action_save_as:
        action_save_as.triggered.connect(lambda: save_database_as(window))
        # Add Ctrl+Shift+S shortcut (standard for Save As)
        from PyQt5.QtGui import QKeySequence
        action_save_as.setShortcut(QKeySequence.SaveAs)  # Ctrl+Shift+S
    
    # File menu: Print (print selected page/section/binder)
    action_print = window.findChild(QtWidgets.QAction, "actionPrint")
    if action_print:
        from PyQt5.QtGui import QKeySequence
        action_print.setShortcut(QKeySequence.Print)  # Ctrl+P
        action_print.triggered.connect(lambda: print_current_selection(window))
    
    # File menu: Print Preview (show preview before printing)
    action_print_preview = window.findChild(QtWidgets.QAction, "actionPrint_Preview")
    if action_print_preview:
        from PyQt5.QtGui import QKeySequence
        action_print_preview.setShortcut(QKeySequence("Ctrl+Shift+P"))
        action_print_preview.triggered.connect(lambda: print_preview_current_selection(window))
    # Insert menu wiring for quick content creation
    act_add_section = window.findChild(QtWidgets.QAction, "actionAdd_Scction")
    if act_add_section:
        act_add_section.triggered.connect(lambda: add_section(window))
    act_add_page = window.findChild(QtWidgets.QAction, "actionAdd_Page")
    if act_add_page:
        act_add_page.triggered.connect(lambda: add_page(window))
    # Insert menu: Collapse All Binders
    act_collapse_all = window.findChild(QtWidgets.QAction, "actionCollapse_All_Binders")
    if act_collapse_all:

        def _collapse_all_binders():
            try:
                tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
                if tree is None:
                    return
                for i in range(tree.topLevelItemCount()):
                    top = tree.topLevelItem(i)
                    top.setExpanded(False)
                try:
                    from settings_manager import set_expanded_notebooks

                    set_expanded_notebooks(set())
                except Exception:
                    pass
            except Exception:
                pass

        act_collapse_all.triggered.connect(_collapse_all_binders)
    # Insert menu: Binder ops duplicates
    act_del_wb_ins = window.findChild(QtWidgets.QAction, "actionDelete_Workbook")
    if act_del_wb_ins:
        act_del_wb_ins.triggered.connect(lambda: delete_binder(window))
    act_ren_wb_ins = window.findChild(QtWidgets.QAction, "actionRename_WorkBook")
    if act_ren_wb_ins:
        act_ren_wb_ins.triggered.connect(lambda: rename_binder(window))
    # Insert menu: Section ops
    act_del_sec = window.findChild(QtWidgets.QAction, "actionDelete_Section")
    if act_del_sec:

        def _del_section_from_menu():
            try:
                tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
                item = tree_widget.currentItem() if tree_widget else None
                if item is None:
                    return
                # If binder selected, try first child section
                if item.parent() is None:
                    # pick selected tab's section instead
                    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                    if tab_widget and tab_widget.count() > 0:
                        sid = tab_widget.tabBar().tabData(tab_widget.currentIndex())
                        current_name = tab_widget.tabText(tab_widget.currentIndex()) or "(untitled)"
                    else:
                        sid = None
                        current_name = "(untitled)"
                else:
                    sid = item.data(0, 1000)
                    current_name = item.text(0) or "(untitled)"
                if sid is None:
                    return
                confirm = QtWidgets.QMessageBox.question(
                    window,
                    "Delete Section",
                    f'Are you sure you want to delete the section "{current_name}" and all its pages?',
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                )
                if confirm != QtWidgets.QMessageBox.Yes:
                    return
                try:
                    save_current_page(window)
                except Exception:
                    pass
                db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                db_delete_section(int(sid), db_path)
                nb_id = getattr(window, "_current_notebook_id", None)
                if nb_id is not None:
                    refresh_for_notebook(window, int(nb_id))
                    ensure_left_tree_sections(window, int(nb_id))
            except Exception:
                pass

        act_del_sec.triggered.connect(_del_section_from_menu)
    act_ren_sec = window.findChild(QtWidgets.QAction, "actionRename_Section")
    if act_ren_sec:

        def _ren_section_from_menu():
            try:
                tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
                item = tree_widget.currentItem() if tree_widget else None
                # Prefer selected section; else active tab section
                sid = None
                if item is not None and item.parent() is not None:
                    sid = item.data(0, 1000)
                    current_text = item.text(0) or ""
                else:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                    if tab_widget and tab_widget.count() > 0:
                        idx = tab_widget.currentIndex()
                        sid = tab_widget.tabBar().tabData(idx)
                        current_text = tab_widget.tabText(idx) or ""
                if sid is None:
                    return
                new_title, ok = QtWidgets.QInputDialog.getText(
                    window, "Rename Section", "New title:", text=current_text
                )
                if not ok or not new_title.strip():
                    return
                db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                db_rename_section(int(sid), new_title.strip(), db_path)
                nb_id = getattr(window, "_current_notebook_id", None)
                if nb_id is not None:
                    refresh_for_notebook(window, int(nb_id), select_section_id=int(sid))
                    ensure_left_tree_sections(window, int(nb_id), select_section_id=int(sid))
            except Exception:
                pass

        act_ren_sec.triggered.connect(_ren_section_from_menu)

    # --- Tools / Maintenance menu: Normalize Page Order ---
    try:
        menubar = window.menuBar()
        tools_menu = None
        # Try to find existing 'Tools' or 'Maintenance' menu
        for act in menubar.actions():
            if act.menu() and act.text().strip().lower() in {"tools", "maintenance"}:
                tools_menu = act.menu()
                break
        if tools_menu is None:
            tools_menu = menubar.addMenu("Tools")
        normalize_action = QtWidgets.QAction("Normalize Page Order", window)
        normalize_action.setToolTip("Resequence order_index values (gap‑free) for all notebooks, sections, and pages")

        def _normalize_order_indexes():
            try:
                from maintenance_order import collect_changes, summarize, apply_changes
            except Exception as e:
                QtWidgets.QMessageBox.warning(window, "Normalize", f"Maintenance module missing: {e}")
                return
            # Flush current edits first
            try:
                save_current_page(window)
            except Exception:
                pass
            db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
            try:
                changes = collect_changes(db_path)
            except Exception as e:
                QtWidgets.QMessageBox.warning(window, "Normalize", f"Failed to collect changes: {e}")
                return
            total = sum(len(changes[k]) for k in changes)
            summary = summarize(changes)
            if total == 0:
                QtWidgets.QMessageBox.information(window, "Normalize Page Order", f"Already normalized.\n\n{summary}")
                return
            # Offer backup + apply
            msg_box = QtWidgets.QMessageBox(window)
            msg_box.setWindowTitle("Normalize Page Order")
            msg_box.setIcon(QtWidgets.QMessageBox.Question)
            msg_box.setText("Proposed resequencing:\n\n" + summary + "\n\nCreate backup and apply normalization?")
            backup_apply = msg_box.addButton("Backup && Apply", QtWidgets.QMessageBox.AcceptRole)
            apply_only = msg_box.addButton("Apply Only", QtWidgets.QMessageBox.DestructiveRole)
            cancel_btn = msg_box.addButton(QtWidgets.QMessageBox.Cancel)
            msg_box.exec_()
            chosen = msg_box.clickedButton()
            if chosen == cancel_btn:
                return
            # Optional backup
            if chosen == backup_apply:
                try:
                    import os, shutil, datetime
                    stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                    base_dir = os.path.dirname(db_path) or "."
                    backup_name = f"notes_pre_normalize_menu_{stamp}.db"
                    backup_path = os.path.join(base_dir, backup_name)
                    shutil.copy2(db_path, backup_path)
                except Exception as e:
                    QtWidgets.QMessageBox.warning(window, "Normalize", f"Backup failed (continuing): {e}")
            # Apply changes
            try:
                apply_changes(db_path, changes)
            except Exception as e:
                QtWidgets.QMessageBox.critical(window, "Normalize", f"Failed to apply updates: {e}")
                return
            # Refresh current notebook tree & editor context
            try:
                nb_id = getattr(window, "_current_notebook_id", None)
                if nb_id is not None:
                    ensure_left_tree_sections(window, int(nb_id))
                    refresh_for_notebook(window, int(nb_id))
            except Exception:
                pass
            QtWidgets.QMessageBox.information(window, "Normalize Page Order", "Normalization complete.\n\n" + summarize(collect_changes(db_path)))

        normalize_action.triggered.connect(_normalize_order_indexes)
        tools_menu.addAction(normalize_action)
        # (Legacy formula actions removed during feature rollback.)
        
        # --- Spell Check toggle ---
        tools_menu.addSeparator()
        spell_check_action = QtWidgets.QAction("Spell Check", window)
        spell_check_action.setCheckable(True)
        try:
            from settings_manager import get_spell_check_enabled
            from spell_check import is_spell_check_available
            spell_available = is_spell_check_available()
            spell_check_action.setEnabled(spell_available)
            if spell_available:
                spell_check_action.setChecked(get_spell_check_enabled())
            else:
                spell_check_action.setChecked(False)
                spell_check_action.setToolTip("Spell check unavailable (pyenchant not installed)")
        except Exception:
            spell_check_action.setChecked(False)
        
        def _toggle_spell_check(checked):
            try:
                from settings_manager import set_spell_check_enabled
                set_spell_check_enabled(checked)
                # Toggle the spell checker on the editor
                spell_checker = getattr(window, "_spell_checker", None)
                if spell_checker:
                    spell_checker.enabled = checked
            except Exception:
                pass
        
        spell_check_action.triggered.connect(_toggle_spell_check)
        tools_menu.addAction(spell_check_action)
    except Exception:
        pass
    # Insert menu: Page ops
    act_del_page = window.findChild(QtWidgets.QAction, "actionDelete_Page")
    if act_del_page:

        def _del_page_from_menu():
            try:
                # Determine current page from active section
                section_id, page_id = _current_page_context(window)
                if page_id is None:
                    QtWidgets.QMessageBox.information(
                        window, "Delete Page", "Please select a page to delete."
                    )
                    return
                confirm = QtWidgets.QMessageBox.question(
                    window,
                    "Delete Page",
                    "Are you sure you want to delete this page?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                )
                if confirm != QtWidgets.QMessageBox.Yes:
                    return
                db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                db_delete_page(int(page_id), db_path)
                # Refresh UI
                try:
                    if _is_two_column_ui(window):
                        # If we deleted the active page for this section, clear mapping
                        try:
                            sid_int = int(section_id) if section_id is not None else None
                            if sid_int is not None:
                                cur_pid = getattr(window, "_current_page_by_section", {}).get(
                                    sid_int
                                )
                                if cur_pid == int(page_id):
                                    window._current_page_by_section[sid_int] = None
                        except Exception:
                            pass
                        # Refresh left tree children for this section's notebook and load first page
                        try:
                            # Determine notebook id for this section
                            import sqlite3

                            con = sqlite3.connect(db_path)
                            cur = con.cursor()
                            cur.execute(
                                "SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),)
                            )
                            row = cur.fetchone()
                            con.close()
                            nb_id = (
                                int(row[0])
                                if row
                                else getattr(window, "_current_notebook_id", None)
                            )
                        except Exception:
                            nb_id = getattr(window, "_current_notebook_id", None)
                        if nb_id is not None:
                            ensure_left_tree_sections(
                                window, int(nb_id), select_section_id=int(section_id)
                            )
                        try:
                            window._current_section_id = int(section_id)
                        except Exception:
                            window._current_section_id = section_id
                            _load_first_page_two_column(window)
                        except Exception:
                            pass
                    else:
                        # Legacy tabs: rebuild panes for current notebook
                        nb_id = getattr(window, "_current_notebook_id", None)
                        if nb_id is not None:
                            refresh_for_notebook(
                                window, int(nb_id), select_section_id=int(section_id)
                            )
                except Exception:
                    # Fallback to legacy refresh
                    nb_id = getattr(window, "_current_notebook_id", None)
                    if nb_id is not None:
                        refresh_for_notebook(window, int(nb_id), select_section_id=int(section_id))
            except Exception:
                pass

        act_del_page.triggered.connect(_del_page_from_menu)
    act_ren_page = window.findChild(QtWidgets.QAction, "actionRename_Page")
    if act_ren_page:

        def _ren_page_from_menu():
            try:
                section_id, page_id = _current_page_context(window)
                if page_id is None:
                    QtWidgets.QMessageBox.information(
                        window, "Rename Page", "Please select a page to rename."
                    )
                    return
                # Prefill current title
                try:
                    from db_pages import get_page_by_id as db_get_page_by_id

                    row = db_get_page_by_id(
                        int(page_id),
                        getattr(window, "_db_path", None) or get_last_db() or "notes.db",
                    )
                    current_title = str(row[2]) if row else ""
                except Exception:
                    current_title = ""
                new_title, ok = QtWidgets.QInputDialog.getText(
                    window, "Rename Page", "New title:", text=current_title
                )
                if not ok or not new_title.strip():
                    return
                db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                db_update_page_title(int(page_id), new_title.strip(), db_path)
                # Reflect in UI: update title field (2-col) and left tree label
                try:
                    if _is_two_column_ui(window):
                        try:
                            title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
                            if title_le is not None:
                                title_le.blockSignals(True)
                                title_le.setText(new_title.strip())
                                title_le.blockSignals(False)
                        except Exception:
                            pass
                        try:
                            # Update left tree item text without full rebuild
                            update_left_tree_page_title(
                                window, int(section_id), int(page_id), new_title.strip()
                            )
                        except Exception:
                            # Fallback: refresh left tree for section's notebook
                            try:
                                # Lookup notebook id for section to refresh children
                                import sqlite3

                                con = sqlite3.connect(db_path)
                                cur = con.cursor()
                                cur.execute(
                                    "SELECT notebook_id FROM sections WHERE id = ?",
                                    (int(section_id),),
                                )
                                row = cur.fetchone()
                                con.close()
                                nb_id = int(row[0]) if row else None
                            except Exception:
                                nb_id = getattr(window, "_current_notebook_id", None)
                            if nb_id is not None:
                                ensure_left_tree_sections(
                                    window, int(nb_id), select_section_id=int(section_id)
                                )
                except Exception:
                    pass
                # Legacy/tabbed: refresh right/left as before
                nb_id = getattr(window, "_current_notebook_id", None)
                if nb_id is not None:
                    try:
                        refresh_for_notebook(window, int(nb_id), select_section_id=int(section_id))
                    except Exception:
                        pass
            except Exception:
                pass

        act_ren_page.triggered.connect(_ren_page_from_menu)
    act_insert_attachment = window.findChild(QtWidgets.QAction, "actionInsert_Attachment")
    if act_insert_attachment:
        act_insert_attachment.triggered.connect(lambda: insert_attachment(window))
    # Tools: Manual Database Backup
    try:
        act_backup_now = window.findChild(QtWidgets.QAction, "actionBackup_Database")
        if act_backup_now is not None:
            act_backup_now.triggered.connect(lambda: backup_database_now(window))
    except Exception:
        pass
    # Tools: Rename Database (handled in backup module for compartmentalization)
    try:
        act_rename_db = window.findChild(QtWidgets.QAction, "actionRename_Database")
        if act_rename_db is not None:
            from backup import show_rename_database_dialog

            act_rename_db.triggered.connect(lambda: show_rename_database_dialog(window))
    except Exception:
        pass
    # File: Export Binder (handled in backup module)
    try:
        act_export_binder = window.findChild(QtWidgets.QAction, "actionExport_Binder")
        if act_export_binder is not None:
            from backup import export_binder

            act_export_binder.triggered.connect(lambda: export_binder(window))
    except Exception:
        pass
    # File: Import Binder (handled in backup module)
    try:
        act_import_binder = window.findChild(QtWidgets.QAction, "actionImport_Binder")
        if act_import_binder is not None:
            from backup import import_binder

            act_import_binder.triggered.connect(lambda: import_binder(window))
    except Exception:
        pass
    # Insert menu: Planning Register
    act_plan_reg = window.findChild(QtWidgets.QAction, "actionPlanning_Register")
    if act_plan_reg:
        def _insert_planning_register_via_dialog():
            try:
                te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                if te is None or not te.isEnabled():
                    QtWidgets.QMessageBox.information(window, "Insert Planning Register", "Please open or create a page first.")
                    return
                # Build options: first 'New Planning Register', then saved presets (if any)
                try:
                    from settings_manager import list_table_preset_names
                    preset_names = list_table_preset_names()
                except Exception:
                    preset_names = []
                options = ["New Planning Register"] + preset_names
                choice, ok = QtWidgets.QInputDialog.getItem(
                    window, "Insert Planning Register", "Choose:", options, 0, False
                )
                if not (ok and choice):
                    return
                if choice == "New Planning Register":
                    insert_planning_register(window)
                else:
                    insert_table_from_preset(te, choice, fit_width_100=True)
            except Exception:
                pass
        act_plan_reg.triggered.connect(_insert_planning_register_via_dialog)
    # Save Planning Register as Preset (Insert menu)
    try:
        act_save_reg_preset = window.findChild(QtWidgets.QAction, "actionSave_Planning_Register_as_Preset")
        act_rename_reg_preset = window.findChild(QtWidgets.QAction, "actionRename_Planning_Register_Preset")
        act_delete_reg_preset = window.findChild(QtWidgets.QAction, "actionDelete_Planning_Register_Preset")
        if act_save_reg_preset is not None:
            def _save_planning_register_as_preset():
                te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                if te is None:
                    return
                cur = te.textCursor()
                tbl = cur.currentTable()
                if tbl is None:
                    QtWidgets.QMessageBox.information(window, "Save Planning Register as Preset", "Place the caret inside the left Planning Register table.")
                    return
                # If the caret is on the outer container, try to descend into the left cell's inner table
                try:
                    if tbl.rows() == 1 and tbl.columns() == 2:
                        left_cell = tbl.cellAt(0, 0)
                        s_pos = left_cell.firstCursorPosition().position()
                        e_pos = left_cell.lastCursorPosition().position()
                        from PyQt5.QtGui import QTextCursor as _QTextCursor
                        scan = _QTextCursor(te.document())
                        scan.setPosition(s_pos)
                        inner = None
                        iters = 0
                        while scan.position() < e_pos and iters < 20000:
                            t = scan.currentTable()
                            if t is not None:
                                inner = t
                                break
                            scan.movePosition(_QTextCursor.NextCharacter)
                            iters += 1
                        if inner is not None:
                            te.setTextCursor(inner.firstCursorPosition())
                            tbl = inner
                except Exception:
                    pass
                # Verify this looks like a Planning Register table (left table)
                try:
                    from ui_planning_register import _is_planning_register_table
                    if not _is_planning_register_table(te, tbl):
                        QtWidgets.QMessageBox.information(window, "Save Planning Register as Preset", "Please place the caret inside the left Planning Register table.")
                        return
                except Exception:
                    pass
                # Use the centralized HTML preset saver
                from ui_richtext import save_current_table_as_preset
                save_current_table_as_preset(te)

            act_save_reg_preset.triggered.connect(_save_planning_register_as_preset)

        # Helper to choose a preset name
        def _choose_preset_name(parent, title: str) -> str:
            try:
                from settings_manager import list_table_preset_names
                names = list_table_preset_names()
            except Exception:
                names = []
            if not names:
                QtWidgets.QMessageBox.information(parent, title, "No presets saved yet.")
                return None
            item, ok = QtWidgets.QInputDialog.getItem(parent, title, "Preset:", names, 0, False)
            return item if ok and item else None

        if act_rename_reg_preset is not None:
            def _rename_planning_register_preset():
                name = _choose_preset_name(window, "Rename Planning Register Preset")
                if not name:
                    return
                new_name, ok = QtWidgets.QInputDialog.getText(window, "Rename Preset", "New name:", text=name)
                if not ok or not new_name or new_name == name:
                    return
                try:
                    from settings_manager import rename_table_preset
                    rename_table_preset(name, new_name)
                except Exception:
                    pass
            act_rename_reg_preset.triggered.connect(_rename_planning_register_preset)

        if act_delete_reg_preset is not None:
            def _delete_planning_register_preset():
                name = _choose_preset_name(window, "Delete Planning Register Preset")
                if not name:
                    return
                if QtWidgets.QMessageBox.question(
                    window, "Delete Preset", f"Delete preset '{name}'?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                ) != QtWidgets.QMessageBox.Yes:
                    return
                try:
                    from settings_manager import delete_table_preset
                    delete_table_preset(name)
                except Exception:
                    pass
            act_delete_reg_preset.triggered.connect(_delete_planning_register_preset)

    except Exception:
        pass

    # --- Show Deleted Items toggle and Empty All Deleted Items (File menu, before Exit) ---
    try:
        menubar = window.menuBar()
        file_menu = None
        for act in menubar.actions():
            if act.menu() and act.text().replace("&", "").strip().lower() == "file":
                file_menu = act.menu()
                break
        if file_menu is not None:
            # Find the Exit action to insert before it
            exit_action = window.findChild(QtWidgets.QAction, "actionExit")
            
            # Create separator
            sep_action = QtWidgets.QAction(window)
            sep_action.setSeparator(True)
            
            # Show Deleted Items - checkable action
            show_deleted_action = QtWidgets.QAction("Show Deleted Items", window)
            show_deleted_action.setCheckable(True)
            try:
                from settings_manager import get_show_deleted
                show_deleted_action.setChecked(get_show_deleted())
            except Exception:
                show_deleted_action.setChecked(False)
            # Store on window for syncing with context menus
            window._show_deleted_action = show_deleted_action
            
            def _toggle_show_deleted(checked):
                try:
                    from settings_manager import set_show_deleted
                    set_show_deleted(checked)
                except Exception:
                    pass
                # Refresh the tree to show/hide deleted items
                try:
                    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                    populate_notebook_names(window, db_path)
                    # Re-expand current notebook if any
                    nb_id = getattr(window, "_current_notebook_id", None)
                    if nb_id is not None:
                        ensure_left_tree_sections(window, int(nb_id))
                except Exception:
                    pass
            
            show_deleted_action.triggered.connect(_toggle_show_deleted)
            
            # Empty All Deleted Items
            empty_deleted_action = QtWidgets.QAction("Empty All Deleted Items...", window)
            
            def _empty_all_deleted():
                try:
                    from db_access import get_deleted_counts, empty_all_deleted
                    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                    counts = get_deleted_counts(db_path)
                    if counts['total'] == 0:
                        QtWidgets.QMessageBox.information(
                            window, "Empty Deleted Items", "No deleted items to remove."
                        )
                        return
                    # Confirm before permanent deletion
                    msg = (
                        f"This will permanently delete:\n\n"
                        f"  • {counts['notebooks']} binder(s)\n"
                        f"  • {counts['sections']} section(s)\n"
                        f"  • {counts['pages']} page(s)\n\n"
                        f"This action cannot be undone. Continue?"
                    )
                    confirm = QtWidgets.QMessageBox.warning(
                        window,
                        "Empty All Deleted Items",
                        msg,
                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    )
                    if confirm != QtWidgets.QMessageBox.Yes:
                        return
                    empty_all_deleted(db_path)
                    # Refresh tree
                    populate_notebook_names(window, db_path)
                    nb_id = getattr(window, "_current_notebook_id", None)
                    if nb_id is not None:
                        ensure_left_tree_sections(window, int(nb_id))
                    QtWidgets.QMessageBox.information(
                        window, "Empty Deleted Items", "All deleted items have been permanently removed."
                    )
                except Exception as e:
                    QtWidgets.QMessageBox.warning(window, "Error", f"Failed to empty deleted items: {e}")
            
            empty_deleted_action.triggered.connect(_empty_all_deleted)
            
            # Insert before Exit action (or append if Exit not found)
            if exit_action is not None:
                file_menu.insertAction(exit_action, sep_action)
                file_menu.insertAction(exit_action, show_deleted_action)
                file_menu.insertAction(exit_action, empty_deleted_action)
                # Add another separator before Exit
                sep_before_exit = QtWidgets.QAction(window)
                sep_before_exit.setSeparator(True)
                file_menu.insertAction(exit_action, sep_before_exit)
            else:
                file_menu.addSeparator()
                file_menu.addAction(show_deleted_action)
                file_menu.addAction(empty_deleted_action)
    except Exception:
        pass
    
    action_exit = window.findChild(QtWidgets.QAction, "actionExit")
    if action_exit:
        action_exit.triggered.connect(window.close)

    # Build/augment a 'Table Presets' submenu under Insert (or reuse one from the .ui to avoid duplicates)
    try:
        menubar = window.menuBar() if hasattr(window, "menuBar") else None
        target_menu = None

        def _find_or_create_table_presets_menu(win) -> QtWidgets.QMenu:
            mb = win.menuBar() if hasattr(win, "menuBar") else None
            if mb is None:
                return None
            # Try TOP-LEVEL first (unlikely for this app but supported)
            for act in mb.actions():
                m = act.menu()
                if m and act.text().replace("&", "").strip().lower() == "table presets":
                    return m
            # Find the Insert menu
            insert_m = None
            for act in mb.actions():
                m = act.menu()
                if m and act.text().replace("&", "").strip().lower() == "insert":
                    insert_m = m
                    break
            if insert_m is None:
                return None
            # If a 'Table Presets' submenu already exists in Insert (defined in .ui), reuse it
            for act in insert_m.actions():
                m = act.menu()
                if m and act.text().replace("&", "").strip().lower() == "table presets":
                    return m
            # Otherwise create it under Insert
            return insert_m.addMenu("Table Presets")

        # If the UI provides explicit actions for presets, wire those and skip creating a separate submenu
        act_insert_preset = window.findChild(QtWidgets.QAction, "actionInsert_Table_Preset")
        act_save_preset = window.findChild(QtWidgets.QAction, "actionSave_Table_as_Preset")
        if act_insert_preset:
            from ui_richtext import choose_and_insert_preset
            act_insert_preset.triggered.connect(lambda: choose_and_insert_preset(window.findChild(QtWidgets.QTextEdit, "pageEdit"), fit_width_100=True))
        if act_save_preset:
            from ui_richtext import save_current_table_as_preset
            act_save_preset.triggered.connect(lambda: save_current_table_as_preset(window.findChild(QtWidgets.QTextEdit, "pageEdit")))

        # Only create a Table Presets submenu if we don't have explicit actions in the UI
        target_menu = None if (act_insert_preset or act_save_preset) else _find_or_create_table_presets_menu(window)
        if target_menu is not None:
            target_menu.clear()
            # Insert submenu
            sub_insert = target_menu.addMenu("Insert Preset")
            try:
                from settings_manager import list_table_preset_names

                names = list_table_preset_names()
            except Exception:
                names = []
            if names:
                for nm in names:
                    act = sub_insert.addAction(nm)
                    act.triggered.connect(lambda chk=False, name=nm: _insert_preset_into_editor(window, name))
            else:
                sub_insert.setEnabled(False)
            target_menu.addSeparator()
            act_ren = target_menu.addAction("Rename Preset…")
            act_del = target_menu.addAction("Delete Preset…")

            def _choose_preset_name(parent, title: str) -> str:
                try:
                    from settings_manager import list_table_preset_names

                    names = list_table_preset_names()
                except Exception:
                    names = []
                if not names:
                    QtWidgets.QMessageBox.information(parent, title, "No presets saved yet.")
                    return None
                item, ok = QtWidgets.QInputDialog.getItem(parent, title, "Preset:", names, 0, False)
                return item if ok and item else None

            def _rename_preset():
                name = _choose_preset_name(window, "Rename Preset")
                if not name:
                    return
                new_name, ok = QtWidgets.QInputDialog.getText(window, "Rename Preset", "New name:", text=name)
                if not ok or not new_name or new_name == name:
                    return
                try:
                    from settings_manager import rename_table_preset

                    rename_table_preset(name, new_name)
                except Exception:
                    pass
                # Rebuild the menu to reflect the change
                QTimer.singleShot(0, lambda: _rebuild_table_presets_menu(window))

            def _delete_preset():
                name = _choose_preset_name(window, "Delete Preset")
                if not name:
                    return
                if QtWidgets.QMessageBox.question(
                    window, "Delete Preset", f"Delete preset '{name}'?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                ) != QtWidgets.QMessageBox.Yes:
                    return
                try:
                    from settings_manager import delete_table_preset

                    delete_table_preset(name)
                except Exception:
                    pass
                QTimer.singleShot(0, lambda: _rebuild_table_presets_menu(window))

            act_ren.triggered.connect(_rename_preset)
            act_del.triggered.connect(_delete_preset)
    except Exception:
        pass

    # Helper to rebuild the Table Presets menu dynamically
    def _rebuild_table_presets_menu(win):
        try:
            # Re-enter main() portion just to rebuild this menu block
            menubar = win.menuBar() if hasattr(win, "menuBar") else None
            if menubar is None:
                return
            # If UI provides actions, nothing to rebuild here
            if win.findChild(QtWidgets.QAction, "actionInsert_Table_Preset") or win.findChild(QtWidgets.QAction, "actionSave_Table_as_Preset"):
                return
            # Find the Table Presets menu either as a top-level entry or under Insert
            target_menu = None
            # Top-level
            for a in menubar.actions():
                if a.menu() and a.text().replace("&", "").strip().lower() == "table presets":
                    target_menu = a.menu()
                    break
            if target_menu is None:
                # Under Insert
                insert_menu = None
                for a in menubar.actions():
                    if a.menu() and a.text().replace("&", "").strip().lower() == "insert":
                        insert_menu = a.menu()
                        break
                if insert_menu is not None:
                    for act in insert_menu.actions():
                        m = act.menu()
                        if m and act.text().replace("&", "").strip().lower() == "table presets":
                            target_menu = m
                            break
            if target_menu is None:
                return
            # Rebuild contents
            target_menu.clear()
            sub_insert = target_menu.addMenu("Insert Preset")
            try:
                from settings_manager import list_table_preset_names
                names = list_table_preset_names()
            except Exception:
                names = []
            if names:
                for nm in names:
                    act = sub_insert.addAction(nm)
                    act.triggered.connect(lambda chk=False, name=nm: _insert_preset_into_editor(win, name))
            else:
                sub_insert.setEnabled(False)
            target_menu.addSeparator()
            act_ren = target_menu.addAction("Rename Preset…")
            act_del = target_menu.addAction("Delete Preset…")
            def _rename_preset_local():
                from settings_manager import list_table_preset_names, rename_table_preset
                name, ok = QtWidgets.QInputDialog.getItem(win, "Rename Preset", "Preset:", list_table_preset_names(), 0, False)
                if not ok or not name:
                    return
                new_name, ok2 = QtWidgets.QInputDialog.getText(win, "Rename Preset", "New name:", text=name)
                if not ok2 or not new_name or new_name == name:
                    return
                rename_table_preset(name, new_name)
                QTimer.singleShot(0, lambda: _rebuild_table_presets_menu(win))
            def _delete_preset_local():
                from settings_manager import list_table_preset_names, delete_table_preset
                name, ok = QtWidgets.QInputDialog.getItem(win, "Delete Preset", "Preset:", list_table_preset_names(), 0, False)
                if not ok or not name:
                    return
                if QtWidgets.QMessageBox.question(win, "Delete Preset", f"Delete preset '{name}'?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No) != QtWidgets.QMessageBox.Yes:
                    return
                delete_table_preset(name)
                QTimer.singleShot(0, lambda: _rebuild_table_presets_menu(win))
            act_ren.triggered.connect(_rename_preset_local)
            act_del.triggered.connect(_delete_preset_local)
        except Exception:
            pass

    def _insert_preset_into_editor(win, name: str):
        try:
            te = win.findChild(QtWidgets.QTextEdit, "pageEdit")
            if te is not None:
                insert_table_from_preset(te, name, fit_width_100=True)
        except Exception:
            pass

    # Keyboard: Ctrl+Up / Ctrl+Down to reorder binders (top-level notebooks)
    try:

        def _move_binder(delta: int):
            try:
                tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
                if tree is None or tree.topLevelItemCount() == 0:
                    return
                cur = tree.currentItem()
                # If a section is selected, operate on its parent binder
                if cur is not None and cur.parent() is not None:
                    cur = cur.parent()
                if cur is None or cur.parent() is not None:
                    return
                idx = tree.indexOfTopLevelItem(cur)
                if idx < 0:
                    return
                new_idx = idx + (1 if delta > 0 else -1)
                if new_idx < 0 or new_idx >= tree.topLevelItemCount():
                    return
                # Cache identifiers and db path up front
                moved_id = cur.data(0, 1000)
                db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                # Build current order of ids
                ordered_ids = []
                for i in range(tree.topLevelItemCount()):
                    nid = tree.topLevelItem(i).data(0, 1000)
                    if nid is not None:
                        ordered_ids.append(int(nid))
                # Swap positions
                ordered_ids[idx], ordered_ids[new_idx] = ordered_ids[new_idx], ordered_ids[idx]
                # Persist order
                try:
                    from db_access import set_notebooks_order

                    set_notebooks_order(ordered_ids, db_path)
                except Exception:
                    pass
                # Repopulate left tree and reselect the moved binder
                try:
                    from ui_logic import populate_notebook_names

                    populate_notebook_names(window, db_path)
                    # Reselect by id
                    if moved_id is not None:
                        for i in range(tree.topLevelItemCount()):
                            top = tree.topLevelItem(i)
                            if int(top.data(0, 1000)) == int(moved_id):
                                tree.setCurrentItem(top)
                                break
                    # Keep UI in sync without changing binder selection
                    try:
                        window._current_notebook_id = int(moved_id)
                        from settings_manager import get_expanded_notebooks

                        # Restore expanded state after repopulate
                        expanded_ids = get_expanded_notebooks()
                        for i in range(tree.topLevelItemCount()):
                            top = tree.topLevelItem(i)
                            tid = top.data(0, 1000)
                            try:
                                tid_int = int(tid)
                            except Exception:
                                tid_int = None
                            if tid_int is not None and tid_int in expanded_ids:
                                top.setExpanded(True)
                        # Refresh center to reflect current binder context
                        refresh_for_notebook(window, int(moved_id), keep_left_tree_selection=True)
                        # Ensure focus stays on the left tree so repeated Ctrl+Up/Down works
                        try:
                            tree.setFocus(Qt.OtherFocusReason)
                        except Exception:
                            pass
                    except Exception:
                        pass
                except Exception:
                    pass
            except Exception:
                pass

        # Bind shortcuts on the LEFT TREE ONLY so right-panel Ctrl+Up/Down won't move binders
        from PyQt5.QtGui import QKeySequence

        _left_tree_for_shortcuts = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if _left_tree_for_shortcuts is not None:
            sc_up = QtWidgets.QShortcut(
                QKeySequence("Ctrl+Up"),
                _left_tree_for_shortcuts,
                activated=lambda: _move_binder(-1),
            )
            sc_down = QtWidgets.QShortcut(
                QKeySequence("Ctrl+Down"),
                _left_tree_for_shortcuts,
                activated=lambda: _move_binder(1),
            )
            try:
                sc_up.setContext(Qt.WidgetWithChildrenShortcut)
                sc_down.setContext(Qt.WidgetWithChildrenShortcut)
            except Exception:
                pass
            # Keep refs
            window._binder_move_shortcuts = [sc_up, sc_down]
    except Exception:
        pass

    # Keyboard dispatch for right panel: Ctrl+Up / Ctrl+Down moves Section or Page based on selection
    def _right_panel_move(delta: int):
        try:
            # Try QTreeWidget first
            right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
            if right_tw is not None:
                cur = right_tw.currentItem()
                if cur is not None:
                    kind = cur.data(0, 1001)
                    if kind == "section":
                        _move_section(delta)
                        return
                    if kind == "page":
                        _move_page(delta)
                        return
            # Fallback to QTreeView
            right_tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
            if right_tv is not None:
                idx = right_tv.currentIndex()
                if idx.isValid():
                    kind = idx.data(1001)
                    if kind == "section":
                        _move_section(delta)
                        return
                    if kind == "page":
                        _move_page(delta)
                        return
        except Exception:
            pass

    # Keyboard: Ctrl+Up / Ctrl+Down to reorder SECTIONS in the right panel (when a section is selected)
    try:

        def _move_section(delta: int):
            try:
                nb_id = getattr(window, "_current_notebook_id", None)
                if nb_id is None:
                    return
                # Prefer QTreeWidget path
                right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
                right_tv = (
                    window.findChild(QtWidgets.QTreeView, "sectionPages")
                    if right_tw is None
                    else None
                )
                section_id = None
                focus_widget = None
                if right_tw is not None:
                    cur_item = right_tw.currentItem()
                    if cur_item is not None and cur_item.data(0, 1001) == "section":
                        section_id = cur_item.data(0, 1000)
                        focus_widget = right_tw
                if section_id is None and right_tv is not None:
                    idx = right_tv.currentIndex()
                    if idx.isValid() and idx.data(1001) == "section":
                        section_id = idx.data(1000)
                        focus_widget = right_tv
                if section_id is None:
                    return
                # Persist move
                db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                try:
                    if delta < 0:
                        db_move_section_up(int(section_id), db_path)
                    else:
                        db_move_section_down(int(section_id), db_path)
                except Exception:
                    pass
                # Ensure the right tree keeps the SECTION selected (not the first page) during refresh
                try:
                    window._keep_right_tree_section_selected_once = True
                except Exception:
                    pass
                # Refresh UI and reselect the moved section
                refresh_for_notebook(window, int(nb_id), select_section_id=int(section_id))
                # Re-assert the section selection in the right panel after the model rebuild
                try:

                    def _reselect_section():
                        try:
                            # QTreeWidget path
                            tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
                            if tw is not None:
                                for i in range(tw.topLevelItemCount()):
                                    top = tw.topLevelItem(i)
                                    try:
                                        if int(top.data(0, 1000)) == int(section_id):
                                            tw.setCurrentItem(top)
                                            tw.setFocus(Qt.OtherFocusReason)
                                            return
                                    except Exception:
                                        pass
                            # QTreeView path
                            tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
                            if tv is not None and tv.model() is not None:
                                model = tv.model()
                                for row in range(model.rowCount()):
                                    idx = model.index(row, 0)
                                    try:
                                        if idx.data(1001) == "section" and int(
                                            idx.data(1000)
                                        ) == int(section_id):
                                            tv.setCurrentIndex(idx)
                                            tv.expand(idx)
                                            tv.setFocus(Qt.OtherFocusReason)
                                            return
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                    QTimer.singleShot(0, _reselect_section)
                except Exception:
                    pass
                try:
                    ensure_left_tree_sections(window, int(nb_id), select_section_id=int(section_id))
                except Exception:
                    pass
                # Return focus to the right panel so repeated Ctrl+Up/Down works
                try:
                    if focus_widget is not None:
                        focus_widget.setFocus(Qt.OtherFocusReason)
                except Exception:
                    pass
            except Exception:
                pass

        # Bind shortcuts on the RIGHT panel (tree or view) only — unified dispatcher
        from PyQt5.QtGui import QKeySequence

        _right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
        _right_tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
        window._section_move_shortcuts = []
        if _right_tw is not None:
            sc_tw_up = QtWidgets.QShortcut(
                QKeySequence("Ctrl+Up"), _right_tw, activated=lambda: _right_panel_move(-1)
            )
            sc_tw_down = QtWidgets.QShortcut(
                QKeySequence("Ctrl+Down"), _right_tw, activated=lambda: _right_panel_move(1)
            )
            try:
                sc_tw_up.setContext(Qt.WidgetWithChildrenShortcut)
                sc_tw_down.setContext(Qt.WidgetWithChildrenShortcut)
            except Exception:
                pass
            window._section_move_shortcuts.extend([sc_tw_up, sc_tw_down])
        if _right_tv is not None:
            sc_tv_up = QtWidgets.QShortcut(
                QKeySequence("Ctrl+Up"), _right_tv, activated=lambda: _right_panel_move(-1)
            )
            sc_tv_down = QtWidgets.QShortcut(
                QKeySequence("Ctrl+Down"), _right_tv, activated=lambda: _right_panel_move(1)
            )
            try:
                sc_tv_up.setContext(Qt.WidgetWithChildrenShortcut)
                sc_tv_down.setContext(Qt.WidgetWithChildrenShortcut)
            except Exception:
                pass
            window._section_move_shortcuts.extend([sc_tv_up, sc_tv_down])
    except Exception:
        pass

    # Keyboard: Ctrl+Up / Ctrl+Down to reorder PAGES within the selected section in the right panel
    try:

        def _move_page(delta: int):
            try:
                # Determine selected page and its parent section
                right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
                right_tv = (
                    window.findChild(QtWidgets.QTreeView, "sectionPages")
                    if right_tw is None
                    else None
                )
                page_id = None
                section_id = None
                focus_widget = None
                if right_tw is not None:
                    cur = right_tw.currentItem()
                    if cur is not None and cur.data(0, 1001) == "page":
                        page_id = cur.data(0, 1000)
                        parent = cur.parent()
                        if parent is not None:
                            section_id = parent.data(0, 1000)
                        focus_widget = right_tw
                if page_id is None and right_tv is not None:
                    idx = right_tv.currentIndex()
                    if idx.isValid() and idx.data(1001) == "page":
                        page_id = idx.data(1000)
                        pidx = idx.parent()
                        if pidx.isValid() and pidx.data(1001) == "section":
                            section_id = pidx.data(1000)
                        focus_widget = right_tv
                if page_id is None or section_id is None:
                    return
                # Build ordered page id list for the section from the right panel
                ordered_ids = []
                if right_tw is not None:
                    # find the section item
                    for i in range(right_tw.topLevelItemCount()):
                        sec_item = right_tw.topLevelItem(i)
                        if int(sec_item.data(0, 1000)) == int(section_id):
                            for j in range(sec_item.childCount()):
                                ch = sec_item.child(j)
                                if ch.data(0, 1001) == "page":
                                    pid = ch.data(0, 1000)
                                    if pid is not None:
                                        ordered_ids.append(int(pid))
                            break
                elif right_tv is not None and right_tv.model() is not None:
                    model = right_tv.model()
                    # iterate top-level to find section, then its children pages
                    for row in range(model.rowCount()):
                        sec_idx = model.index(row, 0)
                        if sec_idx.data(1001) == "section" and int(sec_idx.data(1000)) == int(
                            section_id
                        ):
                            for crow in range(model.rowCount(sec_idx)):
                                child_idx = model.index(crow, 0, sec_idx)
                                if child_idx.data(1001) == "page":
                                    pid = child_idx.data(1000)
                                    if pid is not None:
                                        ordered_ids.append(int(pid))
                            break
                if not ordered_ids or page_id not in ordered_ids:
                    return
                # Compute new index for the page
                cur_idx = ordered_ids.index(int(page_id))
                new_idx = cur_idx + (1 if delta > 0 else -1)
                if new_idx < 0 or new_idx >= len(ordered_ids):
                    return
                ordered_ids[cur_idx], ordered_ids[new_idx] = (
                    ordered_ids[new_idx],
                    ordered_ids[cur_idx],
                )
                # Persist order
                db_path = getattr(window, "_db_path", None)
                if not db_path:
                    return
                try:
                    db_set_pages_order(int(section_id), ordered_ids, db_path)
                except Exception:
                    pass
                # Refresh UI: keep the same section and reselect the moved page
                nb_id = getattr(window, "_current_notebook_id", None)
                if nb_id is not None:
                    # Prevent auto-selecting the first page during refresh; we'll reassert the moved page
                    try:
                        window._keep_right_tree_section_selected_once = True
                    except Exception:
                        pass
                    refresh_for_notebook(window, int(nb_id), select_section_id=int(section_id))

                    # Defer selection + page load until after the model rebuild settles
                    def _finalize_page_selection():
                        try:
                            # Suppress sync signals while we set selection
                            try:
                                window._suppress_sync = True
                            except Exception:
                                pass
                            tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
                            done = False
                            if tw is not None:
                                for i in range(tw.topLevelItemCount()):
                                    sec_item = tw.topLevelItem(i)
                                    try:
                                        if int(sec_item.data(0, 1000)) == int(section_id):
                                            sec_item.setExpanded(True)
                                            for j in range(sec_item.childCount()):
                                                ch = sec_item.child(j)
                                                if ch.data(0, 1001) == "page" and int(
                                                    ch.data(0, 1000)
                                                ) == int(page_id):
                                                    tw.setCurrentItem(ch)
                                                    done = True
                                                    break
                                    except Exception:
                                        pass
                                    if done:
                                        break
                            if not done:
                                tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
                                if tv is not None and tv.model() is not None:
                                    model = tv.model()
                                    for row in range(model.rowCount()):
                                        sec_idx = model.index(row, 0)
                                        try:
                                            if sec_idx.data(1001) == "section" and int(
                                                sec_idx.data(1000)
                                            ) == int(section_id):
                                                tv.expand(sec_idx)
                                                for crow in range(model.rowCount(sec_idx)):
                                                    child_idx = model.index(crow, 0, sec_idx)
                                                    if child_idx.data(1001) == "page" and int(
                                                        child_idx.data(1000)
                                                    ) == int(page_id):
                                                        tv.setCurrentIndex(child_idx)
                                                        done = True
                                                        break
                                        except Exception:
                                            pass
                                        if done:
                                            break
                            try:
                                window._suppress_sync = False
                            except Exception:
                                pass
                            # Update current page mapping and load the page into the editor
                            try:
                                if not hasattr(window, "_current_page_by_section"):
                                    window._current_page_by_section = {}
                                window._current_page_by_section[int(section_id)] = int(page_id)
                            except Exception:
                                pass
                            try:
                                _load_page_two_column(window, int(page_id))
                            except Exception:
                                pass
                            try:
                                set_last_state(section_id=int(section_id), page_id=int(page_id))
                            except Exception:
                                pass
                            # Return focus to right panel for repeated moves
                            try:
                                if focus_widget is not None:
                                    focus_widget.setFocus(Qt.OtherFocusReason)
                            except Exception:
                                pass
                        except Exception:
                            pass

                    QTimer.singleShot(0, _finalize_page_selection)
                # Reselect the page after model rebuild
                try:

                    def _reselect_page():
                        try:
                            tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
                            if tw is not None:
                                # locate section item first
                                for i in range(tw.topLevelItemCount()):
                                    sec_item = tw.topLevelItem(i)
                                    if int(sec_item.data(0, 1000)) == int(section_id):
                                        for j in range(sec_item.childCount()):
                                            ch = sec_item.child(j)
                                            if ch.data(0, 1001) == "page" and int(
                                                ch.data(0, 1000)
                                            ) == int(page_id):
                                                tw.setCurrentItem(ch)
                                                tw.setFocus(Qt.OtherFocusReason)
                                                return
                            tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
                            if tv is not None and tv.model() is not None:
                                model = tv.model()
                                for row in range(model.rowCount()):
                                    sec_idx = model.index(row, 0)
                                    if sec_idx.data(1001) == "section" and int(
                                        sec_idx.data(1000)
                                    ) == int(section_id):
                                        for crow in range(model.rowCount(sec_idx)):
                                            child_idx = model.index(crow, 0, sec_idx)
                                            if child_idx.data(1001) == "page" and int(
                                                child_idx.data(1000)
                                            ) == int(page_id):
                                                tv.setCurrentIndex(child_idx)
                                                tv.expand(sec_idx)
                                                tv.setFocus(Qt.OtherFocusReason)
                                                return
                        except Exception:
                            pass

                    QTimer.singleShot(0, _reselect_page)
                except Exception:
                    pass
                # Ensure focus on right panel
                try:
                    if focus_widget is not None:
                        focus_widget.setFocus(Qt.OtherFocusReason)
                except Exception:
                    pass
            except Exception:
                pass

        # Shortcuts for right panel are handled by the unified dispatcher above; no duplicate page bindings here.
    except Exception:
        pass

    # Edit: Undo/Redo actions
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        act_undo = window.findChild(QtWidgets.QAction, "actionUndo")
        act_redo = window.findChild(QtWidgets.QAction, "actionRedo")
        
        if act_undo and te:
            from PyQt5.QtGui import QKeySequence
            act_undo.setShortcut(QKeySequence.Undo)  # Ctrl+Z
            act_undo.triggered.connect(te.undo)
            # Enable/disable based on availability
            act_undo.setEnabled(te.document().isUndoAvailable())
            te.undoAvailable.connect(act_undo.setEnabled)
        
        if act_redo and te:
            from PyQt5.QtGui import QKeySequence
            act_redo.setShortcut(QKeySequence.Redo)  # Ctrl+Y / Ctrl+Shift+Z
            act_redo.triggered.connect(te.redo)
            # Enable/disable based on availability
            act_redo.setEnabled(te.document().isRedoAvailable())
            te.redoAvailable.connect(act_redo.setEnabled)
    except Exception:
        pass

    # Edit: Paste actions
    try:
        act_paste_plain = window.findChild(QtWidgets.QAction, "actionPaste_Text_Only")
        if act_paste_plain:

            def _paste_plain():
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                    if not tab_widget:
                        return
                    page = tab_widget.currentWidget()
                    if not page:
                        return
                    te = page.findChild(QtWidgets.QTextEdit)
                    if not te:
                        return
                    from ui_richtext import paste_text_only

                    paste_text_only(te)
                    # Persist immediately so closing the app doesn't lose the paste
                    try:
                        save_current_page(window)
                    except Exception:
                        pass
                except Exception:
                    pass

            act_paste_plain.triggered.connect(_paste_plain)
        act_paste_match = window.findChild(QtWidgets.QAction, "actionPaste_and_Match_Style")
        if act_paste_match:

            def _paste_match():
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                    if not tab_widget:
                        return
                    page = tab_widget.currentWidget()
                    if not page:
                        return
                    te = page.findChild(QtWidgets.QTextEdit)
                    if not te:
                        return
                    from ui_richtext import paste_match_style

                    paste_match_style(te)
                    try:
                        save_current_page(window)
                    except Exception:
                        pass
                except Exception:
                    pass

            act_paste_match.triggered.connect(_paste_match)
        act_paste_clean = window.findChild(QtWidgets.QAction, "actionPaste_Clean_Formatting")
        if act_paste_clean:

            def _paste_clean():
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                    if not tab_widget:
                        return
                    page = tab_widget.currentWidget()
                    if not page:
                        return
                    te = page.findChild(QtWidgets.QTextEdit)
                    if not te:
                        return
                    from ui_richtext import paste_clean_formatting

                    paste_clean_formatting(te)
                    try:
                        save_current_page(window)
                    except Exception:
                        pass
                except Exception:
                    pass

            act_paste_clean.triggered.connect(_paste_clean)
    except Exception:
        pass

    # Default Paste Mode submenu wiring
    try:
        # Actions
        am_rich = window.findChild(QtWidgets.QAction, "actionPasteMode_Rich")
        am_text = window.findChild(QtWidgets.QAction, "actionPasteMode_Text_Only")
        am_match = window.findChild(QtWidgets.QAction, "actionPasteMode_Match_Style")
        am_clean = window.findChild(QtWidgets.QAction, "actionPasteMode_Clean")
        group = None
        if am_rich and am_text and am_match and am_clean:
            group = QtWidgets.QActionGroup(window)
            group.setExclusive(True)
            for a in (am_rich, am_text, am_match, am_clean):
                a.setCheckable(True)
                group.addAction(a)
            # Reflect current mode
            mode = getattr(window, "_default_paste_mode", "rich")
            if mode == "rich":
                am_rich.setChecked(True)
            elif mode == "text-only":
                am_text.setChecked(True)
            elif mode == "match-style":
                am_match.setChecked(True)
            elif mode == "clean":
                am_clean.setChecked(True)

            # Persist on change
            def _set_mode(m):
                try:
                    window._default_paste_mode = m
                    from settings_manager import set_default_paste_mode

                    set_default_paste_mode(m)
                except Exception:
                    pass

            am_rich.triggered.connect(lambda: _set_mode("rich"))
            am_text.triggered.connect(lambda: _set_mode("text-only"))
            am_match.triggered.connect(lambda: _set_mode("match-style"))
            am_clean.triggered.connect(lambda: _set_mode("clean"))
    except Exception:
        pass

    # Tools: Settings dialog
    try:
        act_settings = window.findChild(QtWidgets.QAction, "actionSettings")
        if act_settings:

            def _open_settings():
                try:
                    # Centralized loading from ui_loader
                    from ui_loader import load_settings_dialog

                    dlg = load_settings_dialog(window)
                    try:
                        dlg.setWindowModality(Qt.ApplicationModal)
                    except Exception:
                        pass

                    # Populate current settings
                    try:
                        from settings_manager import (
                            get_databases_root,
                            get_default_paste_mode,
                            get_list_schemes_settings,
                            get_plain_indent_px,
                            get_image_insert_long_side,
                            get_video_insert_long_side,
                            get_settings_file_path,
                            get_theme_name,
                            get_exit_backup_dir,
                            get_backup_on_exit_enabled,
                            get_backups_to_keep,
                        )

                        # Paste mode
                        mode = get_default_paste_mode()
                        combo = dlg.findChild(QtWidgets.QComboBox, "comboPasteMode")
                        if combo is not None:
                            mapping = {
                                "rich": "Rich",
                                "text-only": "Text Only",
                                "match-style": "Match Style",
                                "clean": "Clean Formatting",
                            }
                            text = mapping.get(mode, "Rich")
                            idx = combo.findText(text)
                            if idx >= 0:
                                combo.setCurrentIndex(idx)
                        # Plain indent px
                        sp = dlg.findChild(QtWidgets.QSpinBox, "spinIndentPx")
                        if sp is not None:
                            sp.setValue(int(get_plain_indent_px()))
                        # List schemes
                        ord_s, unord_s = get_list_schemes_settings()
                        c_ord = dlg.findChild(QtWidgets.QComboBox, "comboOrdered")
                        if c_ord is not None:
                            idx = c_ord.findText(
                                "Classic (I, A, 1, a, i)"
                                if ord_s == "classic"
                                else "Decimal (1, 2, 3)"
                            )
                            if idx >= 0:
                                c_ord.setCurrentIndex(idx)
                        c_un = dlg.findChild(QtWidgets.QComboBox, "comboUnordered")
                        if c_un is not None:
                            idx = c_un.findText(
                                "Disc → Circle → Square"
                                if unord_s == "disc-circle-square"
                                else "Disc only"
                            )
                            if idx >= 0:
                                c_un.setCurrentIndex(idx)
                        # Databases root
                        try:
                            ed = dlg.findChild(QtWidgets.QLineEdit, "editDbRoot")
                            if ed is not None:
                                ed.setText(get_databases_root())
                        except Exception:
                            pass
                        # Backups on exit settings
                        try:
                            # Field names from settings_dialog.ui
                            ed_back = (
                                dlg.findChild(QtWidgets.QLineEdit, "editDbBackup")
                                or dlg.findChild(QtWidgets.QLineEdit, "exitDbBackup")
                            )
                            chk_on_exit = dlg.findChild(QtWidgets.QCheckBox, "chkBuOnExit")
                            sp_keep = dlg.findChild(QtWidgets.QSpinBox, "spinBuToKeep")
                            if ed_back is not None:
                                ed_back.setText(get_exit_backup_dir())
                            if chk_on_exit is not None:
                                chk_on_exit.setChecked(bool(get_backup_on_exit_enabled()))
                            if sp_keep is not None:
                                sp_keep.setValue(int(get_backups_to_keep()))
                            # Browse button (support both expected names)
                            btn_browse_back = (
                                dlg.findChild(QtWidgets.QPushButton, "btnBrowseDbBackup")
                                or dlg.findChild(QtWidgets.QPushButton, "btnBrowseExitBackup")
                            )
                            if btn_browse_back is not None and ed_back is not None:
                                def _browse_backup_dir():
                                    try:
                                        import os
                                        start = ed_back.text().strip() or os.path.expanduser("~")
                                        dir_path = QtWidgets.QFileDialog.getExistingDirectory(window, "Select Backup Folder", start)
                                        if dir_path:
                                            ed_back.setText(dir_path)
                                    except Exception:
                                        pass
                                btn_browse_back.clicked.connect(_browse_backup_dir)
                        except Exception:
                            pass
                        # Default image insert size
                        try:
                            sp_img = dlg.findChild(QtWidgets.QSpinBox, "spinImageLong")
                            if sp_img is not None:
                                sp_img.setValue(int(get_image_insert_long_side()))
                        except Exception:
                            pass
                        # Default video insert size
                        try:
                            sp_vid = dlg.findChild(QtWidgets.QSpinBox, "spinVideoLong")
                            if sp_vid is not None:
                                sp_vid.setValue(int(get_video_insert_long_side()))
                        except Exception:
                            pass
                        # Theme name
                        try:
                            theme_combo = dlg.findChild(QtWidgets.QComboBox, "comboTheme")
                            if theme_combo is not None:
                                name = get_theme_name()
                                idx = theme_combo.findText(name)
                                if idx >= 0:
                                    theme_combo.setCurrentIndex(idx)
                        except Exception:
                            pass
                        # Settings file path & open folder
                        try:
                            edp = dlg.findChild(QtWidgets.QLineEdit, "editSettingsPath")
                            btn_open = dlg.findChild(QtWidgets.QPushButton, "btnOpenSettingsFolder")
                            btn_browse_settings = dlg.findChild(QtWidgets.QPushButton, "btnBrowseSettingsPath")
                            spath = get_settings_file_path()
                            if edp is not None:
                                edp.setText(spath)
                            if btn_open is not None:

                                def _open_settings_folder():
                                    try:
                                        folder = os.path.dirname(spath)
                                        from PyQt5.QtCore import QUrl
                                        from PyQt5.QtGui import QDesktopServices

                                        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
                                    except Exception:
                                        pass

                                btn_open.clicked.connect(_open_settings_folder)
                            # Browse / Change settings location
                            if btn_browse_settings is not None and edp is not None:
                                def _change_settings_location():
                                    try:
                                        from settings_manager import get_settings_dir
                                        start_dir = os.path.dirname(spath) if os.path.isdir(os.path.dirname(spath)) else get_settings_dir()
                                        new_dir = QtWidgets.QFileDialog.getExistingDirectory(window, "Choose Settings Folder", start_dir)
                                        if not new_dir:
                                            return
                                        new_dir = os.path.abspath(new_dir)
                                        # Validate write access
                                        test_ok = False
                                        try:
                                            os.makedirs(new_dir, exist_ok=True)
                                            test_file = os.path.join(new_dir, ".__nb_test_write__")
                                            with open(test_file, "w", encoding="utf-8") as tf:
                                                tf.write("ok")
                                            os.remove(test_file)
                                            test_ok = True
                                        except Exception as e:
                                            QtWidgets.QMessageBox.warning(window, "Settings", f"Selected folder is not writable:\n{e}")
                                            return
                                        if not test_ok:
                                            return
                                        # Perform migration of settings.json
                                        src = spath
                                        dst = os.path.join(new_dir, os.path.basename(spath))
                                        if os.path.abspath(src) == os.path.abspath(dst):
                                            edp.setText(dst)
                                            return
                                        # Confirm
                                        resp = QtWidgets.QMessageBox.question(
                                            window,
                                            "Move Settings",
                                            f"Move settings file to:\n{dst}?",
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                        )
                                        if resp != QtWidgets.QMessageBox.Yes:
                                            return
                                        # Copy (not move) first, ensure integrity
                                        import shutil
                                        try:
                                            shutil.copy2(src, dst)
                                        except FileNotFoundError:
                                            # No existing file, create empty settings.json
                                            with open(dst, "w", encoding="utf-8") as nf:
                                                nf.write("{}\n")
                                        except Exception as e:
                                            QtWidgets.QMessageBox.warning(window, "Settings", f"Failed to migrate settings:\n{e}")
                                            return
                                        # Tell settings_manager to start using the new location immediately and persist across restarts
                                        try:
                                            import settings_manager as sm
                                            sm.set_settings_file_path(dst)
                                        except Exception:
                                            pass
                                        # Replace source file with new one (optional move)
                                        try:
                                            # Keep original as backup; do not delete automatically
                                            pass
                                        except Exception:
                                            pass
                                        edp.setText(dst)
                                        QtWidgets.QMessageBox.information(window, "Settings", "Settings location updated. It will be used immediately and on next launch.")
                                    except Exception as e:
                                        QtWidgets.QMessageBox.warning(window, "Settings", f"Failed to change settings location:\n{e}")

                                btn_browse_settings.clicked.connect(_change_settings_location)
                        except Exception:
                            pass
                        # Tables tab: load current table theme
                        try:
                            from settings_manager import get_table_theme
                            theme = get_table_theme()
                            ed_gc = dlg.findChild(QtWidgets.QLineEdit, "editGridColor")
                            sp_gw = dlg.findChild(QtWidgets.QDoubleSpinBox, "spinGridWidth")
                            ed_hb = dlg.findChild(QtWidgets.QLineEdit, "editHeaderBg")
                            ed_tb = dlg.findChild(QtWidgets.QLineEdit, "editTotalsBg")
                            ed_cb = dlg.findChild(QtWidgets.QLineEdit, "editCostHeaderBg")
                            if ed_gc is not None:
                                ed_gc.setText(theme.get("grid_color", "#000000"))
                            if sp_gw is not None:
                                sp_gw.setValue(float(theme.get("grid_width", 1.0)))
                            if ed_hb is not None:
                                ed_hb.setText(theme.get("header_bg", "#F5F5F5"))
                            if ed_tb is not None:
                                ed_tb.setText(theme.get("totals_bg", "#F5F5F5"))
                            if ed_cb is not None:
                                ed_cb.setText(theme.get("cost_header_bg", "#F5F5F5"))
                            # Wire simple color pickers
                            def _pick_into(line_edit):
                                col = QtWidgets.QColorDialog.getColor(parent=dlg)
                                if col.isValid() and line_edit is not None:
                                    line_edit.setText(col.name())
                            btn_gc = dlg.findChild(QtWidgets.QPushButton, "btnPickGridColor")
                            if btn_gc is not None and ed_gc is not None:
                                btn_gc.clicked.connect(lambda: _pick_into(ed_gc))
                            btn_hb = dlg.findChild(QtWidgets.QPushButton, "btnPickHeaderBg")
                            if btn_hb is not None and ed_hb is not None:
                                btn_hb.clicked.connect(lambda: _pick_into(ed_hb))
                            btn_tb = dlg.findChild(QtWidgets.QPushButton, "btnPickTotalsBg")
                            if btn_tb is not None and ed_tb is not None:
                                btn_tb.clicked.connect(lambda: _pick_into(ed_tb))
                            btn_cb = dlg.findChild(QtWidgets.QPushButton, "btnPickCostHeaderBg")
                            if btn_cb is not None and ed_cb is not None:
                                btn_cb.clicked.connect(lambda: _pick_into(ed_cb))
                        except Exception:
                            pass
                        # Browse for databases root
                        try:
                            btn_browse = dlg.findChild(QtWidgets.QPushButton, "btnBrowseDbRoot")
                            ed = dlg.findChild(QtWidgets.QLineEdit, "editDbRoot")
                            if btn_browse is not None and ed is not None:

                                def _browse_db_root():
                                    try:
                                        start = ed.text().strip() or get_databases_root()
                                        dir_path = QtWidgets.QFileDialog.getExistingDirectory(
                                            window, "Select Databases Root", start
                                        )
                                        if dir_path:
                                            ed.setText(dir_path)
                                    except Exception:
                                        pass

                                btn_browse.clicked.connect(_browse_db_root)
                        except Exception:
                            pass
                    except Exception:
                        pass

                    if dlg.exec_() != QtWidgets.QDialog.Accepted:
                        return
                    # Persist settings
                    try:
                        from settings_manager import (
                            set_databases_root,
                            set_default_paste_mode,
                            set_list_schemes_settings,
                            set_plain_indent_px,
                            set_image_insert_long_side,
                            set_video_insert_long_side,
                            set_theme_name,
                            set_exit_backup_dir,
                            set_backup_on_exit_enabled,
                            set_backups_to_keep,
                        )

                        # Paste mode
                        combo = dlg.findChild(QtWidgets.QComboBox, "comboPasteMode")
                        if combo is not None:
                            txt = combo.currentText()
                            inv = {
                                "Rich": "rich",
                                "Text Only": "text-only",
                                "Match Style": "match-style",
                                "Clean Formatting": "clean",
                            }
                            m = inv.get(txt, "rich")
                            set_default_paste_mode(m)
                            window._default_paste_mode = m
                        # Indent step
                        sp = dlg.findChild(QtWidgets.QSpinBox, "spinIndentPx")
                        if sp is not None:
                            set_plain_indent_px(int(sp.value()))
                            # Update active editors' indent step immediately
                            try:
                                import ui_richtext as rt

                                rt.INDENT_STEP_PX = float(sp.value())
                            except Exception:
                                pass
                        # Default image insert long side
                        try:
                            sp_img = dlg.findChild(QtWidgets.QSpinBox, "spinImageLong")
                            if sp_img is not None:
                                val = int(sp_img.value())
                                set_image_insert_long_side(val)
                                # Apply immediately to runtime constant
                                try:
                                    import ui_richtext as rt
                                    rt.DEFAULT_IMAGE_LONG_SIDE = int(val)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # Backups on exit: persist
                        try:
                            ed_back = (
                                dlg.findChild(QtWidgets.QLineEdit, "editDbBackup")
                                or dlg.findChild(QtWidgets.QLineEdit, "exitDbBackup")
                            )
                            chk_on_exit = dlg.findChild(QtWidgets.QCheckBox, "chkBuOnExit")
                            sp_keep = dlg.findChild(QtWidgets.QSpinBox, "spinBuToKeep")
                            if ed_back is not None:
                                set_exit_backup_dir((ed_back.text() or "").strip())
                            if chk_on_exit is not None:
                                set_backup_on_exit_enabled(bool(chk_on_exit.isChecked()))
                            if sp_keep is not None:
                                set_backups_to_keep(int(sp_keep.value()))
                        except Exception:
                            pass
                        # Default video insert long side
                        try:
                            sp_vid = dlg.findChild(QtWidgets.QSpinBox, "spinVideoLong")
                            if sp_vid is not None:
                                vval = int(sp_vid.value())
                                set_video_insert_long_side(vval)
                                try:
                                    import ui_richtext as rt
                                    rt.DEFAULT_VIDEO_LONG_SIDE = int(vval)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # List schemes
                        c_ord = dlg.findChild(QtWidgets.QComboBox, "comboOrdered")
                        c_un = dlg.findChild(QtWidgets.QComboBox, "comboUnordered")
                        ordered = "classic"
                        unordered = "disc-circle-square"
                        if c_ord is not None and c_ord.currentText().startswith("Decimal"):
                            ordered = "decimal"
                        if c_un is not None and c_un.currentText().startswith("Disc only"):
                            unordered = "disc-only"
                        set_list_schemes_settings(ordered=ordered, unordered=unordered)
                        try:
                            from ui_richtext import set_list_schemes

                            set_list_schemes(ordered=ordered, unordered=unordered)
                        except Exception:
                            pass
                        # Databases root
                        ed = dlg.findChild(QtWidgets.QLineEdit, "editDbRoot")
                        if ed is not None:
                            path = (ed.text() or "").strip()
                            if path:
                                set_databases_root(path)
                        # Theme name
                        theme_combo = dlg.findChild(QtWidgets.QComboBox, "comboTheme")
                        if theme_combo is not None:
                            name = theme_combo.currentText()
                            set_theme_name(name)
                            # Apply selected theme immediately
                            try:
                                import os

                                themes_dir = os.path.join(os.path.dirname(__file__), "themes")
                                name_to_file = {
                                    "Default": "default.qss",
                                    "High Contrast": "high-contrast.qss",
                                }
                                qss_file = name_to_file.get(name)
                                if qss_file:
                                    path = os.path.join(themes_dir, qss_file)
                                    if os.path.isfile(path):
                                        with open(path, "r", encoding="utf-8") as f:
                                            QtWidgets.QApplication.instance().setStyleSheet(
                                                f.read()
                                            )
                            except Exception:
                                pass
                        # Tables tab: persist and apply immediately
                        try:
                            from settings_manager import set_table_theme, get_table_theme
                            ed_gc = dlg.findChild(QtWidgets.QLineEdit, "editGridColor")
                            sp_gw = dlg.findChild(QtWidgets.QDoubleSpinBox, "spinGridWidth")
                            ed_hb = dlg.findChild(QtWidgets.QLineEdit, "editHeaderBg")
                            ed_tb = dlg.findChild(QtWidgets.QLineEdit, "editTotalsBg")
                            ed_cb = dlg.findChild(QtWidgets.QLineEdit, "editCostHeaderBg")
                            kwargs = {}
                            if ed_gc is not None:
                                kwargs["grid_color"] = ed_gc.text().strip() or "#000000"
                            if sp_gw is not None:
                                kwargs["grid_width"] = float(sp_gw.value())
                            if ed_hb is not None:
                                kwargs["header_bg"] = ed_hb.text().strip() or "#F5F5F5"
                            if ed_tb is not None:
                                kwargs["totals_bg"] = ed_tb.text().strip() or "#F5F5F5"
                            if ed_cb is not None:
                                kwargs["cost_header_bg"] = ed_cb.text().strip() or "#F5F5F5"
                            set_table_theme(**kwargs)
                            # Apply immediately to current editor content
                            try:
                                te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                                if te is not None:
                                    # Re-run refresh and border enforcement with new colors/widths
                                    from ui_planning_register import refresh_planning_register_styles
                                    refresh_planning_register_styles(te)
                                    from ui_richtext import _enforce_uniform_table_borders
                                    _enforce_uniform_table_borders(te)
                            except Exception:
                                pass
                        except Exception:
                            pass
                    except Exception as e:
                        QtWidgets.QMessageBox.warning(
                            window, "Settings", f"Failed to save settings: {e}"
                        )
                except Exception as e:
                    QtWidgets.QMessageBox.warning(
                        window, "Settings", f"Failed to open settings: {e}"
                    )

            act_settings.triggered.connect(_open_settings)
    except Exception:
        pass

    # Tools: Clean Unused Media
    try:
        act_clean_media = window.findChild(QtWidgets.QAction, "actionClean_Unused_Media")
        if act_clean_media:

            def _do_clean_media():
                try:
                    from media_store import garbage_collect_unused_media

                    dbp = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                    removed = garbage_collect_unused_media(dbp)
                    QtWidgets.QMessageBox.information(
                        window,
                        "Clean Unused Media",
                        f"Removed {removed} unreferenced media file(s).",
                    )
                except Exception as e:
                    QtWidgets.QMessageBox.warning(
                        window, "Clean Unused Media", f"Failed to clean media: {e}"
                    )

            act_clean_media.triggered.connect(_do_clean_media)
    except Exception:
        pass

    # Tools: Open Databases Folder
    try:
        act_open_root = window.findChild(QtWidgets.QAction, "actionOpen_Databases_Folder")
        if act_open_root:

            def _open_root():
                try:
                    from PyQt5.QtCore import QUrl
                    from PyQt5.QtGui import QDesktopServices

                    from settings_manager import get_databases_root

                    QDesktopServices.openUrl(QUrl.fromLocalFile(get_databases_root()))
                except Exception:
                    pass

            act_open_root.triggered.connect(_open_root)
    except Exception:
        pass

    # Tools: Migrate current DB into Databases Root
    try:
        act_migrate = window.findChild(QtWidgets.QAction, "actionMigrate_Current_DB_to_Root")
        if act_migrate:

            def _migrate_into_root():
                try:
                    import os
                    import shutil

                    from media_store import media_root_for_db
                    from settings_manager import get_databases_root

                    src_db = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                    root = get_databases_root()
                    base = os.path.basename(src_db)
                    dst_db = os.path.join(root, base)
                    if os.path.abspath(src_db) == os.path.abspath(dst_db):
                        QtWidgets.QMessageBox.information(
                            window, "Migrate", "Current database is already in the Databases root."
                        )
                        return
                    # Confirm
                    resp = QtWidgets.QMessageBox.question(
                        window,
                        "Migrate Database",
                        f"Copy current database to:\n{dst_db}\n\nAnd copy its media folder next to it. Proceed?",
                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    )
                    if resp != QtWidgets.QMessageBox.Yes:
                        return
                    # Ensure root exists
                    os.makedirs(root, exist_ok=True)
                    # Copy DB
                    if os.path.exists(dst_db):
                        os.remove(dst_db)
                    shutil.copy2(src_db, dst_db)
                    # Copy media
                    src_media = media_root_for_db(src_db)
                    dst_media = media_root_for_db(dst_db)
                    if os.path.isdir(src_media):
                        if os.path.exists(dst_media):
                            shutil.rmtree(dst_media, ignore_errors=True)
                        shutil.copytree(src_media, dst_media)
                    # Switch to migrated copy
                    set_last_db(dst_db)
                    clear_last_state()
                    restart_application()
                except Exception as e:
                    QtWidgets.QMessageBox.warning(
                        window, "Migrate", f"Failed to migrate database: {e}"
                    )

            act_migrate.triggered.connect(_migrate_into_root)
    except Exception:
        pass

    # Format > List Scheme (wired to actions defined in main_window_5.ui)
    try:

        def _apply_list_schemes(ordered=None, unordered=None):
            try:
                from settings_manager import set_list_schemes_settings
                from ui_richtext import set_list_schemes

                set_list_schemes(ordered=ordered, unordered=unordered)
                set_list_schemes_settings(ordered=ordered, unordered=unordered)
            except Exception:
                return

        act_ord_classic = window.findChild(QtWidgets.QAction, "actionOrdered_Classic")
        if act_ord_classic:
            act_ord_classic.triggered.connect(lambda: _apply_list_schemes(ordered="classic"))
        act_ord_decimal = window.findChild(QtWidgets.QAction, "actionOrdered_Decimal")
        if act_ord_decimal:
            act_ord_decimal.triggered.connect(lambda: _apply_list_schemes(ordered="decimal"))
        act_un_disc_cs = window.findChild(QtWidgets.QAction, "actionUnordered_Disc_Circle_Square")
        if act_un_disc_cs:
            act_un_disc_cs.triggered.connect(
                lambda: _apply_list_schemes(unordered="disc-circle-square")
            )
        act_un_disc_only = window.findChild(QtWidgets.QAction, "actionUnordered_Disc_Only")
        if act_un_disc_only:
            act_un_disc_only.triggered.connect(lambda: _apply_list_schemes(unordered="disc-only"))
    except Exception:
        pass

    # Help menu
    try:
        act_documentation = window.findChild(QtWidgets.QAction, "actionDocumentation")
        if act_documentation:
            def _open_documentation():
                """Open the README.md file in the default browser."""
                try:
                    from PyQt5.QtCore import QUrl
                    from PyQt5.QtGui import QDesktopServices
                    import os
                    
                    readme_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "README.md")
                    if os.path.exists(readme_path):
                        QDesktopServices.openUrl(QUrl.fromLocalFile(readme_path))
                    else:
                        QtWidgets.QMessageBox.warning(
                            window, "Documentation", "README.md not found."
                        )
                except Exception as e:
                    QtWidgets.QMessageBox.warning(
                        window, "Documentation", f"Failed to open documentation: {e}"
                    )
            act_documentation.triggered.connect(_open_documentation)

        act_shortcuts = window.findChild(QtWidgets.QAction, "actionKeyboard_Shortcuts")
        if act_shortcuts:
            def _show_shortcuts():
                """Show a dialog with keyboard shortcuts."""
                try:
                    msg = """<h3>Keyboard Shortcuts</h3>
<table border="0" cellpadding="5">
<tr><td><b>General</b></td><td></td></tr>
<tr><td>Ctrl+N</td><td>New Database</td></tr>
<tr><td>Ctrl+O</td><td>Open Database</td></tr>
<tr><td>Ctrl+S</td><td>Save (auto-saves on edit)</td></tr>
<tr><td>Ctrl+Shift+S</td><td>Save Database As...</td></tr>
<tr><td></td><td></td></tr>
<tr><td><b>Editing</b></td><td></td></tr>
<tr><td>Ctrl+B</td><td>Bold</td></tr>
<tr><td>Ctrl+I</td><td>Italic</td></tr>
<tr><td>Ctrl+U</td><td>Underline</td></tr>
<tr><td>Ctrl+V</td><td>Paste (mode set in Edit menu)</td></tr>
<tr><td>Ctrl+Shift+V</td><td>Paste as Plain Text</td></tr>
<tr><td></td><td></td></tr>
<tr><td><b>Tables</b></td><td></td></tr>
<tr><td>Tab</td><td>Next cell (or insert row at end)</td></tr>
<tr><td>Shift+Tab</td><td>Previous cell</td></tr>
<tr><td>Right-click</td><td>Table context menu</td></tr>
<tr><td></td><td></td></tr>
<tr><td><b>Currency Columns</b></td><td></td></tr>
<tr><td>Click header</td><td>Mark/unmark column as currency</td></tr>
<tr><td>Auto-format</td><td>Numbers formatted as $#,##0.00</td></tr>
<tr><td>Auto-total</td><td>Sum appears in bottom Total row</td></tr>
</table>
"""
                    dlg = QtWidgets.QMessageBox(window)
                    dlg.setWindowTitle("Keyboard Shortcuts")
                    dlg.setTextFormat(Qt.RichText)
                    dlg.setText(msg)
                    dlg.setIcon(QtWidgets.QMessageBox.Information)
                    dlg.exec_()
                except Exception as e:
                    QtWidgets.QMessageBox.warning(
                        window, "Shortcuts", f"Failed to display shortcuts: {e}"
                    )
            act_shortcuts.triggered.connect(_show_shortcuts)

        act_about = window.findChild(QtWidgets.QAction, "actionAbout")
        if act_about:
            def _show_about():
                """Show About dialog with version and credits."""
                try:
                    msg = """<h2>NoteBook</h2>
<p><b>Version:</b> 1.0.0</p>
<p>A rich-text note-taking application with binders, sections, and pages.</p>
<p><b>Features:</b></p>
<ul>
<li>Rich text editing with tables, images, and attachments</li>
<li>Currency columns with automatic formatting and totals</li>
<li>Planning registers for structured data entry</li>
<li>SQLite-based storage with media management</li>
<li>Customizable themes and settings</li>
</ul>
<p>Built with PyQt5 and Python.</p>
"""
                    dlg = QtWidgets.QMessageBox(window)
                    dlg.setWindowTitle("About NoteBook")
                    dlg.setTextFormat(Qt.RichText)
                    dlg.setText(msg)
                    dlg.setIcon(QtWidgets.QMessageBox.Information)
                    dlg.exec_()
                except Exception as e:
                    QtWidgets.QMessageBox.warning(
                        window, "About", f"Failed to display about dialog: {e}"
                    )
            act_about.triggered.connect(_show_about)
    except Exception:
        pass

    window.show()

    # Ensure the window is actually visible on current monitors (handles monitor changes)
    def _ensure_window_visible():
        try:
            # Collect available geometries for all connected screens
            screens = QtWidgets.QApplication.screens() or []
            if not screens:
                return
            rects = []
            try:
                rects = [s.availableGeometry() for s in screens]
            except Exception:
                # Fallback to full geometries if available ones are not accessible
                rects = [s.geometry() for s in screens]

            # Determine current window frame geometry (includes window frame)
            try:
                g = window.frameGeometry()
            except Exception:
                g = window.geometry()

            def _intersects_any(target):
                try:
                    for r in rects:
                        try:
                            if target.intersects(r):
                                return True
                        except Exception:
                            pass
                except Exception:
                    pass
                return False

            too_small = False
            try:
                too_small = (g.width() < 100) or (g.height() < 100)
            except Exception:
                pass

            # If the window is off-screen or tiny, move it to the primary screen and size reasonably
            if (not _intersects_any(g)) or too_small:
                try:
                    primary = QtWidgets.QApplication.primaryScreen()
                    pr = primary.availableGeometry() if primary is not None else rects[0]
                except Exception:
                    pr = rects[0]
                try:
                    w = min(int(pr.width() * 0.8), 1200)
                    h = min(int(pr.height() * 0.8), 800)
                    w = max(w, 800)
                    h = max(h, 600)
                except Exception:
                    w, h = 1000, 700
                try:
                    x = pr.x() + (pr.width() - w) // 2
                    y = pr.y() + (pr.height() - h) // 2
                except Exception:
                    x, y = 100, 100

                # Ensure window is in a normal state before moving/resizing
                try:
                    if window.isMaximized() or window.isMinimized():
                        window.setWindowState(Qt.WindowNoState)
                except Exception:
                    pass
                try:
                    window.setGeometry(x, y, w, h)
                except Exception:
                    pass

                # Re-apply maximized state if that's the user's preference
                try:
                    if get_window_maximized():
                        window.showMaximized()
                except Exception:
                    pass
        except Exception:
            pass

    # Run visibility correction after the window is initially shown so frameGeometry is valid
    QTimer.singleShot(0, _ensure_window_visible)

    # Restore splitter sizes after the window is shown to ensure geometry exists
    def _apply_saved_splitter_sizes():
        try:
            splitter = window.findChild(QtWidgets.QSplitter, "mainSplitter")
            if splitter is None:
                return
            from settings_manager import get_splitter_sizes, set_splitter_sizes

            sizes = get_splitter_sizes()
            if sizes:
                # Fit the sizes list to current pane count
                count = splitter.count()
                if len(sizes) > count:
                    sizes = sizes[:count]
                elif len(sizes) < count:
                    sizes = sizes + [max(120, 300)] * (count - len(sizes))
                safe = [max(80, int(x)) for x in sizes]
                splitter.setSizes(safe)
            # Save on every move (lightweight) so crashes don’t lose user resize
            try:
                splitter.splitterMoved.connect(
                    lambda pos, index: set_splitter_sizes(splitter.sizes())
                )
            except Exception:
                pass
        except Exception:
            pass

    QTimer.singleShot(0, _apply_saved_splitter_sizes)

    # Save geometry on close
    def save_geometry():
        # Save current page content first to avoid losing last-minute edits
        try:
            save_current_page(window)
        except Exception:
            pass
        g = window.geometry()
        set_window_geometry(g.x(), g.y(), g.width(), g.height())
        set_window_maximized(window.isMaximized())
        # Persist splitter sizes
        try:
            splitter = window.findChild(QtWidgets.QSplitter, "mainSplitter")
            if splitter is not None:
                from settings_manager import set_splitter_sizes

                set_splitter_sizes(splitter.sizes())
        except Exception:
            pass
        # Backup on exit (best-effort, after content and geometry saves)
        try:
            from settings_manager import (
                get_backup_on_exit_enabled,
                get_exit_backup_dir,
                get_backups_to_keep,
            )
            # Allow a runtime override to disable exit backup entirely
            disable_env = os.environ.get("NOTEBOOK_DISABLE_EXIT_BACKUP", "").strip().lower() in {"1", "true", "yes"}
            if (not disable_env) and get_backup_on_exit_enabled():
                dest = (get_exit_backup_dir() or "").strip()
                if dest:
                    dbp = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                    try:
                        from backup import make_exit_backup

                        make_exit_backup(dbp, dest, keep=get_backups_to_keep(), include_media=True)
                    except KeyboardInterrupt:
                        # Ignore Ctrl+C or interrupt during compression on app exit
                        pass
                    except BaseException:
                        # Swallow any other fatal errors to avoid blocking shutdown
                        pass
        except Exception:
            pass

    app.aboutToQuit.connect(save_geometry)

    # Override Ctrl+V to use the selected default paste mode in the current editor
    try:
        paste_shortcut = QtWidgets.QShortcut(QtWidgets.QKeySequence.Paste, window)

        def _on_default_paste():
            try:
                tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
                if not tab_widget:
                    return
                page = tab_widget.currentWidget()
                if not page:
                    return
                te = page.findChild(QtWidgets.QTextEdit)
                if not te:
                    return
                mode = getattr(window, "_default_paste_mode", "rich")
                if mode == "text-only":
                    from ui_richtext import paste_text_only

                    paste_text_only(te)
                elif mode == "match-style":
                    from ui_richtext import paste_match_style

                    paste_match_style(te)
                elif mode == "clean":
                    from ui_richtext import paste_clean_formatting

                    paste_clean_formatting(te)
                else:
                    # default rich paste: let QTextEdit handle as usual
                    te.paste()
                # Save after paste
                try:
                    save_current_page(window)
                except Exception:
                    pass
            except Exception:
                pass

        paste_shortcut.activated.connect(_on_default_paste)
    except Exception:
        pass
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
