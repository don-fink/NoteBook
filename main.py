"""
main.py
Entry point for the NoteBook application. Handles main window setup, menu actions, database creation/opening, and application startup.
"""
import os

import sys
import warnings

from PyQt5 import QtWidgets, uic
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
from ui_tabs import restore_last_position, setup_tab_sync
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


def _enable_faulthandler(log_path: str):
    """Enable Python faulthandler to dump tracebacks on fatal errors (e.g., segfaults).

    Writes native crash backtraces for all threads to the given log file.
    """
    try:
        import faulthandler as _faulthandler
        # Keep a global reference so the file handle stays open for the lifetime of the app
        globals().setdefault("_native_crash_log_file", None)
        try:
            f = open(log_path, "a", encoding="utf-8")
            globals()["_native_crash_log_file"] = f
        except Exception:
            f = None
        if f is not None:
            try:
                _faulthandler.enable(all_threads=True, file=f)
            except Exception:
                try:
                    f.close()
                except Exception:
                    pass
                globals()["_native_crash_log_file"] = None
    except Exception:
        pass


def _install_qt_message_handler(log_path: str):
    """Capture Qt warnings/errors into a log to aid diagnosing native crashes."""
    try:
        from PyQt5.QtCore import qInstallMessageHandler, QtMsgType
        import datetime as _dt

        level_map = {
            QtMsgType.QtDebugMsg: "DEBUG",
            QtMsgType.QtInfoMsg: "INFO",
            QtMsgType.QtWarningMsg: "WARNING",
            QtMsgType.QtCriticalMsg: "CRITICAL",
            QtMsgType.QtFatalMsg: "FATAL",
        }

        def _qt_handler(msgType, context, message):
            try:
                ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
                level = level_map.get(msgType, str(msgType))
                with open(log_path, "a", encoding="utf-8") as _f:
                    _f.write(f"[{ts}] [Qt {level}] {message}\n")
                    try:
                        file = getattr(context, "file", None)
                        line = getattr(context, "line", None)
                        func = getattr(context, "function", None)
                        if file or func:
                            _f.write(f"    at {file or '?'}:{line or '?'} ({func or '?'})\n")
                    except Exception:
                        pass
            except Exception:
                pass

        qInstallMessageHandler(_qt_handler)
    except Exception:
        pass

def create_new_database(window):
    options = QtWidgets.QFileDialog.Options()
    # Default to the configured Databases root
    try:
        import os

        from settings_manager import get_databases_root

        initial = os.path.join(get_databases_root(), "NewNotebook.db")
    except Exception:
        initial = "NewNotebook.db"
    file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
        window,
        "Create New Database",
        initial,
        "SQLite DB Files (*.db);;All Files (*)",
        options=options,
    )
    if not file_name:
        return
    # Ensure .db extension
    if not str(file_name).lower().endswith(".db"):
        file_name = file_name + ".db"
    import sqlite3

    conn = sqlite3.connect(file_name)
    cursor = conn.cursor()
    cursor.executescript(
        """
        PRAGMA foreign_keys = ON;
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
            color_hex TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            modified_at TEXT NOT NULL DEFAULT (datetime('now')),
            order_index INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (notebook_id) REFERENCES notebooks(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS pages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            section_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content_html TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            modified_at TEXT NOT NULL DEFAULT (datetime('now')),
            order_index INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (section_id) REFERENCES sections(id) ON DELETE CASCADE
        );
    """
    )
    conn.commit()
    # Set version to 2 (includes sections.color_hex)
    cursor.execute("PRAGMA user_version = 2")
    conn.commit()
    conn.close()
    # Prepare media root directory for this DB
    try:
        from media_store import ensure_dir, media_root_for_db

        media_base = media_root_for_db(file_name)
        ensure_dir(media_base)
    except Exception:
        pass
    set_last_db(file_name)
    clear_last_state()
    # Force a clean restart so UI initializes with the new database
    restart_application()


def _select_left_tree_notebook(window, notebook_id: int):
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if not tree_widget:
        return
    for i in range(tree_widget.topLevelItemCount()):
        top = tree_widget.topLevelItem(i)
        if top.data(0, 1000) == notebook_id:
            tree_widget.setCurrentItem(top)
            break


def add_binder(window):
    title, ok = QtWidgets.QInputDialog.getText(
        window, "Add Binder", "Binder title:", text="Untitled Binder"
    )
    if not ok:
        return
    title = (title or "").strip() or "Untitled Binder"
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    # Capture current expanded state of top-level binders and persist before refresh
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        expanded_ids = set()
        if tree_widget is not None:
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

    # (context menu wiring moved to main())

    # Create notebook and refresh UI
    nid = db_create_notebook(title, db_path)
    set_last_state(notebook_id=nid, section_id=None, page_id=None)
    populate_notebook_names(window, db_path)
    # Restore previously expanded binders (do not auto-expand the new one)
    try:
        from settings_manager import get_expanded_notebooks
        from left_tree import ensure_left_tree_sections

        persisted_ids = get_expanded_notebooks()
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if tree_widget is not None and persisted_ids:
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
    # Select the new binder but keep it collapsed to preserve current tree state
    _select_left_tree_notebook(window, nid)
    refresh_for_notebook(window, nid)


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
    """Clear and rebuild the entire UI from current db_path and last state."""
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    # Clear widgets
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if tree_widget:
        tree_widget.clear()
    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
    if tab_widget:
        tab_widget.clear()
    right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
    if right_tw:
        right_tw.clear()
    right_tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
    if right_tv and right_tv.model() is not None:
        right_tv.setModel(None)
    populate_notebook_names(window, db_path)
    setup_tab_sync(window)
    # In both legacy (tabs) and 2-column mode, restore last viewed position
    restore_last_position(window)
    # Prepare splitter stretch factors (favor center panel); apply sizes after show
    try:
        splitter = window.findChild(QtWidgets.QSplitter, "mainSplitter")
        if splitter is not None:
            try:
                splitter.setStretchFactor(0, 0)  # left
                splitter.setStretchFactor(1, 2)  # center
                splitter.setStretchFactor(2, 0)  # right
            except Exception:
                pass
    except Exception:
        pass


def add_page(window):
    # Determine active section from: tabs (legacy), right pane, left pane (2-col), or current context
    tab_widget = window.findChild(QtWidgets.QTabWidget, "tabPages")
    tab_bar = tab_widget.tabBar() if tab_widget else None
    section_id = None
    if tab_widget and tab_widget.count() > 0 and tab_bar is not None:
        idx = tab_widget.currentIndex()
        section_id = tab_bar.tabData(idx)
    if section_id is None:
        # Right pane selection (legacy)
        right_tw = window.findChild(QtWidgets.QTreeWidget, "sectionPages")
        if right_tw and right_tw.currentItem() is not None:
            cur = right_tw.currentItem()
            kind = cur.data(0, 1001)
            if kind == "section":
                section_id = cur.data(0, 1000)
            elif kind == "page":
                section_id = cur.data(0, 1002)
        # Model view
        if section_id is None:
            right_tv = window.findChild(QtWidgets.QTreeView, "sectionPages")
            try:
                idx_obj = right_tv.currentIndex() if right_tv is not None else None
                if idx_obj is not None and idx_obj.isValid():
                    kind = idx_obj.data(1001)
                    if kind == "section":
                        section_id = idx_obj.data(1000)
                    elif kind == "page":
                        section_id = idx_obj.data(1002)
            except Exception:
                pass
    # Two-column: try left tree (notebookName)
    if section_id is None:
        try:
            tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
            cur = tree.currentItem() if tree is not None else None
            if cur is not None and cur.parent() is not None:
                kind = cur.data(0, 1001)
                if kind == "section":
                    section_id = cur.data(0, 1000)
                elif kind == "page":
                    # Prefer explicit parent section id role, else derive from parent
                    section_id = cur.data(0, 1002) or (
                        cur.parent().data(0, 1000) if cur.parent() is not None else None
                    )
        except Exception:
            pass
    # Fallback to current section context in two-column mode
    if section_id is None:
        try:
            if _is_two_column_ui(window):
                section_id = getattr(window, "_current_section_id", None)
        except Exception:
            pass
    if section_id is None:
        QtWidgets.QMessageBox.information(
            window, "Add Page", "Please select or create a section first."
        )
        return
    db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    pid = db_create_page(int(section_id), "Untitled Page", db_path)
    # Update UI depending on mode
    try:
        if _is_two_column_ui(window):
            # Refresh left tree under the owning binder so the new page appears
            try:
                # Find owning notebook for this section
                import sqlite3

                con = sqlite3.connect(db_path)
                cur = con.cursor()
                cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
                row = cur.fetchone()
                con.close()
                nb_id = int(row[0]) if row else None
            except Exception:
                nb_id = None
            try:
                if nb_id is not None:
                    ensure_left_tree_sections(
                        window, nb_id, select_section_id=int(section_id)
                    )
            except Exception:
                pass
            # Set current context and load the new page into the editor
            try:
                window._current_section_id = int(section_id)
                if not hasattr(window, "_current_page_by_section"):
                    window._current_page_by_section = {}
                window._current_page_by_section[int(section_id)] = int(pid)
            except Exception:
                pass
            try:
                _load_page_two_column(window, int(pid))
                # Ensure left tree selects the new page and section remains expanded
                select_left_tree_page(window, int(section_id), int(pid))
            except Exception:
                pass
            # Persist last state
            try:
                set_last_state(section_id=int(section_id), page_id=pid)
            except Exception:
                pass
            return
    except Exception:
        pass
    # Legacy tabs: persist and restore selection via tab logic
    set_last_state(section_id=int(section_id), page_id=pid)
    restore_last_position(window)


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
    """Prompt for a file and attach it to the current page via media store; no inline HTML yet."""
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
        from media_store import add_media_ref, save_file_into_store

        media_id, rel_path = save_file_into_store(db_path, file_path)
        add_media_ref(db_path, media_id, page_id=page_id, role="attachment")
        QtWidgets.QMessageBox.information(
            window, "Insert Attachment", f"Attached file saved to media store.\n{rel_path}"
        )
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, "Insert Attachment", f"Failed to attach file: {e}")


def migrate_database_if_needed(db_path):
    from db_version import get_db_version, set_db_version

    current_version = get_db_version(db_path)
    target_version = 3  # Update as needed for future migrations
    import sqlite3

    if current_version < 1:
        # baseline
        set_db_version(1, db_path)
        current_version = 1
    if current_version < 2:
        # Add color_hex to sections
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute("ALTER TABLE sections ADD COLUMN color_hex TEXT")
            conn.commit()
        except sqlite3.OperationalError:
            # Column may already exist; ignore
            pass
        finally:
            try:
                conn.close()
            except Exception:
                pass
        set_db_version(2, db_path)
    if current_version < 3:
        # Add media storage tables and a database metadata table (uuid)
        import uuid

        conn = sqlite3.connect(db_path)
        try:
            cur = conn.cursor()
            cur.executescript(
                """
                CREATE TABLE IF NOT EXISTS db_metadata (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    uuid TEXT NOT NULL
                );
                
                CREATE TABLE IF NOT EXISTS media (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sha256 TEXT NOT NULL UNIQUE,
                    mime_type TEXT NOT NULL,
                    ext TEXT NOT NULL,
                    original_filename TEXT,
                    size_bytes INTEGER,
                    created_at TEXT NOT NULL DEFAULT (datetime('now'))
                );
                
                CREATE TABLE IF NOT EXISTS media_refs (
                    media_id INTEGER NOT NULL REFERENCES media(id) ON DELETE CASCADE,
                    page_id INTEGER REFERENCES pages(id) ON DELETE CASCADE,
                    section_id INTEGER REFERENCES sections(id) ON DELETE CASCADE,
                    notebook_id INTEGER REFERENCES notebooks(id) ON DELETE CASCADE,
                    role TEXT NOT NULL,
                    CHECK (
                        (page_id IS NOT NULL AND section_id IS NULL AND notebook_id IS NULL) OR
                        (page_id IS NULL AND section_id IS NOT NULL AND notebook_id IS NULL) OR
                        (page_id IS NULL AND section_id IS NULL AND notebook_id IS NOT NULL)
                    )
                );
                CREATE INDEX IF NOT EXISTS idx_media_refs_media ON media_refs(media_id);
                CREATE INDEX IF NOT EXISTS idx_media_refs_page ON media_refs(page_id);
                CREATE INDEX IF NOT EXISTS idx_media_refs_section ON media_refs(section_id);
                CREATE INDEX IF NOT EXISTS idx_media_refs_notebook ON media_refs(notebook_id);
                """
            )
            # Initialize db_metadata uuid if missing
            cur.execute("SELECT uuid FROM db_metadata WHERE id=1")
            row = cur.fetchone()
            if not row or not row[0]:
                cur.execute(
                    "INSERT OR REPLACE INTO db_metadata(id, uuid) VALUES (1, ?)",
                    (str(uuid.uuid4()),),
                )
            conn.commit()
        finally:
            try:
                conn.close()
            except Exception:
                pass
        set_db_version(3, db_path)


def open_database(window):
    options = QtWidgets.QFileDialog.Options()
    # Default to the configured Databases root
    try:
        import os

        from settings_manager import get_databases_root

        initial_dir = get_databases_root()
    except Exception:
        initial_dir = ""
    file_name, _ = QtWidgets.QFileDialog.getOpenFileName(
        window,
        "Open Database",
        initial_dir,
        "SQLite DB Files (*.db);;All Files (*)",
        options=options,
    )
    if not file_name:
        return
    migrate_database_if_needed(file_name)
    set_last_db(file_name)
    clear_last_state()
    # Force a clean restart so UI initializes with the opened database
    restart_application()


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


def main():
    # Suppress noisy SIP deprecation warning from PyQt5 about sipPyTypeDict
    warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*sipPyTypeDict.*")
    # Install a global exception hook so unexpected errors surface in a dialog instead of closing silently
    try:
        _install_global_excepthook()
    except Exception:
        pass
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
    # Ensure database is migrated before any queries
    try:
        migrate_database_if_needed(db_path)
    except Exception:
        pass
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
    except Exception:
        pass
    populate_notebook_names(window, db_path)
    setup_tab_sync(window)
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
                    # Blank area: offer New Binder
                    if item is None:
                        m = QtWidgets.QMenu(tree)
                        act_new = m.addAction("New Binder")
                        m.addSeparator()
                        act_collapse_all = m.addAction("Collapse All Binders")
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
                        return
                    # Top-level binder item
                    if item.parent() is None:
                        tree.setCurrentItem(item)
                        m = QtWidgets.QMenu(tree)
                        # Place 'New Section' at the very top, followed by a separator
                        act_new_section = m.addAction("New Section")
                        m.addSeparator()
                        # Binder operations
                        act_new = m.addAction("New Binder")
                        act_rename = m.addAction("Rename Binder")
                        act_delete = m.addAction("Delete Binder")
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
                        return
                    # Non top-level (section or page)
                    tree.setCurrentItem(item)
                    m = QtWidgets.QMenu(tree)
                    kind = item.data(0, 1001)
                    if kind == "section":
                        # Section menu
                        act_add_page = m.addAction("Add Page")
                        m.addSeparator()
                        act_new_section = m.addAction("New Section")
                        act_rename_section = m.addAction("Rename Section")
                        act_delete_section = m.addAction("Delete Section")
                        chosen = m.exec_(global_pos)
                        if chosen is None:
                            return
                        if chosen == act_add_page:
                            add_page(window)
                            return
                        if chosen == act_new_section:
                            add_section(window)
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
                        # Page menu
                        act_add_page = m.addAction("Add Page")
                        act_rename_page = m.addAction("Rename Page")
                        act_delete_page = m.addAction("Delete Page")
                        chosen = m.exec_(global_pos)
                        if chosen is None:
                            return
                        # Context: ids
                        page_id = item.data(0, 1000)
                        section_id = item.data(0, 1002)
                        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
                        if chosen == act_add_page:
                            add_page(window)
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
                                        from ui_tabs import USER_ROLE_ID, USER_ROLE_KIND
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
                            # Two-column: refresh section's children and load first page
                            try:
                                if _is_two_column_ui(window):
                                    # Determine notebook id for this section
                                    nb_id = getattr(window, "_current_notebook_id", None)
                                    if nb_id is None and section_id is not None:
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
                                    if nb_id is not None:
                                        ensure_left_tree_sections(
                                            window, int(nb_id), select_section_id=int(section_id) if section_id is not None else None
                                        )
                                    # Clear current if we deleted the active page, then load first page
                                    try:
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
    action_open = window.findChild(QtWidgets.QAction, "actionOpen")
    if action_open:
        action_open.triggered.connect(lambda: open_database(window))
    # Save As: copy current db and media to new path
    action_save_as = window.findChild(QtWidgets.QAction, "actionSave_As")
    if action_save_as:
        action_save_as.triggered.connect(lambda: save_database_as(window))
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
                            # Ensure section context and load first page (or clear)
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
                            from ui_tabs import _load_page_for_current_tab as _load_page

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
                                _load_page(window, int(page_id))
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
                    import os

                    # Load as a top-level QDialog with normal window chrome
                    ui_path = os.path.join(os.path.dirname(__file__), "settings_dialog.ui")
                    dlg = uic.loadUi(ui_path)
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

    window.show()

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
            from ui_tabs import save_current_page

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
                    from ui_tabs import save_current_page

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
