"""
main.py
Entry point for the NoteBook application. Handles main window setup, menu actions, database creation/opening, and application startup.
"""
import sys
import warnings
from PyQt5 import QtWidgets
from PyQt5.QtCore import QProcess, QTimer
from ui_loader import load_main_window
from ui_logic import populate_notebook_names
from ui_tabs import setup_tab_sync, restore_last_position, refresh_for_notebook, ensure_left_tree_sections
from settings_manager import set_last_state
from settings_manager import (
    get_last_db,
    set_last_db,
    get_window_geometry,
    set_window_geometry,
    get_window_maximized,
    set_window_maximized,
    clear_last_state,
)
from db_access import create_notebook as db_create_notebook, rename_notebook as db_rename_notebook, delete_notebook as db_delete_notebook
from db_sections import create_section as db_create_section, get_sections_by_notebook_id as db_get_sections_by_notebook_id
from db_pages import create_page as db_create_page
from db_pages import update_page_title as db_update_page_title

def create_new_database(window):
    options = QtWidgets.QFileDialog.Options()
    file_name, _ = QtWidgets.QFileDialog.getSaveFileName(window, "Create New Database", "", "SQLite DB Files (*.db);;All Files (*)", options=options)
    if not file_name:
        return
    import sqlite3
    conn = sqlite3.connect(file_name)
    cursor = conn.cursor()
    cursor.executescript('''
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
    ''')
    conn.commit()
    # Set version to 2 (includes sections.color_hex)
    cursor.execute('PRAGMA user_version = 2')
    conn.commit()
    conn.close()
    set_last_db(file_name)
    clear_last_state()
    # Force a clean restart so UI initializes with the new database
    restart_application()

def _select_left_tree_notebook(window, notebook_id: int):
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    if not tree_widget:
        return
    for i in range(tree_widget.topLevelItemCount()):
        top = tree_widget.topLevelItem(i)
        if top.data(0, 1000) == notebook_id:
            tree_widget.setCurrentItem(top)
            break

def add_binder(window):
    title, ok = QtWidgets.QInputDialog.getText(window, "Add Binder", "Binder title:", text="Untitled Binder")
    if not ok:
        return
    title = (title or "").strip() or "Untitled Binder"
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    # Capture current expanded state of top-level binders and persist before refresh
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
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

    # Create notebook and refresh UI
    nid = db_create_notebook(title, db_path)
    set_last_state(notebook_id=nid, section_id=None, page_id=None)
    populate_notebook_names(window, db_path)
    # Restore previously expanded binders (do not auto-expand the new one)
    try:
        from ui_tabs import ensure_left_tree_sections
        from settings_manager import get_expanded_notebooks
        persisted_ids = get_expanded_notebooks()
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
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
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    if not tree_widget:
        return
    item = tree_widget.currentItem()
    if item is None or item.parent() is not None:
        # fallback to first notebook
        item = tree_widget.topLevelItem(0) if tree_widget.topLevelItemCount() > 0 else None
    if item is None:
        QtWidgets.QMessageBox.information(window, "Rename Binder", "No binder selected.")
        return
    nid = item.data(0, 1000)
    current = item.text(0) or ""
    new_title, ok = QtWidgets.QInputDialog.getText(window, "Rename Binder", "New title:", text=current)
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
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    db_rename_notebook(int(nid), new_title.strip(), db_path)
    populate_notebook_names(window, db_path)
    # Restore expansion from persisted state
    try:
        from ui_tabs import ensure_left_tree_sections
        from settings_manager import get_expanded_notebooks
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
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    if not tree_widget:
        return
    item = tree_widget.currentItem()
    if item is None or item.parent() is not None:
        # If a section is selected, delete its parent binder; otherwise fallback to first binder
        if item is not None and item.parent() is not None:
            item = item.parent()
        else:
            item = tree_widget.topLevelItem(0) if tree_widget.topLevelItemCount() > 0 else None
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
    confirm = QtWidgets.QMessageBox.question(
        window,
        "Delete Binder",
        "Are you sure you want to delete this binder and all its sections and pages?",
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
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    db_delete_notebook(nid, db_path)
    # Clear any remembered state that points to this notebook
    clear_last_state()
    # Refresh UI: repopulate binders (selection will change shortly)
    populate_notebook_names(window, db_path)
    # Restore previously expanded binders (excluding the one we just deleted), based on persisted state
    try:
        from ui_tabs import ensure_left_tree_sections
        from settings_manager import get_expanded_notebooks
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
                    from ui_tabs import ensure_left_tree_sections
                    ensure_left_tree_sections(window, nb_id)
            except Exception:
                pass
            # Single unified refresh
            refresh_for_notebook(window, nb_id)
            # Fallback: if binder has sections but tabs are empty, force full UI refresh once
            try:
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                sections = db_get_sections_by_notebook_id(nb_id, db_path)
                if sections and (not tab_widget or tab_widget.count() == 0):
                    _full_ui_refresh(window)
                    refresh_for_notebook(window, nb_id)
            except Exception:
                pass
    else:
        # No binders left: clear tabs and right pane explicitly
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if tab_widget:
            tab_widget.clear()
        right_tw = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
        if right_tw:
            right_tw.clear()
        right_tv = window.findChild(QtWidgets.QTreeView, 'sectionPages')
        if right_tv and right_tv.model() is not None:
            right_tv.setModel(None)

def add_section(window):
    # Determine target notebook: current selection in left tree
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
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
    title, ok = QtWidgets.QInputDialog.getText(window, "Add Section", "Section title:", text="Untitled Section")
    if not ok:
        return
    title = (title or "").strip() or "Untitled Section"
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    sid = db_create_section(int(nb_id), title, db_path)
    # Preserve left-tree state: avoid full repopulate; refresh only the target binder children
    try:
        from ui_tabs import refresh_for_notebook, ensure_left_tree_sections
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
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    # Clear widgets
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    if tree_widget:
        tree_widget.clear()
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if tab_widget:
        tab_widget.clear()
    right_tw = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
    if right_tw:
        right_tw.clear()
    right_tv = window.findChild(QtWidgets.QTreeView, 'sectionPages')
    if right_tv and right_tv.model() is not None:
        right_tv.setModel(None)
    populate_notebook_names(window, db_path)
    setup_tab_sync(window)
    restore_last_position(window)
    # Prepare splitter stretch factors (favor center panel); apply sizes after show
    try:
        splitter = window.findChild(QtWidgets.QSplitter, 'mainSplitter')
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
    # Determine active section from the current tab or right pane selection
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    tab_bar = tab_widget.tabBar() if tab_widget else None
    section_id = None
    if tab_widget and tab_widget.count() > 0 and tab_bar is not None:
        idx = tab_widget.currentIndex()
        section_id = tab_bar.tabData(idx)
    if section_id is None:
        # Try right pane selection (QTreeWidget)
        right_tw = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
        if right_tw and right_tw.currentItem() is not None:
            cur = right_tw.currentItem()
            kind = cur.data(0, 1001)
            if kind == 'section':
                section_id = cur.data(0, 1000)
            elif kind == 'page':
                section_id = cur.data(0, 1002)
        # Try model view
        if section_id is None:
            right_tv = window.findChild(QtWidgets.QTreeView, 'sectionPages')
            if right_tv and right_tv.currentIndex().isValid():
                idx = right_tv.currentIndex()
                kind = idx.data(1001)
                if kind == 'section':
                    section_id = idx.data(1000)
                elif kind == 'page':
                    section_id = idx.data(1002)
    if section_id is None:
        QtWidgets.QMessageBox.information(window, "Add Page", "Please select or create a section first.")
        return
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    pid = db_create_page(int(section_id), "Untitled Page", db_path)
    # Ensure the UI reflects the new page; selecting the section will populate and load it
    set_last_state(section_id=int(section_id), page_id=pid)
    restore_last_position(window)

def _current_page_context(window):
    """Return (section_id, page_id) for the current tab/page if available."""
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
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
        QtWidgets.QMessageBox.information(window, "Insert Attachment", "Please open or create a page first.")
        return
    db_path = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
    options = QtWidgets.QFileDialog.Options()
    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(window, "Select Attachment", "", "All Files (*);;Images (*.png *.jpg *.jpeg *.gif *.bmp);;PDF (*.pdf)", options=options)
    if not file_path:
        return
    try:
        from media_store import save_file_into_store, add_media_ref
        media_id, rel_path = save_file_into_store(db_path, file_path)
        add_media_ref(db_path, media_id, page_id=page_id, role='attachment')
        QtWidgets.QMessageBox.information(window, "Insert Attachment", f"Attached file saved to media store.\n{rel_path}")
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
                cur.execute("INSERT OR REPLACE INTO db_metadata(id, uuid) VALUES (1, ?)", (str(uuid.uuid4()),))
            conn.commit()
        finally:
            try:
                conn.close()
            except Exception:
                pass
        set_db_version(3, db_path)

def open_database(window):
    options = QtWidgets.QFileDialog.Options()
    file_name, _ = QtWidgets.QFileDialog.getOpenFileName(window, "Open Database", "", "SQLite DB Files (*.db);;All Files (*)", options=options)
    if not file_name:
        return
    migrate_database_if_needed(file_name)
    set_last_db(file_name)
    clear_last_state()
    # Force a clean restart so UI initializes with the opened database
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
    app = QtWidgets.QApplication(sys.argv)
    window = load_main_window()
    # Restore window geometry and maximized state
    geom = get_window_geometry()
    if geom and all(k in geom for k in ('x','y','w','h')):
        window.setGeometry(int(geom['x']), int(geom['y']), int(geom['w']), int(geom['h']))
    if get_window_maximized():
        window.showMaximized()
    db_path = get_last_db() or 'notes.db'
    # Ensure database is migrated before any queries
    try:
        migrate_database_if_needed(db_path)
    except Exception:
        pass
    window._db_path = db_path
    # Prepare media root path for this database (not yet used by UI)
    try:
        from media_store import media_root_for_db, ensure_dir
        window._media_root = media_root_for_db(db_path)
        ensure_dir(window._media_root)
    except Exception:
        window._media_root = None
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
    # Apply default paste mode to override Ctrl+V behavior
    try:
        from settings_manager import get_default_paste_mode
        window._default_paste_mode = get_default_paste_mode()
    except Exception:
        window._default_paste_mode = 'rich'
    # Restore left-panel expanded binders from settings after initial build
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
        from settings_manager import get_expanded_notebooks
        from ui_tabs import ensure_left_tree_sections
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

    # Connect menu actions
    # Updated QAction name from UI: actionNew_Database
    action_newdb = window.findChild(QtWidgets.QAction, 'actionNew_Database')
    if action_newdb:
        action_newdb.triggered.connect(lambda: create_new_database(window))
    # Binder (notebook) actions
    act_add_wb_variants = [
        window.findChild(QtWidgets.QAction, 'actionAdd_WorkBook'),
        window.findChild(QtWidgets.QAction, 'actionAdd_Workbook'),
    ]
    for act in act_add_wb_variants:
        if act:
            act.triggered.connect(lambda: add_binder(window))
    act_rename_wb = window.findChild(QtWidgets.QAction, 'actionRename_WorkBook')
    if act_rename_wb:
        act_rename_wb.triggered.connect(lambda: rename_binder(window))
    act_delete_wb = window.findChild(QtWidgets.QAction, 'actionDelete_Workbook')
    if act_delete_wb:
        act_delete_wb.triggered.connect(lambda: delete_binder(window))
    action_open = window.findChild(QtWidgets.QAction, 'actionOpen')
    if action_open:
        action_open.triggered.connect(lambda: open_database(window))
    # Insert menu wiring for quick content creation
    act_add_section = window.findChild(QtWidgets.QAction, 'actionAdd_Scction')
    if act_add_section:
        act_add_section.triggered.connect(lambda: add_section(window))
    act_add_page = window.findChild(QtWidgets.QAction, 'actionAdd_Page')
    if act_add_page:
        act_add_page.triggered.connect(lambda: add_page(window))
    act_insert_attachment = window.findChild(QtWidgets.QAction, 'actionInsert_Attachment')
    if act_insert_attachment:
        act_insert_attachment.triggered.connect(lambda: insert_attachment(window))
    action_exit = window.findChild(QtWidgets.QAction, 'actionExit')
    if action_exit:
        action_exit.triggered.connect(window.close)

    # Edit: Paste actions
    try:
        act_paste_plain = window.findChild(QtWidgets.QAction, 'actionPaste_Text_Only')
        if act_paste_plain:
            def _paste_plain():
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
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
                        from ui_tabs import save_current_page
                        save_current_page(window)
                    except Exception:
                        pass
                except Exception:
                    pass
            act_paste_plain.triggered.connect(_paste_plain)
        act_paste_match = window.findChild(QtWidgets.QAction, 'actionPaste_and_Match_Style')
        if act_paste_match:
            def _paste_match():
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
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
                        from ui_tabs import save_current_page
                        save_current_page(window)
                    except Exception:
                        pass
                except Exception:
                    pass
            act_paste_match.triggered.connect(_paste_match)
        act_paste_clean = window.findChild(QtWidgets.QAction, 'actionPaste_Clean_Formatting')
        if act_paste_clean:
            def _paste_clean():
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
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
                        from ui_tabs import save_current_page
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
        am_rich = window.findChild(QtWidgets.QAction, 'actionPasteMode_Rich')
        am_text = window.findChild(QtWidgets.QAction, 'actionPasteMode_Text_Only')
        am_match = window.findChild(QtWidgets.QAction, 'actionPasteMode_Match_Style')
        am_clean = window.findChild(QtWidgets.QAction, 'actionPasteMode_Clean')
        group = None
        if am_rich and am_text and am_match and am_clean:
            group = QtWidgets.QActionGroup(window)
            group.setExclusive(True)
            for a in (am_rich, am_text, am_match, am_clean):
                a.setCheckable(True)
                group.addAction(a)
            # Reflect current mode
            mode = getattr(window, '_default_paste_mode', 'rich')
            if mode == 'rich':
                am_rich.setChecked(True)
            elif mode == 'text-only':
                am_text.setChecked(True)
            elif mode == 'match-style':
                am_match.setChecked(True)
            elif mode == 'clean':
                am_clean.setChecked(True)
            # Persist on change
            def _set_mode(m):
                try:
                    window._default_paste_mode = m
                    from settings_manager import set_default_paste_mode
                    set_default_paste_mode(m)
                except Exception:
                    pass
            am_rich.triggered.connect(lambda: _set_mode('rich'))
            am_text.triggered.connect(lambda: _set_mode('text-only'))
            am_match.triggered.connect(lambda: _set_mode('match-style'))
            am_clean.triggered.connect(lambda: _set_mode('clean'))
    except Exception:
        pass

    # Tools: Clean Unused Media
    try:
        act_clean_media = window.findChild(QtWidgets.QAction, 'actionClean_Unused_Media')
        if act_clean_media:
            def _do_clean_media():
                try:
                    from media_store import garbage_collect_unused_media
                    dbp = getattr(window, '_db_path', None) or get_last_db() or 'notes.db'
                    removed = garbage_collect_unused_media(dbp)
                    QtWidgets.QMessageBox.information(window, "Clean Unused Media", f"Removed {removed} unreferenced media file(s).")
                except Exception as e:
                    QtWidgets.QMessageBox.warning(window, "Clean Unused Media", f"Failed to clean media: {e}")
            act_clean_media.triggered.connect(_do_clean_media)
    except Exception:
        pass

    # Format > List Scheme (wired to actions defined in main_window_5.ui)
    try:
        def _apply_list_schemes(ordered=None, unordered=None):
            try:
                from ui_richtext import set_list_schemes
                from settings_manager import set_list_schemes_settings
                set_list_schemes(ordered=ordered, unordered=unordered)
                set_list_schemes_settings(ordered=ordered, unordered=unordered)
            except Exception:
                return
        act_ord_classic = window.findChild(QtWidgets.QAction, 'actionOrdered_Classic')
        if act_ord_classic:
            act_ord_classic.triggered.connect(lambda: _apply_list_schemes(ordered='classic'))
        act_ord_decimal = window.findChild(QtWidgets.QAction, 'actionOrdered_Decimal')
        if act_ord_decimal:
            act_ord_decimal.triggered.connect(lambda: _apply_list_schemes(ordered='decimal'))
        act_un_disc_cs = window.findChild(QtWidgets.QAction, 'actionUnordered_Disc_Circle_Square')
        if act_un_disc_cs:
            act_un_disc_cs.triggered.connect(lambda: _apply_list_schemes(unordered='disc-circle-square'))
        act_un_disc_only = window.findChild(QtWidgets.QAction, 'actionUnordered_Disc_Only')
        if act_un_disc_only:
            act_un_disc_only.triggered.connect(lambda: _apply_list_schemes(unordered='disc-only'))
    except Exception:
        pass

    window.show()

    # Restore splitter sizes after the window is shown to ensure geometry exists
    def _apply_saved_splitter_sizes():
        try:
            splitter = window.findChild(QtWidgets.QSplitter, 'mainSplitter')
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
                splitter.splitterMoved.connect(lambda pos, index: set_splitter_sizes(splitter.sizes()))
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
            splitter = window.findChild(QtWidgets.QSplitter, 'mainSplitter')
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
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                if not tab_widget:
                    return
                page = tab_widget.currentWidget()
                if not page:
                    return
                te = page.findChild(QtWidgets.QTextEdit)
                if not te:
                    return
                mode = getattr(window, '_default_paste_mode', 'rich')
                if mode == 'text-only':
                    from ui_richtext import paste_text_only
                    paste_text_only(te)
                elif mode == 'match-style':
                    from ui_richtext import paste_match_style
                    paste_match_style(te)
                elif mode == 'clean':
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