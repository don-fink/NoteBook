"""
ui_tabs.py
Keeps the tab widget (tabPages) and the right-side section/pages tree in sync.
- Tabs represent Sections of the selected Notebook
- Right column (QTreeWidget: sectionPages) lists all sections of the active notebook with their pages as children
- Selecting a tab or a section in either tree switches the active section
- Selecting a page in the right tree loads that page in the editor
"""
import warnings as _warnings
_warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*sipPyTypeDict.*")
from PyQt5 import QtWidgets
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtCore import QModelIndex, Qt, QObject, QEvent, QTimer
from PyQt5.QtWidgets import QAbstractItemView
from db_sections import (
    get_sections_by_notebook_id,
    update_section_color,
    get_section_color_map,
    create_section,
    rename_section,
    delete_section,
    move_section_up,
    move_section_down,
    set_sections_order,
)
from db_pages import (
    get_pages_by_section_id,
    create_page,
    update_page_title,
    update_page_content,
    delete_page,
    set_pages_order,
    set_pages_parent_and_order,
)
from settings_manager import set_last_state, get_last_state
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QKeySequence
from PyQt5 import uic
import os
from ui_richtext import add_rich_text_toolbar
from ui_richtext import sanitize_html_for_storage
import sqlite3

# Roles for storing ids/kinds in tree items and models
USER_ROLE_ID = 1000  # same role used to store ids on tree items
USER_ROLE_KIND = 1001  # 'section' or 'page'
USER_ROLE_PARENT_SECTION = 1002  # section_id for page items

# Special sentinel used for an "+ New Section" tab when a binder has no sections
ADD_SECTION_SENTINEL = "__ADD_SECTION__"

def _is_add_section_sentinel(value) -> bool:
    return isinstance(value, str) and value == ADD_SECTION_SENTINEL

# Lightweight debug toggle and helper for this module
DEBUG_UI_TABS = False
def _dbg_print(*args, **kwargs):
    if DEBUG_UI_TABS:
        try:
            print(*args, **kwargs)
        except Exception:
            pass

def _ensure_attrs(window):
    if not hasattr(window, "_tabs_setup_done"):
        window._tabs_setup_done = False
    if not hasattr(window, "_db_path"):
        window._db_path = "notes.db"
    if not hasattr(window, "_suppress_sync"):
        window._suppress_sync = False
    if not hasattr(window, "_current_notebook_id"):
        window._current_notebook_id = None
    if not hasattr(window, "_current_page_by_section"):
        window._current_page_by_section = {}
    if not hasattr(window, "_last_tab_index"):
        window._last_tab_index = -1
    if not hasattr(window, "_refresh_generation"):
        window._refresh_generation = 0
    # One-shot guard to skip left-tree section selection (used when user clicks a Binder)
    if not hasattr(window, "_skip_left_tree_selection_once"):
        window._skip_left_tree_selection_once = False
    # One-shot guard to keep right-tree selection on the section when that section was clicked
    if not hasattr(window, "_keep_right_tree_section_selected_once"):
        window._keep_right_tree_section_selected_once = False

def _prompt_new_section(window):
    try:
        nb_id = getattr(window, '_current_notebook_id', None)
        if nb_id is None:
            return
        title, ok = QtWidgets.QInputDialog.getText(window, "New Section", "Section title:", text="Untitled Section")
        if not ok:
            return
        title = (title or "").strip() or "Untitled Section"
        new_id = create_section(nb_id, title, window._db_path)
        refresh_for_notebook(window, nb_id, select_section_id=new_id)
        try:
            _refresh_left_tree_children(window, nb_id, select_section_id=new_id)
        except Exception:
            pass
    except Exception:
        pass


def setup_tab_sync(window):
    """Connect handlers once and use window._db_path for queries."""
    _ensure_attrs(window)
    if window._tabs_setup_done:
        return

    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    right_tree = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
    right_view = window.findChild(QtWidgets.QTreeView, 'sectionPages')
    if not tree_widget or not tab_widget:
        return
    # Ensure the tab bar is visible even if there's only one tab
    try:
        tab_widget.setTabBarAutoHide(False)
    except Exception:
        pass
    # Take ownership of left tree signals by disconnecting existing connections if any
    try:
        tree_widget.itemClicked.disconnect()
    except Exception:
        pass
    # Ensure right/left clicks select the item under cursor to avoid flaky selection
    try:
        class _LeftTreeSelectFilter(QObject):
            def eventFilter(self, obj, event):
                try:
                    if event.type() == QEvent.MouseButtonPress:
                        tw = obj if isinstance(obj, QtWidgets.QTreeWidget) else obj.parent()
                        if isinstance(tw, QtWidgets.QTreeWidget):
                            pos = event.pos() if hasattr(obj, 'pos') else event.pos()
                            # Map to viewport if needed
                            if hasattr(tw, 'viewport') and obj is tw:
                                vp = tw.viewport()
                                pos = vp.mapFrom(tw, pos)
                                item = tw.itemAt(pos)
                            else:
                                item = tw.itemAt(event.pos())
                            if item is not None:
                                tw.setCurrentItem(item)
                except Exception:
                    pass
                return False

        if not hasattr(window, '_left_tree_select_filter'):
            window._left_tree_select_filter = _LeftTreeSelectFilter()
            tree_widget.installEventFilter(window._left_tree_select_filter)
            if hasattr(tree_widget, 'viewport') and tree_widget.viewport() is not None:
                tree_widget.viewport().installEventFilter(window._left_tree_select_filter)
    except Exception:
        pass
    # Left tree context menu for section CRUD
    try:
        tree_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        tree_widget.customContextMenuRequested.connect(lambda pos: _on_left_tree_context_menu(window, tree_widget, pos))
    except Exception:
        pass
    # Persist expand/collapse state of binders for future sessions
    try:
        def _on_item_expanded(item):
            if item is not None and item.parent() is None:
                from settings_manager import add_expanded_notebook
                nid = item.data(0, USER_ROLE_ID)
                if nid is not None:
                    add_expanded_notebook(int(nid))
        def _on_item_collapsed(item):
            if item is not None and item.parent() is None:
                from settings_manager import remove_expanded_notebook
                nid = item.data(0, USER_ROLE_ID)
                if nid is not None:
                    remove_expanded_notebook(int(nid))
        tree_widget.itemExpanded.connect(_on_item_expanded)
        tree_widget.itemCollapsed.connect(_on_item_collapsed)
    except Exception:
        pass

    # Enable context menu on tab bar for setting colors
    tab_bar = tab_widget.tabBar()
    if tab_bar is not None:
        tab_bar.setContextMenuPolicy(Qt.CustomContextMenu)
        tab_bar.customContextMenuRequested.connect(lambda pos: _on_tab_bar_context_menu(window, pos))
        # Enable drag to reorder section tabs
        try:
            tab_bar.setMovable(True)
            def _on_tab_moved(_from, _to):
                _persist_tab_order(window)
            tab_bar.tabMoved.connect(_on_tab_moved)
        except Exception:
            pass
        # Clicking the current sentinel tab should still open the New Section prompt
        try:
            def _on_tab_bar_clicked(idx:int):
                tb = tab_widget.tabBar()
                if tb is None:
                    return
                sid = tb.tabData(idx)
                if _is_add_section_sentinel(sid):
                    _prompt_new_section(window)
            tab_bar.tabBarClicked.connect(_on_tab_bar_clicked)
        except Exception:
            pass

    # Ctrl+S saves current page
    try:
        QtWidgets.QShortcut(QKeySequence.Save, window, activated=lambda: save_current_page(window))
    except Exception:
        pass

    def on_tree_item_clicked(item, column):
        if getattr(window, "_suppress_sync", False):
            return
        # Save current page before switching context
        save_current_page(window)
        # Notebook item
        if item.parent() is None:
            notebook_id = item.data(0, USER_ROLE_ID)
            if notebook_id is None:
                return
            # Persist notebook and rebuild UI explicitly
            set_last_state(notebook_id=notebook_id)
            window._current_notebook_id = notebook_id
            # Ensure this binder's sections are visible under the clicked item
            try:
                _refresh_left_tree_children(window, notebook_id)
                item.setExpanded(True)
            except Exception:
                pass
            # Unified, deferred refresh to avoid reentrancy and ensure selection.
            # Keep the left tree focused on the binder (don't immediately jump selection to a section).
            try:
                window._skip_left_tree_selection_once = True
            except Exception:
                pass
            refresh_for_notebook(window, notebook_id, keep_left_tree_selection=True)
        else:
            # Section item: select matching tab and load first page
            section_id = item.data(0, USER_ROLE_ID)
            if section_id is None:
                return
            # If the clicked section belongs to a different notebook than the current tabs,
            # switch notebooks first and then select the section's tab.
            parent_item = item.parent()
            parent_notebook_id = parent_item.data(0, USER_ROLE_ID) if parent_item is not None else None
            if parent_notebook_id is not None and parent_notebook_id != getattr(window, "_current_notebook_id", None):
                set_last_state(notebook_id=parent_notebook_id)
                _populate_tabs_for_notebook(window, parent_notebook_id)
                _build_right_tree_for_notebook(window, parent_notebook_id)
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, section_id)
            window._suppress_sync = False
            _load_first_page_for_current_tab(window)
            _select_right_tree_section(window, section_id)

    def on_tab_changed(index):
        if getattr(window, "_suppress_sync", False):
            window._last_tab_index = index
            return
        if index < 0:
            return
        # Diagnostics to ensure signal fires and tab widget is enabled
        try:
            tw = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            tb = tw.tabBar() if tw else None
            sid_dbg = tb.tabData(index) if tb is not None else None
            # print(f"[ui_tabs] on_tab_changed idx={index}, sid={sid_dbg}")
            if tw is not None:
                tw.setEnabled(True)
                if tb is not None:
                    tb.setEnabled(True)
        except Exception:
            pass
        # Save previous tab page before switching
        if window._last_tab_index is not None and window._last_tab_index != -1 and window._last_tab_index != index:
            _save_page_for_tab_index(window, window._last_tab_index)
        tab_bar = tab_widget.tabBar()
        # Intercept special "+ New Section" sentinel tab
        if tab_bar is not None:
            section_id = tab_bar.tabData(index)
            if _is_add_section_sentinel(section_id):
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is None:
                    window._last_tab_index = index
                    return
                title, ok = QtWidgets.QInputDialog.getText(tab_widget, "New Section", "Section title:", text="Untitled Section")
                if ok:
                    title = (title or "").strip() or "Untitled Section"
                    new_id = create_section(nb_id, title, window._db_path)
                    refresh_for_notebook(window, nb_id, select_section_id=new_id)
                window._last_tab_index = index
                return
        # Normal tab change: load, reflect selection and persist last section
        _load_first_page_for_current_tab(window)
        if tab_bar is not None:
            section_id = tab_bar.tabData(index)
            if section_id is not None and not _is_add_section_sentinel(section_id):
                # Respect one-shot guard: when switching binders, keep binder selected in left tree
                if not getattr(window, "_skip_left_tree_selection_once", False):
                    _select_tree_section(window, section_id)
                try:
                    nb_id = getattr(window, '_current_notebook_id', None)
                    if nb_id is not None:
                        set_last_state(notebook_id=nb_id, section_id=section_id)
                    else:
                        set_last_state(section_id=section_id)
                except Exception:
                    pass
                _select_right_tree_section(window, section_id)
        window._last_tab_index = index
    def on_right_tree_clicked(item, column):
        if getattr(window, "_suppress_sync", False):
            return
        if item is None:
            return
        # Save edits before switching
        save_current_page(window)
        kind = item.data(0, USER_ROLE_KIND)
        if kind == 'section':
            section_id = item.data(0, USER_ROLE_ID)
            if section_id is None:
                return
            # Select the corresponding tab; first page will load via tab change handler
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, section_id)
            window._suppress_sync = False
            # Keep the right tree's selection on the section for this user action
            try:
                window._keep_right_tree_section_selected_once = True
            except Exception:
                pass
            _load_first_page_for_current_tab(window)
            # Explicitly reselect the section in the right tree to ensure it stays highlighted
            try:
                _select_right_tree_section(window, section_id)
                # Also sync the left panel to the same section
                _select_tree_section(window, section_id)
            except Exception:
                pass
            try:
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    set_last_state(notebook_id=nb_id, section_id=section_id)
                else:
                    set_last_state(section_id=section_id)
            except Exception:
                pass
        elif kind == 'page':
            page_id = item.data(0, USER_ROLE_ID)
            parent_section_id = item.data(0, USER_ROLE_PARENT_SECTION)
            if parent_section_id is None or page_id is None:
                return
            # Ensure the correct section/tab is active
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, parent_section_id)
            window._suppress_sync = False
            _load_page_for_current_tab(window, page_id)
            # Sync left panel to the parent section of the selected page
            try:
                _select_tree_section(window, parent_section_id)
            except Exception:
                pass
            try:
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    set_last_state(notebook_id=nb_id, section_id=parent_section_id, page_id=page_id)
                else:
                    set_last_state(section_id=parent_section_id, page_id=page_id)
            except Exception:
                pass

    tree_widget.itemClicked.connect(on_tree_item_clicked)
    tab_widget.currentChanged.connect(on_tab_changed)
    if right_tree is not None:
        right_tree.itemClicked.connect(on_right_tree_clicked)
        right_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        right_tree.customContextMenuRequested.connect(lambda pos: _on_right_tree_context_menu(window, right_tree, pos))
        # Enable drag/drop reorder for pages, including cross-section reparenting; prevent nesting under pages
        try:
            right_tree.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
            right_tree.setDragDropMode(QAbstractItemView.InternalMove)
            right_tree.setDefaultDropAction(Qt.MoveAction)
            right_tree.setAcceptDrops(True)
            right_tree.setDragEnabled(True)
            right_tree.setDropIndicatorShown(True)
            if hasattr(right_tree, 'viewport') and right_tree.viewport() is not None:
                right_tree.viewport().setAcceptDrops(True)
            # Install a drop event filter to persist order after drop
            if not hasattr(window, '_right_tree_dnd_filter'):
                window._right_tree_dnd_filter = _RightTreeDnDFilter(window)
                # Install on both the widget and its viewport to ensure events are captured
                right_tree.installEventFilter(window._right_tree_dnd_filter)
                if hasattr(right_tree, 'viewport') and right_tree.viewport() is not None:
                    right_tree.viewport().installEventFilter(window._right_tree_dnd_filter)
        except Exception:
            pass
    if right_view is not None:
        right_view.clicked.connect(lambda index: _on_right_view_clicked(window, index))
        right_view.setContextMenuPolicy(Qt.CustomContextMenu)
        right_view.customContextMenuRequested.connect(lambda pos: _on_right_view_context_menu(window, right_view, pos))
        # Enable drag/drop reorder for pages within a section in model view
        try:
            right_view.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
            right_view.setDragDropMode(QAbstractItemView.InternalMove)
            right_view.setDefaultDropAction(Qt.MoveAction)
            right_view.setAcceptDrops(True)
            right_view.setDragEnabled(True)
            right_view.setDropIndicatorShown(True)
            if hasattr(right_view, 'viewport') and right_view.viewport() is not None:
                right_view.viewport().setAcceptDrops(True)
            if not hasattr(window, '_right_view_dnd_filter'):
                window._right_view_dnd_filter = _RightViewDnDFilter(window)
                # Install on both the view and its viewport to capture drop events
                right_view.installEventFilter(window._right_view_dnd_filter)
                if hasattr(right_view, 'viewport') and right_view.viewport() is not None:
                    right_view.viewport().installEventFilter(window._right_view_dnd_filter)
        except Exception:
            pass
    window._tabs_setup_done = True


def select_tab_for_section(window, section_id):
    """Public helper: select the tab that corresponds to section_id."""
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if not tab_widget:
        return
    _select_tab_for_section(tab_widget, section_id)


def load_first_page_for_current_tab(window):
    """Public helper: load the first page for the currently selected tab."""
    _load_first_page_for_current_tab(window)

def ensure_left_tree_sections(window, notebook_id: int, select_section_id: int = None):
    """Public helper: ensure the left tree shows sections under the given notebook and expand it."""
    try:
        _refresh_left_tree_children(window, notebook_id, select_section_id=select_section_id)
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
        if not tree_widget:
            return
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            if top.data(0, USER_ROLE_ID) == notebook_id:
                top.setExpanded(True)
                break
    except Exception:
        pass

def refresh_for_notebook(window, notebook_id: int, select_section_id: int = None, keep_left_tree_selection: bool = False):
    """Public helper: rebuild tabs and right pane for a notebook and optionally select a section.
    Uses a deferred finalize step to ensure the UI updates after model changes.
    """
    # Ensure core attrs and reset any stuck suppression so interactions work
    _ensure_attrs(window)
    try:
        window._suppress_sync = False
    except Exception:
        pass
    # Track the active notebook context for downstream actions and persistence
    try:
        window._current_notebook_id = notebook_id
        set_last_state(notebook_id=notebook_id)
    except Exception:
        pass
    # Bump generation to invalidate any prior scheduled finalizers
    try:
        window._refresh_generation += 1
    except Exception:
        window._refresh_generation = 1
    generation = window._refresh_generation
    # If we want to keep the binder highlighted in the left tree (when user clicked a Binder),
    # set a one-shot guard so tabChanged handlers won't override the left selection.
    if keep_left_tree_selection:
        try:
            window._skip_left_tree_selection_once = True
        except Exception:
            pass
    _populate_tabs_for_notebook(window, notebook_id)
    _build_right_tree_for_notebook(window, notebook_id)
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if not tab_widget:
        return
    tab_widget.setVisible(True)
    try:
        tab_widget.setEnabled(True)
        if tab_widget.tabBar() is not None:
            tab_widget.tabBar().setEnabled(True)
    except Exception:
        pass

    def _finalize():
        # Drop if a newer refresh was scheduled
        if getattr(window, "_refresh_generation", 0) != generation:
            return
        try:
            sections = get_sections_by_notebook_id(notebook_id, window._db_path)
        except Exception:
            sections = []
        if sections and tab_widget.count() == 0:
            # Try repopulating once more if tabs are still empty
            _populate_tabs_for_notebook(window, notebook_id)
        # Perform selection
        if select_section_id is not None:
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, select_section_id)
            window._suppress_sync = False
            _select_tree_section(window, select_section_id)
            _select_right_tree_section(window, select_section_id)
            _load_first_page_for_current_tab(window)
        else:
            if tab_widget.count() > 0:
                # Prefer first real section tab (skip sentinel)
                first_idx = 0
                try:
                    tb = tab_widget.tabBar()
                    if tb is not None:
                        for i in range(tab_widget.count()):
                            data = tb.tabData(i)
                            if data is not None and not _is_add_section_sentinel(data):
                                first_idx = i
                                break
                except Exception:
                    pass
                window._suppress_sync = True
                tab_widget.setCurrentIndex(first_idx)
                window._suppress_sync = False
                _load_first_page_for_current_tab(window)
                # Reflect selection in left and right trees
                try:
                    tb = tab_widget.tabBar()
                    sid = tb.tabData(first_idx) if tb is not None else None
                    if sid is not None and not _is_add_section_sentinel(sid):
                        # Only change left tree selection if not instructed to keep binder selection
                        if not keep_left_tree_selection:
                            _select_tree_section(window, sid)
                        _select_right_tree_section(window, sid)
                except Exception:
                    pass
        # Ensure visibility and raise the tab widget
        try:
            tab_widget.setTabBarAutoHide(False)
            tab_widget.show()
            tab_widget.raise_()
            tb = tab_widget.tabBar()
            if tb is not None:
                tb.show()
        except Exception:
            pass
        # Debug: print tab count after finalize
        try:
            tb_cnt = tab_widget.count()
            tb = tab_widget.tabBar()
            cur = tab_widget.currentIndex()
            sid = tb.tabData(cur) if (tb is not None and cur >= 0) else None
            vis = tab_widget.isVisible()
            geo = tab_widget.geometry()
            # print(f"[ui_tabs] finalize nb={notebook_id}, tabs={tb_cnt}, cur={cur}, sid={sid}, visible={vis}, h={geo.height()}")
        except Exception:
            pass
        # If somehow collapsed, try to enforce a minimal size and relayout
        try:
            if tab_widget.height() <= 1:
                tab_widget.setMinimumHeight(120)
                tab_widget.updateGeometry()
                tw_parent = tab_widget.parentWidget()
                if tw_parent is not None:
                    tw_parent.updateGeometry()
                window.updateGeometry()
                window.repaint()
        except Exception:
            pass
        tab_widget.update()
        # Clear any one-shot left-tree skip to avoid affecting future user-driven tab changes
        try:
            window._skip_left_tree_selection_once = False
        except Exception:
            pass

    # Defer finalize to the next event loop cycle
    QTimer.singleShot(0, _finalize)


def _populate_tabs_for_notebook(window, notebook_id):
    """Create one tab per section under the given notebook."""
    _ensure_attrs(window)
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if not tab_widget:
        return
    # print(f"[ui_tabs] clear tabs nb={notebook_id}")
    tab_widget.clear()
    tab_widget.setVisible(True)

    sections = get_sections_by_notebook_id(notebook_id, window._db_path)
    # print(f"[ui_tabs] populate tabs for nb={notebook_id}, sections={len(sections)}")
    created_tabs = 0
    for section in sections:
        # section: (id, notebook_id, title, ...)
        section_id = section[0]
        section_title = str(section[2])
        # Load the tab page UI file once
        # Resolve tab_page.ui robustly
        tab_ui_path = os.path.join(os.path.dirname(__file__), 'tab_page.ui')
        if not os.path.exists(tab_ui_path):
            try:
                tab_ui_path = os.path.abspath('tab_page.ui')
            except Exception:
                pass
        # Create the tab content from UI file (centralized styling)
        try:
            tab = uic.loadUi(tab_ui_path)
        except Exception:
            # Fallback to a simple container if UI fails to load
            tab = QtWidgets.QWidget()
            layout = QtWidgets.QVBoxLayout(tab)
            title_edit = QtWidgets.QLineEdit(tab)
            title_edit.setObjectName('pageTitleEdit')
            title_edit.setPlaceholderText('Page title')
            layout.addWidget(title_edit)
            text_edit = QtWidgets.QTextEdit(tab)
            text_edit.setObjectName('textEdit')
            layout.addWidget(text_edit)
        index = tab_widget.addTab(tab, section_title)
        # Debug
        _dbg_print(f"[ui_tabs] add tab sid={section_id}, title='{section_title}', idx={index}")
        # Store section_id in the QTabBar's tab data
        tab_bar = tab_widget.tabBar()
        if tab_bar is not None:
            tab_bar.setTabData(index, section_id)
        # Connect title edit commit to rename current page of this section
        def _bind_commit(te_tab=tab):
            def _commit():
                _commit_title_edit(window, te_tab)
            return _commit
        title_edit = tab.findChild(QtWidgets.QLineEdit, 'pageTitleEdit')
        if title_edit is not None:
            title_edit.editingFinished.connect(_bind_commit())
        try:
            if title_edit is not None:
                title_edit.returnPressed.connect(_bind_commit())
        except Exception:
            pass
        # Hook text change to show modified indicator
        text_edit = tab.findChild(QtWidgets.QTextEdit, 'textEdit')
        if text_edit is not None:
            def _on_text_changed(te_tab=tab, sid=section_id):
                try:
                    # If user starts editing in a section with no pages yet, auto-create a page
                    current_pid = getattr(window, "_current_page_by_section", {}).get(sid)
                    if not current_pid:
                        te_local = te_tab.findChild(QtWidgets.QTextEdit, 'textEdit')
                        if te_local is not None:
                            plain = (te_local.toPlainText() or "").strip()
                            placeholder = "No pages in this section yet."
                            # Create a new page when there is user content beyond the placeholder
                            if plain and plain != placeholder:
                                new_pid = create_page(sid, "Untitled Page", window._db_path)
                                try:
                                    window._current_page_by_section[sid] = new_pid
                                except Exception:
                                    pass
                                # Enable title edit and set default title
                                try:
                                    title_edit = te_tab.findChild(QtWidgets.QLineEdit, 'pageTitleEdit')
                                    if title_edit is not None:
                                        title_edit.setEnabled(True)
                                        title_edit.setText("Untitled Page")
                                except Exception:
                                    pass
                                # If placeholder text is present, remove it while preserving user input
                                try:
                                    full_plain = te_local.toPlainText()
                                    if placeholder in full_plain:
                                        cleaned = full_plain.replace(placeholder, "", 1).lstrip("\n\r ")
                                        te_local.blockSignals(True)
                                        te_local.setPlainText(cleaned)
                                        te_local.blockSignals(False)
                                except Exception:
                                    pass
                                # Reflect in right pane and selection
                                try:
                                    nb_id = getattr(window, '_current_notebook_id', None)
                                    if nb_id is not None:
                                        _build_right_tree_for_notebook(window, nb_id)
                                    _select_right_tree_page(window, sid, new_pid)
                                    set_last_state(section_id=sid, page_id=new_pid)
                                except Exception:
                                    pass
                except Exception:
                    pass
                _set_modified_indicator_for_tab(te_tab, True)
            text_edit.textChanged.connect(_on_text_changed)
            # Attach a compact rich text toolbar at the top of the tab
            try:
                add_rich_text_toolbar(tab, text_edit)
            except Exception:
                pass
        created_tabs += 1
    # If no sections exist, add a "+ New Section" sentinel tab as an affordance
    if created_tabs == 0:
        add_tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(add_tab)
        layout.setContentsMargins(16, 16, 16, 16)
        label = QtWidgets.QLabel("No sections yet. Click this tab to create your first section.")
        label.setWordWrap(True)
        layout.addWidget(label)
        try:
            btn = QtWidgets.QPushButton("New Section")
            btn.clicked.connect(lambda: _prompt_new_section(window))
            layout.addWidget(btn)
        except Exception:
            pass
        idx = tab_widget.addTab(add_tab, "+ New Section")
        tab_bar = tab_widget.tabBar()
        if tab_bar is not None:
            tab_bar.setTabData(idx, ADD_SECTION_SENTINEL)
    # After creating tabs, apply color icons if any
    color_map = get_section_color_map(notebook_id, window._db_path)
    try:
        tab_bar = tab_widget.tabBar()
        if tab_bar is not None and color_map:
            for i in range(tab_widget.count()):
                sid = tab_bar.tabData(i)
                if sid is not None and not _is_add_section_sentinel(sid):
                    col = color_map.get(int(sid))
                    if col:
                        _apply_tab_color(tab_widget, i, col)
    except Exception:
        pass
def _select_tab_for_section(tab_widget, section_id):
    tab_bar = tab_widget.tabBar()
    if tab_bar is None:
        return
    for i in range(tab_widget.count()):
        if tab_bar.tabData(i) == section_id:
            tab_widget.setCurrentIndex(i)
            return


def _load_first_page_for_current_tab(window):
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if not tab_widget or tab_widget.count() == 0:
        return
    idx = tab_widget.currentIndex()
    tab_bar = tab_widget.tabBar()
    section_id = tab_bar.tabData(idx) if tab_bar is not None else None
    if section_id is None or _is_add_section_sentinel(section_id):
        return

    # Get first page for the section; prefer order_index then id
    pages = get_pages_by_section_id(section_id, window._db_path)
    if pages:
        # Try to sort by order_index (index 6) then id (index 0) if present
        try:
            pages_sorted = sorted(pages, key=lambda p: (p[6], p[0]))
        except Exception:
            pages_sorted = pages
        first = pages_sorted[0]
        html = first[3] or ""
        # Persist last page id and track current page per section
        try:
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                set_last_state(notebook_id=nb_id, section_id=section_id, page_id=first[0])
            else:
                set_last_state(section_id=section_id, page_id=first[0])
        except Exception:
            pass
        try:
            window._current_page_by_section[section_id] = first[0]
        except Exception:
            pass
    else:
        html = "<i>No pages in this section yet.</i>"

    # Set HTML and title into widgets
    tab = tab_widget.widget(idx)
    text_edit = tab.findChild(QtWidgets.QTextEdit, 'textEdit')
    title_edit = tab.findChild(QtWidgets.QLineEdit, 'pageTitleEdit')
    if text_edit:
        try:
            text_edit.blockSignals(True)
            # Normalize HTML by injecting default font family/size into root style to avoid Qt fallback to 8pt
            try:
                from ui_richtext import DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT
                if isinstance(html, str) and '<html' in html.lower():
                    # Inject a style into the head/body if none exists
                    if '<body' in html.lower():
                        html = html.replace('<body', f'<body style="font-family: {DEFAULT_FONT_FAMILY}; font-size: {int(DEFAULT_FONT_SIZE_PT)}pt"', 1)
                    else:
                        # simple prepend
                        html = f'<div style="font-family: {DEFAULT_FONT_FAMILY}; font-size: {int(DEFAULT_FONT_SIZE_PT)}pt">{html}</div>'
            except Exception:
                pass
            text_edit.setHtml(html)
        finally:
            text_edit.blockSignals(False)
        # Reset modified indicator after programmatic load
        _set_modified_indicator_for_tab(tab, False)
        # Also set document default font to ensure consistent fallback
        try:
            from PyQt5.QtGui import QFont
            from ui_richtext import DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT
            text_edit.document().setDefaultFont(QFont(DEFAULT_FONT_FAMILY, int(DEFAULT_FONT_SIZE_PT)))
        except Exception:
            pass
    if title_edit:
        if pages:
            try:
                title_edit.setEnabled(True)
                title_edit.setText(str(first[2]) if first[2] else "Untitled Page")
            except Exception:
                pass
        else:
            title_edit.setEnabled(False)
            title_edit.setText("")
    # Reflect in right tree selection if there is a page, unless the user explicitly clicked a section
    if pages:
        if getattr(window, "_keep_right_tree_section_selected_once", False):
            try:
                window._keep_right_tree_section_selected_once = False
            except Exception:
                pass
        else:
            _select_right_tree_page(window, section_id, first[0])


def _load_page_for_current_tab(window, page_id):
    """Load a specific page by id for the current tab's section; fallback to first if not found."""
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if not tab_widget or tab_widget.count() == 0:
        return
    idx = tab_widget.currentIndex()
    tab_bar = tab_widget.tabBar()
    section_id = tab_bar.tabData(idx) if tab_bar is not None else None
    if section_id is None or _is_add_section_sentinel(section_id):
        return
    pages = get_pages_by_section_id(section_id, window._db_path)
    page = None
    if pages:
        for p in pages:
            if p[0] == page_id:
                page = p
                break
        if page is None:
            # fallback to first sorted
            try:
                pages = sorted(pages, key=lambda p: (p[6], p[0]))
            except Exception:
                pass
            page = pages[0]
    html = page[3] if page else "<i>No pages in this section yet.</i>"
    tab = tab_widget.widget(idx)
    text_edit = tab.findChild(QtWidgets.QTextEdit, 'textEdit')
    title_edit = tab.findChild(QtWidgets.QLineEdit, 'pageTitleEdit')
    if text_edit:
        try:
            text_edit.blockSignals(True)
            text_edit.setHtml(html)
        finally:
            text_edit.blockSignals(False)
        # Reset modified indicator after programmatic load
        _set_modified_indicator_for_tab(tab, False)
        # Ensure document default font is set for consistent fallback
        try:
            from PyQt5.QtGui import QFont
            from ui_richtext import DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT
            text_edit.document().setDefaultFont(QFont(DEFAULT_FONT_FAMILY, int(DEFAULT_FONT_SIZE_PT)))
        except Exception:
            pass
    if title_edit:
        if page:
            title_edit.setEnabled(True)
            try:
                title_edit.setText(str(page[2]) if page[2] else "Untitled Page")
            except Exception:
                pass
        else:
            title_edit.setEnabled(False)
            title_edit.setText("")
    if page:
        try:
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                set_last_state(notebook_id=nb_id, section_id=section_id, page_id=page[0])
            else:
                set_last_state(page_id=page[0])
        except Exception:
            pass
        try:
            window._current_page_by_section[section_id] = page[0]
        except Exception:
            pass
        _select_right_tree_page(window, section_id, page[0])


def _commit_title_edit(window, tab_widget_page: QtWidgets.QWidget):
    """Commit the page title from the tab's line edit to the database and refresh right panel."""
    try:
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if tab_widget is None:
            return
        idx = tab_widget.indexOf(tab_widget_page)
        if idx < 0:
            return
        tab_bar = tab_widget.tabBar()
        section_id = tab_bar.tabData(idx) if tab_bar is not None else None
        if section_id is None:
            return
        page_id = getattr(window, "_current_page_by_section", {}).get(section_id)
        if not page_id:
            return
        title_edit = tab_widget_page.findChild(QtWidgets.QLineEdit, 'pageTitleEdit')
        if title_edit is None:
            return
        new_title = (title_edit.text() or "").strip() or "Untitled Page"
        update_page_title(page_id, new_title, window._db_path)
        # Rebuild right tree and reselect this page so label updates
        current_nb = getattr(window, '_current_notebook_id', None)
        if current_nb is not None:
            _build_right_tree_for_notebook(window, current_nb)
        _select_right_tree_page(window, section_id, page_id)
    except Exception:
        pass


def _select_tree_section(window, section_id):
    """Select the corresponding section item in the tree without collapsing tabs."""
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    if not tree_widget:
        return
    # Search all top-level notebook items
    for i in range(tree_widget.topLevelItemCount()):
        top = tree_widget.topLevelItem(i)
        # Ensure sections are present under this notebook item
        if top.childCount() == 0:
            try:
                from ui_sections import add_sections_as_children
                notebook_id = top.data(0, USER_ROLE_ID)
                if notebook_id is not None:
                    add_sections_as_children(tree_widget, notebook_id, top, window._db_path)
            except Exception:
                pass
        for j in range(top.childCount()):
            child = top.child(j)
            if child.data(0, USER_ROLE_ID) == section_id:
                window._suppress_sync = True
                tree_widget.setCurrentItem(child)
                # Expand only the matching binder so the selection is visible
                try:
                    if not top.isExpanded():
                        top.setExpanded(True)
                except Exception:
                    pass
                window._suppress_sync = False
                return

def _build_right_tree_for_notebook(window, notebook_id):
    """Build the right-side QTreeWidget listing sections and their pages for the active notebook."""
    # QTreeWidget path
    right_tree = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
    if right_tree is not None:
        right_tree.clear()
        sections = get_sections_by_notebook_id(notebook_id, window._db_path)
        color_map = get_section_color_map(notebook_id, window._db_path)
        # Restore expanded sections from settings
        try:
            from settings_manager import get_expanded_sections_by_notebook
            expanded_map = get_expanded_sections_by_notebook()
            expanded_for_nb = expanded_map.get(str(int(notebook_id)), set())
        except Exception:
            expanded_for_nb = set()
        for section in sections:
            section_id = section[0]
            section_title = str(section[2])
            section_color = color_map.get(section_id) if color_map else None
            sec_item = QtWidgets.QTreeWidgetItem([section_title])
            sec_item.setData(0, USER_ROLE_ID, section_id)
            sec_item.setData(0, USER_ROLE_KIND, 'section')
            # Accept drops (pages) and enable selection
            sec_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDropEnabled)
            # Apply colored icon if available
            if section_color:
                sec_item.setIcon(0, _make_color_icon(section_color))
            right_tree.addTopLevelItem(sec_item)

            pages = get_pages_by_section_id(section_id, window._db_path)
            if pages:
                try:
                    pages_sorted = sorted(pages, key=lambda p: (p[6], p[0]))
                except Exception:
                    pages_sorted = pages
                for p in pages_sorted:
                    page_id = p[0]
                    page_title = str(p[2])
                    page_item = QtWidgets.QTreeWidgetItem([page_title])
                    page_item.setData(0, USER_ROLE_ID, page_id)
                    page_item.setData(0, USER_ROLE_KIND, 'page')
                    page_item.setData(0, USER_ROLE_PARENT_SECTION, section_id)
                    # Draggable pages; explicitly not drop targets (prevents sub-pages)
                    page_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled)
                    if section_color:
                        page_item.setIcon(0, _make_color_icon(section_color, is_page=True))
                    sec_item.addChild(page_item)
            else:
                placeholder = QtWidgets.QTreeWidgetItem(["No pages"])
                placeholder.setDisabled(True)
                sec_item.addChild(placeholder)
            # Restore expanded state for this section
            try:
                if int(section_id) in expanded_for_nb:
                    sec_item.setExpanded(True)
            except Exception:
                pass
        # Hook expand/collapse to persist state
        try:
            def _on_sec_expanded(item):
                if item is not None and item.parent() is None and item.data(0, USER_ROLE_KIND) == 'section':
                    from settings_manager import add_expanded_section
                    add_expanded_section(int(notebook_id), int(item.data(0, USER_ROLE_ID)))
            def _on_sec_collapsed(item):
                if item is not None and item.parent() is None and item.data(0, USER_ROLE_KIND) == 'section':
                    from settings_manager import remove_expanded_section
                    remove_expanded_section(int(notebook_id), int(item.data(0, USER_ROLE_ID)))
            # Avoid duplicate connections
            try:
                right_tree.itemExpanded.disconnect()
            except Exception:
                pass
            try:
                right_tree.itemCollapsed.disconnect()
            except Exception:
                pass
            right_tree.itemExpanded.connect(_on_sec_expanded)
            right_tree.itemCollapsed.connect(_on_sec_collapsed)
        except Exception:
            pass
        return

    # QTreeView path (model-based)
    right_view = window.findChild(QtWidgets.QTreeView, 'sectionPages')
    if right_view is not None:
        model = QStandardItemModel(right_view)
        # Hide header text if headerHidden is true; content still OK
        sections = get_sections_by_notebook_id(notebook_id, window._db_path)
        color_map = get_section_color_map(notebook_id, window._db_path)
        for section in sections:
            section_id = section[0]
            section_title = str(section[2])
            section_color = color_map.get(section_id) if color_map else None
            sec_item = QStandardItem(section_title)
            sec_item.setEditable(False)
            sec_item.setData(section_id, USER_ROLE_ID)
            sec_item.setData('section', USER_ROLE_KIND)
            # Accept drops (pages) but do not drag sections
            sec_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDropEnabled)
            if section_color:
                sec_item.setIcon(_make_color_icon(section_color))

            pages = get_pages_by_section_id(section_id, window._db_path)
            if pages:
                try:
                    pages_sorted = sorted(pages, key=lambda p: (p[6], p[0]))
                except Exception:
                    pages_sorted = pages
                for p in pages_sorted:
                    page_id = p[0]
                    page_title = str(p[2])
                    page_item = QStandardItem(page_title)
                    page_item.setEditable(False)
                    page_item.setData(page_id, USER_ROLE_ID)
                    page_item.setData('page', USER_ROLE_KIND)
                    page_item.setData(section_id, USER_ROLE_PARENT_SECTION)
                    # Draggable pages; not drop targets
                    page_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled)
                    if section_color:
                        page_item.setIcon(_make_color_icon(section_color, is_page=True))
                    sec_item.appendRow(page_item)
            else:
                placeholder = QStandardItem("No pages")
                placeholder.setEnabled(False)
                sec_item.appendRow(placeholder)

            model.appendRow(sec_item)
        right_view.setModel(model)
        # Restore expanded sections for model view
        try:
            from settings_manager import get_expanded_sections_by_notebook
            expanded_map = get_expanded_sections_by_notebook()
            expanded_for_nb = expanded_map.get(str(int(notebook_id)), set())
            for row in range(model.rowCount()):
                idx = model.index(row, 0)
                sid = idx.data(USER_ROLE_ID)
                if sid is not None and int(sid) in expanded_for_nb:
                    right_view.expand(idx)
        except Exception:
            pass
        # Hook expand/collapse to persist state in model view
        try:
            def _on_view_expanded(idx: QModelIndex):
                if idx.isValid() and idx.data(USER_ROLE_KIND) == 'section':
                    from settings_manager import add_expanded_section
                    add_expanded_section(int(notebook_id), int(idx.data(USER_ROLE_ID)))
            def _on_view_collapsed(idx: QModelIndex):
                if idx.isValid() and idx.data(USER_ROLE_KIND) == 'section':
                    from settings_manager import remove_expanded_section
                    remove_expanded_section(int(notebook_id), int(idx.data(USER_ROLE_ID)))
            # Avoid duplicate connections
            try:
                right_view.expanded.disconnect()
            except Exception:
                pass
            try:
                right_view.collapsed.disconnect()
            except Exception:
                pass
            right_view.expanded.connect(_on_view_expanded)
            right_view.collapsed.connect(_on_view_collapsed)
        except Exception:
            pass
        return


def _make_color_icon(color_hex: str, is_page: bool = False) -> QIcon:
    """Create a small colored icon. Sections: circle; Pages: smaller circle or square."""
    try:
        color = QColor(color_hex)
        if not color.isValid():
            raise ValueError
    except Exception:
        color = QColor('#888888')
    size = 14 if not is_page else 10
    pm = QPixmap(size, size)
    pm.fill(QtWidgets.QApplication.palette().base().color())
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setBrush(color)
    painter.setPen(color)
    if is_page:
        # draw a small square for page
        painter.drawRect(1, 1, size - 2, size - 2)
    else:
        # draw a circle for section
        painter.drawEllipse(1, 1, size - 2, size - 2)
    painter.end()
    return QIcon(pm)


def _on_tab_bar_context_menu(window, pos):
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if tab_widget is None:
        return
    tab_bar = tab_widget.tabBar()
    if tab_bar is None:
        return
    index = tab_bar.tabAt(pos)
    if index < 0:
        return
    section_id = tab_bar.tabData(index)
    if section_id is None:
        return
    # If user right-clicked the "+ New Section" sentinel, only offer New Section
    if _is_add_section_sentinel(section_id):
        menu = QtWidgets.QMenu(tab_bar)
        act_new_section = menu.addAction("New Section")
        action = menu.exec_(tab_bar.mapToGlobal(pos))
        if action == act_new_section:
            title, ok = QtWidgets.QInputDialog.getText(tab_bar, "New Section", "Section title:", text="Untitled Section")
            if ok:
                title = (title or "").strip() or "Untitled Section"
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    new_id = create_section(nb_id, title, window._db_path)
                    refresh_for_notebook(window, nb_id, select_section_id=new_id)
        return

    menu = QtWidgets.QMenu(tab_bar)
    # Section CRUD
    act_new_section = menu.addAction("New Section")
    act_rename_section = menu.addAction("Rename Section")
    act_delete_section = menu.addAction("Delete Section")
    menu.addSeparator()
    act_set = menu.addAction("Set Color…")
    act_clear = menu.addAction("Clear Color")
    action = menu.exec_(tab_bar.mapToGlobal(pos))
    if action is None:
        return
    if action == act_new_section:
        title, ok = QtWidgets.QInputDialog.getText(tab_bar, "New Section", "Section title:", text="Untitled Section")
        if ok:
            title = title.strip() or "Untitled Section"
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                new_id = create_section(nb_id, title, window._db_path)
                # Rebuild tabs/right pane and select the new section
                refresh_for_notebook(window, nb_id, select_section_id=new_id)
    elif action == act_rename_section:
        current_text = tab_bar.tabText(index)
        new_title, ok = QtWidgets.QInputDialog.getText(tab_bar, "Rename Section", "New title:", text=current_text)
        if ok and new_title.strip():
            rename_section(section_id, new_title.strip(), window._db_path)
            tab_widget.setTabText(index, new_title.strip())
            # Refresh right pane, update left tree children
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                _build_right_tree_for_notebook(window, nb_id)
                _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
    elif action == act_delete_section:
        confirm = QtWidgets.QMessageBox.question(
            tab_bar,
            "Delete Section",
            "Are you sure you want to delete this section and all its pages?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        )
        if confirm == QtWidgets.QMessageBox.Yes:
            # Save edits before deletion
            try:
                save_current_page(window)
            except Exception:
                pass
            nb_id = getattr(window, '_current_notebook_id', None)
            delete_section(section_id, window._db_path)
            if nb_id is not None:
                _populate_tabs_for_notebook(window, nb_id)
                _build_right_tree_for_notebook(window, nb_id)
                _refresh_left_tree_children(window, nb_id)
                # Select first tab if available
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                if tab_widget and tab_widget.count() > 0:
                    tab_widget.setCurrentIndex(0)
                    _load_first_page_for_current_tab(window)
    elif action == act_set:
        color = QtWidgets.QColorDialog.getColor(parent=tab_bar, title="Pick Section Color")
        if color.isValid():
            update_section_color(section_id, color.name(), window._db_path)
            _apply_tab_color(tab_widget, index, color.name())
            # rebuild right panel to reflect new color
            current_nb = getattr(window, '_current_notebook_id', None)
            if current_nb is not None:
                _build_right_tree_for_notebook(window, current_nb)
            _select_right_tree_section(window, section_id)
    elif action == act_clear:
        update_section_color(section_id, None, window._db_path)
        _apply_tab_color(tab_widget, index, None)
        current_nb = getattr(window, '_current_notebook_id', None)
        if current_nb is not None:
            _build_right_tree_for_notebook(window, current_nb)
        _select_right_tree_section(window, section_id)


def _persist_tab_order(window):
    """Persist the current tab order (sections) into the database order_index."""
    try:
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if not tab_widget:
            return
        tab_bar = tab_widget.tabBar()
        if not tab_bar:
            return
        nb_id = getattr(window, '_current_notebook_id', None)
        if nb_id is None:
            return
        # Debounce: avoid overlapping refreshes when tabMoved emits multiple times
        if getattr(window, '_persisting_tab_order', False):
            return
        window._persisting_tab_order = True
        ordered_ids = []
        for i in range(tab_widget.count()):
            sid = tab_bar.tabData(i)
            if sid is not None and not _is_add_section_sentinel(sid):
                try:
                    ordered_ids.append(int(sid))
                except Exception:
                    pass
        if ordered_ids:
            set_sections_order(nb_id, ordered_ids, window._db_path)
            # Refresh trees and keep selection
            current_idx = tab_widget.currentIndex()
            current_sid = tab_bar.tabData(current_idx) if current_idx >= 0 else None
            _build_right_tree_for_notebook(window, nb_id)
            if current_sid is not None:
                _refresh_left_tree_children(window, nb_id, select_section_id=current_sid)
                _select_right_tree_section(window, current_sid)
    except Exception:
        pass
    finally:
        try:
            window._persisting_tab_order = False
        except Exception:
            pass


class _RightTreeDnDFilter(QObject):
    """Event filter to persist page reorder and reparenting in QTreeWidget; disallow nesting under pages."""
    def __init__(self, window):
        super().__init__()
        self._window = window

    def eventFilter(self, obj, event):
        try:
            # Resolve tree and viewport/pos consistently
            tree = None
            viewport = None
            if isinstance(obj, QtWidgets.QTreeWidget):
                tree = obj
                viewport = tree.viewport()
                pos_vp = viewport.mapFrom(tree, event.pos())
            elif isinstance(obj, QtWidgets.QWidget) and isinstance(obj.parent(), QtWidgets.QTreeWidget):
                tree = obj.parent()
                viewport = obj
                pos_vp = event.pos()
            else:
                return False

            if event.type() in (QEvent.DragEnter, QEvent.DragMove):
                # Allow drops when hovering a valid item (section or page); disallow root whitespace
                item = tree.itemAt(pos_vp)
                if item is None:
                    event.ignore()
                    return True
                # Expand section under hover for easier targeting
                if item.data(0, USER_ROLE_KIND) == 'section' and not item.isExpanded():
                    item.setExpanded(True)
                # Auto-scroll near edges
                vp_h = viewport.height() if viewport is not None else tree.viewport().height()
                top_thresh, bot_thresh = 20, vp_h - 20
                vbar = tree.verticalScrollBar()
                if pos_vp.y() < top_thresh:
                    vbar.setValue(max(0, vbar.value() - 10))
                elif pos_vp.y() > bot_thresh:
                    vbar.setValue(vbar.value() + 10)
            if event.type() == QEvent.Drop:
                # Perform controlled move with persistence; block default handling
                target_item = tree.itemAt(pos_vp)
                src_item = tree.currentItem()
                # Validate source is a page and target is a section or a page (use its parent section)
                if src_item is None or src_item.data(0, USER_ROLE_KIND) != 'page':
                    event.ignore()
                    return True
                if target_item is None:
                    event.ignore()
                    return True
                if target_item.data(0, USER_ROLE_KIND) == 'page':
                    # Drop relative to this page into its parent section
                    parent_sec = target_item.parent()
                    if parent_sec is None or parent_sec.data(0, USER_ROLE_KIND) != 'section':
                        event.ignore()
                        return True
                    target_section_item = parent_sec
                    # Determine insert index based on cursor position relative to target page
                    rect = tree.visualItemRect(target_item)  # viewport coords
                    before = pos_vp.y() < rect.center().y()
                    insert_idx = 0
                    # find index of target_item among pages
                    for j in range(target_section_item.childCount()):
                        if target_section_item.child(j) is target_item:
                            insert_idx = j if before else j + 1
                            break
                elif target_item.data(0, USER_ROLE_KIND) == 'section':
                    target_section_item = target_item
                    insert_idx = target_section_item.childCount()  # append at end
                else:
                    event.ignore()
                    return True

                moved_pid = int(src_item.data(0, USER_ROLE_ID))
                src_section_item = src_item.parent()
                src_section_id = src_section_item.data(0, USER_ROLE_ID) if src_section_item else None
                tgt_section_id = int(target_section_item.data(0, USER_ROLE_ID))

                # Build ordered page ids for target section
                target_pages = []
                for j in range(target_section_item.childCount()):
                    ch = target_section_item.child(j)
                    if ch and ch.data(0, USER_ROLE_KIND) == 'page':
                        target_pages.append(int(ch.data(0, USER_ROLE_ID)))
                # Remove moved if already listed (moving within same section)
                if moved_pid in target_pages:
                    target_pages.remove(moved_pid)
                # Clamp insert index
                if insert_idx < 0:
                    insert_idx = 0
                if insert_idx > len(target_pages):
                    insert_idx = len(target_pages)
                target_pages.insert(insert_idx, moved_pid)

                # Persist target section (reparent + reorder)
                set_pages_parent_and_order(tgt_section_id, target_pages, self._window._db_path)

                # If moving across sections, reindex the source section remaining pages
                if src_section_id is not None and src_section_id != tgt_section_id:
                    src_pages = []
                    for j in range(src_section_item.childCount()):
                        ch = src_section_item.child(j)
                        if ch and ch is not src_item and ch.data(0, USER_ROLE_KIND) == 'page':
                            src_pages.append(int(ch.data(0, USER_ROLE_ID)))
                    if src_pages:
                        set_pages_order(src_section_id, src_pages, self._window._db_path)

                # Refresh view and select the moved page in its new section
                nb_id = getattr(self._window, '_current_notebook_id', None)
                if nb_id is not None:
                    _build_right_tree_for_notebook(self._window, nb_id)
                _select_right_tree_page(self._window, tgt_section_id, moved_pid)
                event.accept()
                return True
        except Exception:
            pass
        return False  # continue default handling


def _persist_right_tree_page_orders(window):
    """Persist page order per section from the right QTreeWidget after a drop."""
    try:
        right_tree = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
        if right_tree is None:
            return
        nb_id = getattr(window, '_current_notebook_id', None)
        # Update each section's page order
        for i in range(right_tree.topLevelItemCount()):
            sec_item = right_tree.topLevelItem(i)
            if sec_item is None or sec_item.data(0, USER_ROLE_KIND) != 'section':
                continue
            section_id = sec_item.data(0, USER_ROLE_ID)
            # Build ordered list of page ids
            ordered_page_ids = []
            for j in range(sec_item.childCount()):
                child = sec_item.child(j)
                if child and child.data(0, USER_ROLE_KIND) == 'page':
                    pid = child.data(0, USER_ROLE_ID)
                    if pid is not None:
                        ordered_page_ids.append(int(pid))
            if ordered_page_ids:
                # Move any listed pages into this section (if needed) and set order
                set_pages_parent_and_order(section_id, ordered_page_ids, window._db_path)
        # Rebuild view to snap back any illegal reparenting and keep selection
        if nb_id is not None:
            _build_right_tree_for_notebook(window, nb_id)
            # Keep current tab's section selected in right tree
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            if tab_widget is not None:
                tab_bar = tab_widget.tabBar()
                if tab_bar is not None and tab_widget.currentIndex() >= 0:
                    sid = tab_bar.tabData(tab_widget.currentIndex())
                    if sid is not None:
                        _select_right_tree_section(window, sid)
    except Exception:
        pass


class _RightViewDnDFilter(QObject):
    """Event filter to persist page reorder and reparenting in the QTreeView; disallow nesting under pages."""
    def __init__(self, window):
        super().__init__()
        self._window = window

    def _top_section_index(self, view, idx: QModelIndex) -> QModelIndex:
        if not idx.isValid():
            return QModelIndex()
        while idx.parent().isValid():
            idx = idx.parent()
        return idx

    def eventFilter(self, obj, event):
        try:
            # Resolve view and pos consistently
            if isinstance(obj, QtWidgets.QTreeView):
                view = obj
                viewport = view.viewport()
                pos_vp = viewport.mapFrom(view, event.pos())
            elif isinstance(obj, QtWidgets.QWidget) and isinstance(obj.parent(), QtWidgets.QTreeView):
                view = obj.parent()
                viewport = obj
                pos_vp = event.pos()
            else:
                return False

            if event.type() in (QEvent.DragEnter, QEvent.DragMove):
                idx = view.indexAt(pos_vp)
                # Allow hovering over section or page; disallow invalid/root whitespace
                if not idx.isValid():
                    event.ignore()
                    return True
                # Expand section under hover
                if idx.data(USER_ROLE_KIND) == 'section':
                    try:
                        view.expand(idx)
                    except Exception:
                        pass
                # Auto-scroll near edges
                vp_h = viewport.height()
                top_thresh, bot_thresh = 20, vp_h - 20
                vbar = view.verticalScrollBar()
                if pos_vp.y() < top_thresh:
                    vbar.setValue(max(0, vbar.value() - 10))
                elif pos_vp.y() > bot_thresh:
                    vbar.setValue(vbar.value() + 10)
            if event.type() == QEvent.Drop:
                if not isinstance(view, QtWidgets.QTreeView):
                    return False
                # Allow default processing (including cross-section reparenting), then persist
                QTimer.singleShot(0, lambda: _persist_right_view_page_orders(self._window))
        except Exception:
            pass
        return False


def _persist_right_view_page_orders(window):
    """Persist page order and section reparenting from the right QTreeView after a drop."""
    try:
        right_view = window.findChild(QtWidgets.QTreeView, 'sectionPages')
        if right_view is None or right_view.model() is None:
            return
        model = right_view.model()
        for row in range(model.rowCount()):
            sec_idx = model.index(row, 0)
            if sec_idx.data(USER_ROLE_KIND) != 'section':
                continue
            section_id = sec_idx.data(USER_ROLE_ID)
            ordered_page_ids = []
            for crow in range(model.rowCount(sec_idx)):
                child_idx = model.index(crow, 0, sec_idx)
                if child_idx.data(USER_ROLE_KIND) == 'page':
                    pid = child_idx.data(USER_ROLE_ID)
                    if pid is not None:
                        ordered_page_ids.append(int(pid))
            if ordered_page_ids:
                # Move any pages to this section and set sequential order
                set_pages_parent_and_order(section_id, ordered_page_ids, window._db_path)
        # Rebuild right tree to unify behavior and ensure any visual inconsistencies are resolved
        nb_id = getattr(window, '_current_notebook_id', None)
        if nb_id is not None:
            _build_right_tree_for_notebook(window, nb_id)
            # Reselect the active tab's section in the right pane
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            if tab_widget is not None and tab_widget.count() > 0:
                tab_bar = tab_widget.tabBar()
                if tab_bar is not None:
                    sid = tab_bar.tabData(tab_widget.currentIndex())
                    if sid is not None:
                        _select_right_tree_section(window, sid)
    except Exception:
        pass


def _on_left_tree_context_menu(window, tree_widget: QtWidgets.QTreeWidget, pos):
    item = tree_widget.itemAt(pos)
    if item is None:
        return
    is_notebook = (item.parent() is None)
    menu = QtWidgets.QMenu(tree_widget)
    if is_notebook:
        act_new = menu.addAction("New Section")
        action = menu.exec_(tree_widget.mapToGlobal(pos))
        if action == act_new:
            nb_id = item.data(0, USER_ROLE_ID)
            if nb_id is None:
                return
            title, ok = QtWidgets.QInputDialog.getText(tree_widget, "New Section", "Section title:", text="Untitled Section")
            if ok:
                title = title.strip() or "Untitled Section"
                new_id = create_section(nb_id, title, window._db_path)
                _refresh_left_tree_children(window, nb_id, select_section_id=new_id)
                refresh_for_notebook(window, nb_id, select_section_id=new_id)
    else:
        section_id = item.data(0, USER_ROLE_ID)
        act_rename = menu.addAction("Rename Section")
        act_delete = menu.addAction("Delete Section")
        action = menu.exec_(tree_widget.mapToGlobal(pos))
        if action == act_rename:
            current_text = item.text(0)
            new_title, ok = QtWidgets.QInputDialog.getText(tree_widget, "Rename Section", "New title:", text=current_text)
            if ok and new_title.strip():
                rename_section(section_id, new_title.strip(), window._db_path)
                item.setText(0, new_title.strip())
                # Update tab text
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                idx = _tab_index_for_section(tab_widget, section_id)
                if idx is not None:
                    tab_widget.setTabText(idx, new_title.strip())
                # Update right pane
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    _build_right_tree_for_notebook(window, nb_id)
        elif action == act_delete:
            confirm = QtWidgets.QMessageBox.question(
                tree_widget,
                "Delete Section",
                "Are you sure you want to delete this section and all its pages?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if confirm == QtWidgets.QMessageBox.Yes:
                try:
                    save_current_page(window)
                except Exception:
                    pass
                nb_id = getattr(window, '_current_notebook_id', None)
                delete_section(section_id, window._db_path)
                if nb_id is not None:
                    _refresh_left_tree_children(window, nb_id)
                    _populate_tabs_for_notebook(window, nb_id)
                    _build_right_tree_for_notebook(window, nb_id)
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                    if tab_widget and tab_widget.count() > 0:
                        tab_widget.setCurrentIndex(0)
                        _load_first_page_for_current_tab(window)


def _tab_index_for_section(tab_widget: QtWidgets.QTabWidget, section_id: int):
    if tab_widget is None:
        return None
    tab_bar = tab_widget.tabBar()
    if tab_bar is None:
        return None
    for i in range(tab_widget.count()):
        if tab_bar.tabData(i) == section_id:
            return i
    return None


def _refresh_left_tree_children(window, notebook_id: int, select_section_id: int = None):
    tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
    if not tree_widget:
        return
    # Only rebuild the target notebook's children; keep others intact
    # Find top-level notebook item
    for i in range(tree_widget.topLevelItemCount()):
        top = tree_widget.topLevelItem(i)
        if top.data(0, USER_ROLE_ID) == notebook_id:
            # Clear existing children and rebuild
            try:
                # Remove any existing children to avoid duplicates on repeated refreshes
                top.takeChildren()
                from ui_sections import add_sections_as_children
                add_sections_as_children(tree_widget, notebook_id, top, window._db_path)
            except Exception:
                pass
            # If no sections, add a disabled placeholder so the expander arrow appears
            try:
                if top.childCount() == 0:
                    ph = QtWidgets.QTreeWidgetItem(["No sections"])
                    ph.setDisabled(True)
                    top.addChild(ph)
            except Exception:
                pass
            if select_section_id is not None:
                for j in range(top.childCount()):
                    if top.child(j).data(0, USER_ROLE_ID) == select_section_id:
                        tree_widget.setCurrentItem(top.child(j))
                        break
            try:
                top.setExpanded(True)
                tree_widget.setEnabled(True)
            except Exception:
                pass
            break


def _on_right_tree_context_menu(window, right_tree, pos):
    item = right_tree.itemAt(pos)
    if item is None:
        return
    kind = item.data(0, USER_ROLE_KIND)
    section_id = item.data(0, USER_ROLE_ID) if kind in ('section','page') else None
    menu = QtWidgets.QMenu(right_tree)
    if kind == 'section':
        # Build full section context menu: section ops first, then New Page, then color ops
        act_new_section = menu.addAction("New Section")
        act_rename_section = menu.addAction("Rename Section")
        act_delete_section = menu.addAction("Delete Section")
        act_move_up = menu.addAction("Move Up")
        act_move_down = menu.addAction("Move Down")
        menu.addSeparator()
        act_new_page = menu.addAction("New Page")
        menu.addSeparator()
        act_set = menu.addAction("Set Color…")
        act_clear = menu.addAction("Clear Color")

        action = menu.exec_(right_tree.mapToGlobal(pos))
        if action is None:
            return
        # Handle actions
        if action == act_new_page:
            # Create and open a new page under this section
            pid = create_page(section_id, "Untitled Page", window._db_path)
            current_nb = getattr(window, '_current_notebook_id', None)
            if current_nb is not None:
                _build_right_tree_for_notebook(window, current_nb)
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, section_id)
            window._suppress_sync = False
            _load_page_for_current_tab(window, pid)
            _select_right_tree_page(window, section_id, pid)
            return
        elif action == act_new_section:
            title, ok = QtWidgets.QInputDialog.getText(right_tree, "New Section", "Section title:", text="Untitled Section")
            if ok:
                title = title.strip() or "Untitled Section"
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    new_id = create_section(nb_id, title, window._db_path)
                    refresh_for_notebook(window, nb_id, select_section_id=new_id)
                    _refresh_left_tree_children(window, nb_id, select_section_id=new_id)
        elif action == act_rename_section:
            current_text = item.text(0)
            new_title, ok = QtWidgets.QInputDialog.getText(right_tree, "Rename Section", "New title:", text=current_text)
            if ok and new_title.strip():
                rename_section(section_id, new_title.strip(), window._db_path)
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    _populate_tabs_for_notebook(window, nb_id)
                    _build_right_tree_for_notebook(window, nb_id)
                    _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
                _select_right_tree_section(window, section_id)
        elif action == act_delete_section:
            confirm = QtWidgets.QMessageBox.question(
                right_tree,
                "Delete Section",
                "Are you sure you want to delete this section and all its pages?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if confirm == QtWidgets.QMessageBox.Yes:
                try:
                    save_current_page(window)
                except Exception:
                    pass
                nb_id = getattr(window, '_current_notebook_id', None)
                delete_section(section_id, window._db_path)
                if nb_id is not None:
                    _populate_tabs_for_notebook(window, nb_id)
                    _build_right_tree_for_notebook(window, nb_id)
                    _refresh_left_tree_children(window, nb_id)
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                    if tab_widget and tab_widget.count() > 0:
                        tab_widget.setCurrentIndex(0)
                        _load_first_page_for_current_tab(window)
        elif action == act_move_up:
            try:
                move_section_up(section_id, window._db_path)
            except Exception:
                pass
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                _populate_tabs_for_notebook(window, nb_id)
                _build_right_tree_for_notebook(window, nb_id)
                _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                _select_tab_for_section(tab_widget, section_id)
                _select_right_tree_section(window, section_id)
        elif action == act_move_down:
            try:
                move_section_down(section_id, window._db_path)
            except Exception:
                pass
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                _populate_tabs_for_notebook(window, nb_id)
                _build_right_tree_for_notebook(window, nb_id)
                _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                _select_tab_for_section(tab_widget, section_id)
                _select_right_tree_section(window, section_id)
        elif action in (act_set, act_clear):
            # Determine current tab index for this section to update its icon
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            tab_bar = tab_widget.tabBar() if tab_widget else None
            tab_index = None
            if tab_bar is not None:
                for i in range(tab_widget.count()):
                    if tab_bar.tabData(i) == section_id:
                        tab_index = i
                        break
            if action == act_set:
                color = QtWidgets.QColorDialog.getColor(parent=right_tree, title="Pick Section Color")
                if color.isValid():
                    update_section_color(section_id, color.name(), window._db_path)
                    if tab_index is not None:
                        _apply_tab_color(tab_widget, tab_index, color.name())
                    _build_right_tree_for_notebook(window, getattr(window, '_current_notebook_id', None))
                    _select_right_tree_section(window, section_id)
            elif action == act_clear:
                update_section_color(section_id, None, window._db_path)
                if tab_index is not None:
                    _apply_tab_color(tab_widget, tab_index, None)
                current_nb = getattr(window, '_current_notebook_id', None)
                if current_nb is not None:
                    _build_right_tree_for_notebook(window, current_nb)
                _select_right_tree_section(window, section_id)
    elif kind == 'page':
        page_id = item.data(0, USER_ROLE_ID)
        parent_section_id = item.data(0, USER_ROLE_PARENT_SECTION)
        act_new_page = menu.addAction("New Page")
        menu.addSeparator()
        act_rename_page = menu.addAction("Rename Page")
        act_delete_page = menu.addAction("Delete Page")
        action = menu.exec_(right_tree.mapToGlobal(pos))
        if action is None:
            return
        if action == act_new_page:
            # Create and open a new page under the same section
            pid = create_page(parent_section_id, "Untitled Page", window._db_path)
            current_nb = getattr(window, '_current_notebook_id', None)
            if current_nb is not None:
                _build_right_tree_for_notebook(window, current_nb)
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, parent_section_id)
            window._suppress_sync = False
            _load_page_for_current_tab(window, pid)
            _select_right_tree_page(window, parent_section_id, pid)
        elif action == act_rename_page:
            new_title, ok = QtWidgets.QInputDialog.getText(right_tree, "Rename Page", "New title:", text=item.text(0))
            if ok and new_title.strip():
                update_page_title(page_id, new_title.strip(), window._db_path)
                current_nb = getattr(window, '_current_notebook_id', None)
                if current_nb is not None:
                    _build_right_tree_for_notebook(window, current_nb)
                _select_right_tree_page(window, parent_section_id, page_id)
        elif action == act_delete_page:
            confirm = QtWidgets.QMessageBox.question(
                right_tree,
                "Delete Page",
                "Are you sure you want to delete this page?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if confirm == QtWidgets.QMessageBox.Yes:
                # Save before delete to avoid losing edits
                try:
                    html = None
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                    if tab_widget is not None:
                        tab_bar = tab_widget.tabBar()
                        current_section_id = tab_bar.tabData(tab_widget.currentIndex()) if tab_bar is not None else None
                        is_current = (current_section_id == parent_section_id)
                        if is_current:
                            # Save current page edits if deleting the active one
                            tab = tab_widget.currentWidget()
                            te = tab.findChild(QtWidgets.QTextEdit, 'textEdit') if tab else None
                            if te is not None:
                                html = te.toHtml()
                    # Persist any captured edits to this page
                    if html is not None:
                        update_page_content(page_id, html, window._db_path)
                except Exception:
                    pass
                delete_page(page_id, window._db_path)
                current_nb = getattr(window, '_current_notebook_id', None)
                if current_nb is not None:
                    _build_right_tree_for_notebook(window, current_nb)
                # If we deleted the active page of the active section, load first remaining page
                try:
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                    tab_bar = tab_widget.tabBar() if tab_widget else None
                    active_section_id = tab_bar.tabData(tab_widget.currentIndex()) if tab_bar is not None else None
                    if active_section_id == parent_section_id:
                        _load_first_page_for_current_tab(window)
                except Exception:
                    pass


def _apply_tab_color(tab_widget: QtWidgets.QTabWidget, index: int, color_hex: str):
    """Apply or clear color style on a single tab label."""
    if tab_widget is None or index < 0 or index >= tab_widget.count():
        return
    tab_bar = tab_widget.tabBar()
    label = tab_bar.tabText(index)
    if color_hex:
        # Style just this tab via Qt stylesheet property 'background-color'
        # Note: Per-tab styling with only stylesheets is limited; using a colored icon hint instead.
        tab_bar.setTabIcon(index, _make_color_icon(color_hex))
    else:
        # Clear icon
        tab_bar.setTabIcon(index, QIcon())


def _save_page_for_tab_index(window, index: int):
    """Persist the QTextEdit HTML for the page currently loaded in the given tab index."""
    try:
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if tab_widget is None or index < 0 or index >= tab_widget.count():
            return
        tab_bar = tab_widget.tabBar()
        if tab_bar is None:
            return
        section_id = tab_bar.tabData(index)
        if section_id is None or _is_add_section_sentinel(section_id):
            return
        page_id = getattr(window, "_current_page_by_section", {}).get(section_id)
        if not page_id:
            return
        tab = tab_widget.widget(index)
        text_edit = tab.findChild(QtWidgets.QTextEdit, 'textEdit') if tab else None
        if text_edit is None:
            return
        html = text_edit.toHtml()
        try:
            html = sanitize_html_for_storage(html)
        except Exception:
            pass
        update_page_content(page_id, html, window._db_path)
        _set_modified_indicator_for_tab(tab, False)
    except Exception:
        # Do not crash on autosave
        pass


def save_current_page(window):
    """Save the current tab's page content to the database if available."""
    try:
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        if tab_widget is None:
            return
        _save_page_for_tab_index(window, tab_widget.currentIndex())
    except Exception:
        pass


def _set_modified_indicator_for_tab(tab_widget_page: QtWidgets.QWidget, visible: bool):
    try:
        label = tab_widget_page.findChild(QtWidgets.QLabel, 'modifiedIndicator')
        if label is not None:
            label.setVisible(bool(visible))
    except Exception:
        pass


def _select_right_tree_section(window, section_id):
    # QTreeWidget path
    right_tree = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
    if right_tree is not None:
        for i in range(right_tree.topLevelItemCount()):
            sec_item = right_tree.topLevelItem(i)
            if sec_item.data(0, USER_ROLE_ID) == section_id:
                window._suppress_sync = True
                right_tree.setCurrentItem(sec_item)
                sec_item.setExpanded(True)
                window._suppress_sync = False
                return
    # QTreeView path
    right_view = window.findChild(QtWidgets.QTreeView, 'sectionPages')
    if right_view is not None and right_view.model() is not None:
        model = right_view.model()
        # iterate top-level rows
        for row in range(model.rowCount()):
            idx = model.index(row, 0)
            if idx.data(USER_ROLE_KIND) == 'section' and idx.data(USER_ROLE_ID) == section_id:
                window._suppress_sync = True
                right_view.setCurrentIndex(idx)
                right_view.expand(idx)
                window._suppress_sync = False
                return


def _select_right_tree_page(window, section_id, page_id):
    # QTreeWidget path
    right_tree = window.findChild(QtWidgets.QTreeWidget, 'sectionPages')
    if right_tree is not None:
        for i in range(right_tree.topLevelItemCount()):
            sec_item = right_tree.topLevelItem(i)
            if sec_item.data(0, USER_ROLE_ID) == section_id:
                for j in range(sec_item.childCount()):
                    child = sec_item.child(j)
                    if child.data(0, USER_ROLE_KIND) != 'page':
                        continue
                    if child.data(0, USER_ROLE_ID) == page_id:
                        window._suppress_sync = True
                        right_tree.setCurrentItem(child)
                        sec_item.setExpanded(True)
                        window._suppress_sync = False
                        return
    # QTreeView path
    right_view = window.findChild(QtWidgets.QTreeView, 'sectionPages')
    if right_view is not None and right_view.model() is not None:
        model = right_view.model()
        for row in range(model.rowCount()):
            sec_idx = model.index(row, 0)
            if sec_idx.data(USER_ROLE_KIND) == 'section' and sec_idx.data(USER_ROLE_ID) == section_id:
                # iterate children
                for crow in range(model.rowCount(sec_idx)):
                    child_idx = model.index(crow, 0, sec_idx)
                    if child_idx.data(USER_ROLE_KIND) != 'page':
                        continue
                    if child_idx.data(USER_ROLE_ID) == page_id:
                        window._suppress_sync = True
                        right_view.setCurrentIndex(child_idx)
                        right_view.expand(sec_idx)
                        window._suppress_sync = False
                        return


def _on_right_view_clicked(window, index: QModelIndex):
    if getattr(window, "_suppress_sync", False):
        return
    if not index.isValid():
        return
    # Save edits before switching
    try:
        save_current_page(window)
    except Exception:
        pass
    kind = index.data(USER_ROLE_KIND)
    if kind == 'section':
        section_id = index.data(USER_ROLE_ID)
        if section_id is None:
            return
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        window._suppress_sync = True
        _select_tab_for_section(tab_widget, section_id)
        window._suppress_sync = False
        # Keep the right tree's selection on the section for this user action
        try:
            window._keep_right_tree_section_selected_once = True
        except Exception:
            pass
        _load_first_page_for_current_tab(window)
        # Explicitly reselect the section in the right view to ensure it stays highlighted
        try:
            _select_right_tree_section(window, section_id)
            # Also sync the left panel to the same section
            _select_tree_section(window, section_id)
        except Exception:
            pass
        set_last_state(section_id=section_id)
    elif kind == 'page':
        page_id = index.data(USER_ROLE_ID)
        parent_section_id = index.data(USER_ROLE_PARENT_SECTION)
        if parent_section_id is None or page_id is None:
            return
        tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
        window._suppress_sync = True
        _select_tab_for_section(tab_widget, parent_section_id)
        window._suppress_sync = False
        _load_page_for_current_tab(window, page_id)
        # Sync left panel to the parent section of the selected page
        try:
            _select_tree_section(window, parent_section_id)
        except Exception:
            pass
        set_last_state(section_id=parent_section_id, page_id=page_id)


def _on_right_view_context_menu(window, right_view: QtWidgets.QTreeView, pos):
    index = right_view.indexAt(pos)
    if not index.isValid():
        return
    kind = index.data(USER_ROLE_KIND)
    if kind not in ('section', 'page'):
        return
    menu = QtWidgets.QMenu(right_view)
    if kind == 'section':
        section_id = index.data(USER_ROLE_ID)
        # Build full section context menu: section ops first, then New Page, then color ops
        act_new_section = menu.addAction("New Section")
        act_rename_section = menu.addAction("Rename Section")
        act_delete_section = menu.addAction("Delete Section")
        act_move_up = menu.addAction("Move Up")
        act_move_down = menu.addAction("Move Down")
        menu.addSeparator()
        act_new_page = menu.addAction("New Page")
        menu.addSeparator()
        act_set = menu.addAction("Set Color…")
        act_clear = menu.addAction("Clear Color")

        action = menu.exec_(right_view.mapToGlobal(pos))
        if action is None:
            return
        if action == act_new_page:
            pid = create_page(section_id, "Untitled Page", window._db_path)
            current_nb = getattr(window, '_current_notebook_id', None)
            if current_nb is not None:
                _build_right_tree_for_notebook(window, current_nb)
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, section_id)
            window._suppress_sync = False
            _load_page_for_current_tab(window, pid)
            _select_right_tree_page(window, section_id, pid)
            return
        elif action == act_new_section:
            title, ok = QtWidgets.QInputDialog.getText(right_view, "New Section", "Section title:", text="Untitled Section")
            if ok:
                title = title.strip() or "Untitled Section"
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    new_id = create_section(nb_id, title, window._db_path)
                    refresh_for_notebook(window, nb_id, select_section_id=new_id)
                    _refresh_left_tree_children(window, nb_id, select_section_id=new_id)
        elif action == act_rename_section:
            current_title = index.data()
            new_title, ok = QtWidgets.QInputDialog.getText(right_view, "Rename Section", "New title:", text=str(current_title or ""))
            if ok and new_title.strip():
                rename_section(section_id, new_title.strip(), window._db_path)
                nb_id = getattr(window, '_current_notebook_id', None)
                if nb_id is not None:
                    _populate_tabs_for_notebook(window, nb_id)
                    _build_right_tree_for_notebook(window, nb_id)
                    _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
                _select_right_tree_section(window, section_id)
        elif action == act_delete_section:
            confirm = QtWidgets.QMessageBox.question(
                right_view,
                "Delete Section",
                "Are you sure you want to delete this section and all its pages?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if confirm == QtWidgets.QMessageBox.Yes:
                try:
                    save_current_page(window)
                except Exception:
                    pass
                nb_id = getattr(window, '_current_notebook_id', None)
                delete_section(section_id, window._db_path)
                if nb_id is not None:
                    _populate_tabs_for_notebook(window, nb_id)
                    _build_right_tree_for_notebook(window, nb_id)
                    _refresh_left_tree_children(window, nb_id)
                    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                    if tab_widget and tab_widget.count() > 0:
                        tab_widget.setCurrentIndex(0)
                        _load_first_page_for_current_tab(window)
        elif action == act_move_up:
            try:
                move_section_up(section_id, window._db_path)
            except Exception:
                pass
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                _populate_tabs_for_notebook(window, nb_id)
                _build_right_tree_for_notebook(window, nb_id)
                _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                _select_tab_for_section(tab_widget, section_id)
                _select_right_tree_section(window, section_id)
        elif action == act_move_down:
            try:
                move_section_down(section_id, window._db_path)
            except Exception:
                pass
            nb_id = getattr(window, '_current_notebook_id', None)
            if nb_id is not None:
                _populate_tabs_for_notebook(window, nb_id)
                _build_right_tree_for_notebook(window, nb_id)
                _refresh_left_tree_children(window, nb_id, select_section_id=section_id)
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                _select_tab_for_section(tab_widget, section_id)
                _select_right_tree_section(window, section_id)
        elif action in (act_set, act_clear):
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            tab_bar = tab_widget.tabBar() if tab_widget else None
            tab_index = None
            if tab_bar is not None:
                for i in range(tab_widget.count()):
                    if tab_bar.tabData(i) == section_id:
                        tab_index = i
                        break
            if action == act_set:
                color = QtWidgets.QColorDialog.getColor(parent=right_view, title="Pick Section Color")
                if color.isValid():
                    update_section_color(section_id, color.name(), window._db_path)
                    if tab_index is not None:
                        _apply_tab_color(tab_widget, tab_index, color.name())
                    _build_right_tree_for_notebook(window, getattr(window, '_current_notebook_id', None))
                    _select_right_tree_section(window, section_id)
            elif action == act_clear:
                update_section_color(section_id, None, window._db_path)
                if tab_index is not None:
                    _apply_tab_color(tab_widget, tab_index, None)
                current_nb = getattr(window, '_current_notebook_id', None)
                if current_nb is not None:
                    _build_right_tree_for_notebook(window, current_nb)
                _select_right_tree_section(window, section_id)
    elif kind == 'page':
        page_id = index.data(USER_ROLE_ID)
        section_id = index.data(USER_ROLE_PARENT_SECTION)
        act_new_page = menu.addAction("New Page")
        menu.addSeparator()
        act_rename_page = menu.addAction("Rename Page")
        act_delete_page = menu.addAction("Delete Page")
        action = menu.exec_(right_view.mapToGlobal(pos))
        if action is None:
            return
        if action == act_new_page:
            pid = create_page(section_id, "Untitled Page", window._db_path)
            current_nb = getattr(window, '_current_notebook_id', None)
            if current_nb is not None:
                _build_right_tree_for_notebook(window, current_nb)
            tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
            window._suppress_sync = True
            _select_tab_for_section(tab_widget, section_id)
            window._suppress_sync = False
            _load_page_for_current_tab(window, pid)
            _select_right_tree_page(window, section_id, pid)
        elif action == act_rename_page:
            current_title = index.data()
            new_title, ok = QtWidgets.QInputDialog.getText(right_view, "Rename Page", "New title:", text=str(current_title or ""))
            if ok and new_title.strip():
                update_page_title(page_id, new_title.strip(), window._db_path)
                current_nb = getattr(window, '_current_notebook_id', None)
                if current_nb is not None:
                    _build_right_tree_for_notebook(window, current_nb)
                _select_right_tree_page(window, section_id, page_id)
        elif action == act_delete_page:
            confirm = QtWidgets.QMessageBox.question(
                right_view,
                "Delete Page",
                "Are you sure you want to delete this page?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if confirm == QtWidgets.QMessageBox.Yes:
                try:
                    save_current_page(window)
                except Exception:
                    pass
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                tab_bar = tab_widget.tabBar() if tab_widget else None
                current_section_id = tab_bar.tabData(tab_widget.currentIndex()) if tab_bar is not None else None
                is_current = (current_section_id == section_id)
                delete_page(page_id, window._db_path)
                current_nb = getattr(window, '_current_notebook_id', None)
                if current_nb is not None:
                    _build_right_tree_for_notebook(window, current_nb)
                if is_current:
                    _load_first_page_for_current_tab(window)


def restore_last_position(window):
    """Restore last notebook/section/page if possible."""
    _ensure_attrs(window)
    last = get_last_state()
    if not isinstance(last, dict):
        last = {}
    notebook_id = last.get('last_notebook_id')
    section_id = last.get('last_section_id')
    page_id = last.get('last_page_id')
    # If a section is stored, ensure notebook_id matches that section's notebook
    try:
        if section_id is not None:
            sec_nb = _get_section_notebook_id(window, section_id)
            if sec_nb is not None and sec_nb != notebook_id:
                notebook_id = sec_nb
                set_last_state(notebook_id=notebook_id)
    except Exception:
        pass
    if not notebook_id:
        # Fallback: select the first notebook in the left tree to ensure tabs are shown
        tree_widget = window.findChild(QtWidgets.QTreeWidget, 'notebookName')
        if tree_widget and tree_widget.topLevelItemCount() > 0:
            top_item = tree_widget.topLevelItem(0)
            nb_id = top_item.data(0, USER_ROLE_ID)
            if nb_id is not None:
                set_last_state(notebook_id=nb_id)
                _populate_tabs_for_notebook(window, nb_id)
                _build_right_tree_for_notebook(window, nb_id)
                # Select first tab and reflect selection
                tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
                if tab_widget and tab_widget.count() > 0:
                    # Prefer first real section tab (skip sentinel)
                    first_idx = 0
                    try:
                        tb = tab_widget.tabBar()
                        if tb is not None:
                            for i in range(tab_widget.count()):
                                data = tb.tabData(i)
                                if data is not None and not _is_add_section_sentinel(data):
                                    first_idx = i
                                    break
                    except Exception:
                        pass
                    window._suppress_sync = True
                    tab_widget.setCurrentIndex(first_idx)
                    window._suppress_sync = False
                    _load_first_page_for_current_tab(window)
                # Select in trees
                try:
                    tree_widget.setCurrentItem(top_item)
                    # Also select the first section in the right panel
                    tab_bar = tab_widget.tabBar() if tab_widget else None
                    first_section_id = None
                    if tab_bar and tab_widget and tab_widget.count() > 0:
                        for i in range(tab_widget.count()):
                            data = tab_bar.tabData(i)
                            if data is not None and not _is_add_section_sentinel(data):
                                first_section_id = data
                                break
                    if first_section_id is not None and not _is_add_section_sentinel(first_section_id):
                        _select_right_tree_section(window, first_section_id)
                except Exception:
                    pass
        return
    # Populate tabs for notebook
    _populate_tabs_for_notebook(window, notebook_id)
    try:
        window._current_notebook_id = notebook_id
    except Exception:
        pass
    tab_widget = window.findChild(QtWidgets.QTabWidget, 'tabPages')
    if not tab_widget:
        return
    if tab_widget.count() == 0:
        # Even if there are no tabs yet (no sections), build the right pane so
        # the user can add sections/pages from there.
        _build_right_tree_for_notebook(window, notebook_id)
        return
    # Select section tab if provided
    if section_id is not None:
        window._suppress_sync = True
        _select_tab_for_section(tab_widget, section_id)
        window._suppress_sync = False
        _select_tree_section(window, section_id)
    # Load the last page if available, else first page
    if page_id is not None:
        _load_page_for_current_tab(window, page_id)
    else:
        _load_first_page_for_current_tab(window)
    # Reflect selection in the right tree as well
    active_tab_bar = tab_widget.tabBar()
    if active_tab_bar is not None:
        idx = tab_widget.currentIndex()
        active_section_id = active_tab_bar.tabData(idx)
        if active_section_id is not None:
            _build_right_tree_for_notebook(window, notebook_id)
            # If a page is selected for this section, highlight it; else pick first page
            try:
                current_page_id = getattr(window, '_current_page_by_section', {}).get(active_section_id)
                if current_page_id is None:
                    pages = get_pages_by_section_id(active_section_id, window._db_path)
                    if pages:
                        try:
                            pages_sorted = sorted(pages, key=lambda p: (p[6], p[0]))
                        except Exception:
                            pages_sorted = pages
                        current_page_id = pages_sorted[0][0]
                        window._current_page_by_section[active_section_id] = current_page_id
                if current_page_id is not None:
                    _select_right_tree_page(window, active_section_id, current_page_id)
                else:
                    _select_right_tree_section(window, active_section_id)
            except Exception:
                _select_right_tree_section(window, active_section_id)

def _get_section_notebook_id(window, section_id: int):
    try:
        db_path = getattr(window, '_db_path', 'notes.db')
        con = sqlite3.connect(db_path)
        cur = con.cursor()
        cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
        row = cur.fetchone()
        cur.close()
        con.close()
        return int(row[0]) if row else None
    except Exception:
        return None
