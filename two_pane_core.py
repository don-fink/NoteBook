from PyQt5 import QtWidgets
from PyQt5.QtCore import QEvent, QObject, QTimer, Qt, QUrl

# Local role constants used across the left/right trees
USER_ROLE_ID = 1000
USER_ROLE_KIND = 1001
USER_ROLE_PARENT_SECTION = 1002


def is_two_column_ui(window) -> bool:
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        tabs = window.findChild(QtWidgets.QTabWidget, "tabPages")
        return te is not None and tabs is None
    except Exception:
        return False


def _ensure_state(window):
    if not hasattr(window, "_current_notebook_id"):
        window._current_notebook_id = None
    if not hasattr(window, "_current_section_id"):
        window._current_section_id = None
    if not hasattr(window, "_current_page_by_section"):
        window._current_page_by_section = {}


def _set_page_edit_html(window, html: str):
    te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
    if te is None:
        return
    try:
        te.blockSignals(True)
        if not html:
            te.setHtml("")
        else:
            try:
                from ui_richtext import DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT

                if isinstance(html, str) and "<html" in html.lower():
                    if "<body" in html.lower():
                        html = html.replace(
                            "<body",
                            f'<body style="font-family: {DEFAULT_FONT_FAMILY}; font-size: {int(DEFAULT_FONT_SIZE_PT)}pt"',
                            1,
                        )
            except Exception:
                pass
            te.setHtml(html)
    finally:
        te.blockSignals(False)


def load_page(window, page_id: int = None, html: str = None):
    te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
    if te is None:
        return
    if page_id is None:
        _set_page_edit_html(window, "")
        try:
            te.setReadOnly(True)
        except Exception:
            pass
        try:
            title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
            if title_le is not None:
                title_le.blockSignals(True)
                title_le.setEnabled(False)
                title_le.setText("")
                title_le.blockSignals(False)
        except Exception:
            pass
        try:
            if hasattr(window, "_autosave_timer") and window._autosave_timer.isActive():
                window._autosave_timer.stop()
            window._two_col_dirty = False
            window._autosave_ctx = None
        except Exception:
            pass
        try:
            _ensure_state(window)
            sid = getattr(window, "_current_section_id", None)
            if sid is not None:
                window._current_page_by_section[int(sid)] = None
        except Exception:
            pass
        return

    # Fetch from DB if html not provided
    page_row = None
    try:
        from db_pages import get_page_by_id

        page_row = get_page_by_id(int(page_id), window._db_path)
    except Exception:
        page_row = None
    try:
        if html is None:
            html = page_row[3] if page_row else ""
    except Exception:
        pass
    _set_page_edit_html(window, html or "")
    try:
        te.setReadOnly(False)
    except Exception:
        pass
    try:
        title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
        if title_le is not None:
            title = None
            try:
                title = str(page_row[2]) if page_row else None
            except Exception:
                title = None
            title_le.blockSignals(True)
            title_le.setText(title if title else "Untitled Page")
            title_le.setEnabled(True)
            title_le.blockSignals(False)
    except Exception:
        pass
    try:
        if hasattr(window, "_autosave_timer") and window._autosave_timer.isActive():
            window._autosave_timer.stop()
        window._two_col_dirty = False
        sid = getattr(window, "_current_section_id", None)
        if sid is not None:
            window._autosave_ctx = (int(sid), int(page_id))
    except Exception:
        pass
    try:
        sid = getattr(window, "_current_section_id", None)
        if sid is not None:
            _ensure_state(window)
            window._current_page_by_section[int(sid)] = int(page_id)
    except Exception:
        pass


def save_current_title(window):
    try:
        sid = getattr(window, "_current_section_id", None)
        if sid is None:
            return
        pid = getattr(window, "_current_page_by_section", {}).get(int(sid))
        if not pid:
            return
        title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
        if title_le is None:
            return
        new_title = (title_le.text() or "").strip() or "Untitled Page"
        from db_pages import update_page_title

        update_page_title(int(pid), new_title, window._db_path)
        update_left_tree_page_title(window, int(sid), int(pid), new_title)
        try:
            from settings_manager import set_last_state

            set_last_state(section_id=int(sid), page_id=int(pid))
        except Exception:
            pass
    except Exception:
        pass


def update_left_tree_page_title(window, section_id: int, page_id: int, new_title: str):
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if not tree_widget:
            return
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            for j in range(top.childCount()):
                sec_item = top.child(j)
                if (
                    sec_item
                    and sec_item.data(0, USER_ROLE_KIND) == "section"
                    and int(sec_item.data(0, USER_ROLE_ID)) == int(section_id)
                ):
                    for k in range(sec_item.childCount()):
                        page_item = sec_item.child(k)
                        if (
                            page_item
                            and page_item.data(0, USER_ROLE_KIND) == "page"
                            and int(page_item.data(0, USER_ROLE_ID)) == int(page_id)
                        ):
                            page_item.setText(0, new_title)
                            return
    except Exception:
        pass


def save_current_page(window):
    try:
        sid = getattr(window, "_current_section_id", None)
        if sid is None:
            return
        page_id = getattr(window, "_current_page_by_section", {}).get(int(sid))
        if not page_id:
            return
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is None:
            return
        try:
            if te.isReadOnly():
                return
        except Exception:
            pass
        html = te.toHtml()
        try:
            from ui_richtext import sanitize_html_for_storage

            html = sanitize_html_for_storage(html)
        except Exception:
            pass
        from db_pages import update_page_content

        update_page_content(int(page_id), html, window._db_path)
        try:
            if hasattr(window, "_autosave_timer") and window._autosave_timer.isActive():
                window._autosave_timer.stop()
        except Exception:
            pass
        try:
            window._two_col_dirty = False
            window._autosave_ctx = (int(sid), int(page_id))
        except Exception:
            pass
    except Exception:
        pass


def load_first_page(window):
    try:
        sid = getattr(window, "_current_section_id", None)
        if sid is None:
            nb_id = getattr(window, "_current_notebook_id", None)
            if nb_id is None:
                return
            from db_sections import get_sections_by_notebook_id

            sections = get_sections_by_notebook_id(nb_id, window._db_path)
            if not sections:
                _set_page_edit_html(window, "")
                try:
                    te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                    if te is not None:
                        te.setReadOnly(True)
                except Exception:
                    pass
                return
            sid = sections[0][0]
            window._current_section_id = sid
        from db_pages import get_pages_by_section_id

        pages = get_pages_by_section_id(sid, window._db_path)
        page = None
        if pages:
            try:
                pages_sorted = sorted(pages, key=lambda p: (p[6], p[0]))
            except Exception:
                pages_sorted = pages
            page = pages_sorted[0]
        load_page(window, page_id=(page[0] if page else None), html=(page[3] if page else None))
        try:
            from settings_manager import set_last_state

            if page:
                set_last_state(section_id=int(sid), page_id=int(page[0]))
            else:
                set_last_state(section_id=int(sid), page_id=None)
        except Exception:
            pass
    except Exception:
        pass


def select_section(window, section_id):
    try:
        try:
            window._current_section_id = int(section_id)
        except Exception:
            window._current_section_id = section_id
        _set_page_edit_html(window, "")
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        if te is not None:
            te.setReadOnly(True)
        title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
        if title_le is not None:
            title_le.blockSignals(True)
            title_le.setEnabled(False)
            title_le.setText("")
            title_le.blockSignals(False)
    except Exception:
        pass


def restore_last_position(window):
    """Restore last notebook/section/page in two-pane mode using persisted state.

    Reads last state from settings_manager and attempts to:
      - Select notebook (binder) in left tree
      - Populate its sections/pages
      - Select stored section
      - Load stored page (or first page) into editor
    Safe if any widgets are missing; all operations wrapped in try/except.
    """
    try:
        from settings_manager import get_last_state, set_last_state
    except Exception:
        return
    try:
        last = get_last_state()
        if not isinstance(last, dict):
            last = {}
        notebook_id = last.get("last_notebook_id")
        section_id = last.get("last_section_id")
        page_id = last.get("last_page_id")
        # Align notebook_id with section's notebook if only section stored
        if section_id is not None and not notebook_id:
            try:
                import sqlite3
                db_path = getattr(window, "_db_path", None) or "notes.db"
                con = sqlite3.connect(db_path)
                cur = con.cursor()
                cur.execute("SELECT notebook_id FROM sections WHERE id = ?", (int(section_id),))
                row = cur.fetchone()
                con.close()
                if row and row[0] is not None:
                    notebook_id = int(row[0])
            except Exception:
                pass
        # Fallback: select first binder if none stored
        if notebook_id is None:
            try:
                tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
                if tree_widget and tree_widget.topLevelItemCount() > 0:
                    top_item = tree_widget.topLevelItem(0)
                    nb_id = top_item.data(0, USER_ROLE_ID)
                    if nb_id is not None:
                        notebook_id = int(nb_id)
            except Exception:
                pass
        if notebook_id is None:
            return
        try:
            window._current_notebook_id = int(notebook_id)
            set_last_state(notebook_id=int(notebook_id))
        except Exception:
            pass
        # Build sections/pages and select binder
        try:
            ensure_left_tree_sections(window, int(notebook_id))
            _select_left_binder(window, int(notebook_id))
        except Exception:
            pass
        # Restore section
        if section_id is not None:
            try:
                window._current_section_id = int(section_id)
            except Exception:
                window._current_section_id = section_id
            try:
                ensure_left_tree_sections(window, int(notebook_id), select_section_id=int(section_id))
                _select_tree_section(window, int(section_id))
            except Exception:
                pass
        # Load page or first page
        try:
            if page_id is not None and section_id is not None:
                if not hasattr(window, "_current_page_by_section"):
                    window._current_page_by_section = {}
                try:
                    window._current_page_by_section[int(section_id)] = int(page_id)
                except Exception:
                    window._current_page_by_section[section_id] = page_id
                load_page(window, int(page_id))
                try:
                    from left_tree import select_left_tree_page as _select_left_tree_page
                    _select_left_tree_page(window, int(section_id), int(page_id))
                except Exception:
                    pass
            else:
                load_first_page(window)
        except Exception:
            pass
    except Exception:
        pass


def ensure_left_tree_sections(window, notebook_id: int, select_section_id: int = None):
    """Ensure the left tree shows Sections and Pages under the given binder.

    - Finds the top-level binder item with id == notebook_id
    - Rebuilds its children using ui_sections.add_sections_as_children
    - Expands the binder and optionally selects a Section
    """
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if tree_widget is None:
            return
        # Helper to find the binder item by id
        def _find_binder_item(nid: int):
            for i in range(tree_widget.topLevelItemCount()):
                top = tree_widget.topLevelItem(i)
                try:
                    if int(top.data(0, USER_ROLE_ID)) == int(nid):
                        return top
                except Exception:
                    pass
            return None

        binder_item = _find_binder_item(int(notebook_id))
        if binder_item is None:
            # Repopulate top-level binders first
            try:
                from ui_logic import populate_notebook_names

                populate_notebook_names(window, getattr(window, "_db_path", None) or "notes.db")
            except Exception:
                pass
            binder_item = _find_binder_item(int(notebook_id))
            if binder_item is None:
                return

        # Clear current children and rebuild
        try:
            binder_item.takeChildren()
        except Exception:
            pass
        try:
            from ui_sections import add_sections_as_children

            add_sections_as_children(
                tree_widget,
                int(notebook_id),
                binder_item,
                getattr(window, "_db_path", None) or "notes.db",
            )
        except Exception:
            pass
        # Expand binder
        try:
            binder_item.setExpanded(True)
        except Exception:
            pass
        # Optionally select a section in the rebuilt tree
        if select_section_id is not None:
            try:
                _select_tree_section(window, int(select_section_id))
            except Exception:
                pass
    except Exception:
        pass


def _select_left_binder(window, notebook_id: int):
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if not tree_widget:
            return
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            try:
                if int(top.data(0, USER_ROLE_ID)) == int(notebook_id):
                    tree_widget.setCurrentItem(top)
                    return
            except Exception:
                pass
    except Exception:
        pass


def _select_tree_section(window, section_id: int):
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if not tree_widget:
            return
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            for j in range(top.childCount()):
                sec_item = top.child(j)
                try:
                    if sec_item.data(0, USER_ROLE_KIND) == "section" and int(sec_item.data(0, USER_ROLE_ID)) == int(section_id):
                        try:
                            if not top.isExpanded():
                                top.setExpanded(True)
                        except Exception:
                            pass
                        tree_widget.setCurrentItem(sec_item)
                        return
                except Exception:
                    pass
    except Exception:
        pass


def refresh_for_notebook(window, notebook_id: int, select_section_id: int = None, keep_left_tree_selection: bool = False):
    """Two-pane refresh: rebuild left tree for the binder and optional selection.

    Keeps semantics compatible with callers expecting ui_tabs.refresh_for_notebook.
    """
    try:
        _ensure_state(window)
        try:
            window._suppress_sync = False
        except Exception:
            pass
        try:
            window._current_notebook_id = int(notebook_id)
        except Exception:
            window._current_notebook_id = notebook_id
        try:
            from settings_manager import set_last_state

            set_last_state(notebook_id=int(notebook_id))
        except Exception:
            pass
        ensure_left_tree_sections(window, int(notebook_id))
        if select_section_id is not None and not keep_left_tree_selection:
            try:
                window._current_section_id = int(select_section_id)
            except Exception:
                window._current_section_id = select_section_id
            _select_tree_section(window, int(select_section_id))
    except Exception:
        pass


def setup_two_pane(window):
    _ensure_state(window)
    # Install rich text toolbar and autosave wires
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
        title_le_found = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
        if te is not None and not hasattr(window, "_two_col_toolbar_added"):
            container = te.parentWidget() or window
            before_w = title_le_found if title_le_found is not None else te
            from ui_richtext import add_rich_text_toolbar, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT
            from PyQt5.QtGui import QFont

            add_rich_text_toolbar(container, te, before_widget=before_w)
            window._two_col_toolbar_added = True
            te.document().setDefaultFont(QFont(DEFAULT_FONT_FAMILY, int(DEFAULT_FONT_SIZE_PT)))
            try:
                te.setReadOnly(True)
            except Exception:
                pass
            if not hasattr(window, "_autosave_timer"):
                window._autosave_timer = QTimer(window)
                window._autosave_timer.setSingleShot(True)
                window._autosave_timer.setInterval(1200)

                def _autosave_fire():
                    try:
                        ctx = getattr(window, "_autosave_ctx", None)
                        sid_now = getattr(window, "_current_section_id", None)
                        pid_now = (
                            getattr(window, "_current_page_by_section", {}).get(int(sid_now))
                            if sid_now is not None
                            else None
                        )
                        if isinstance(ctx, tuple) and len(ctx) == 2 and ctx[0] == sid_now and ctx[1] == pid_now:
                            save_current_page(window)
                    except Exception:
                        pass

                window._autosave_timer.timeout.connect(_autosave_fire)

            def _on_text_changed():
                try:
                    sid = getattr(window, "_current_section_id", None)
                    pid = (
                        getattr(window, "_current_page_by_section", {}).get(int(sid))
                        if sid is not None
                        else None
                    )
                    if pid is not None:
                        try:
                            window._two_col_dirty = True
                        except Exception:
                            pass
                        try:
                            window._autosave_ctx = (int(sid), int(pid))
                        except Exception:
                            window._autosave_ctx = (sid, pid)
                        window._autosave_timer.start()
                except Exception:
                    pass

            te.textChanged.connect(_on_text_changed)

            class _FocusSaveFilter(QObject):
                def eventFilter(self, obj, event):
                    try:
                        if event.type() == QEvent.FocusOut:
                            save_current_page(window)
                    except Exception:
                        pass
                    return False

            window._page_edit_focus_filter = _FocusSaveFilter(te)
            te.installEventFilter(window._page_edit_focus_filter)

        try:
            title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
            if title_le is not None:
                try:
                    title_le.setEnabled(False)
                except Exception:
                    pass
                if not hasattr(window, "_two_col_title_wired"):
                    if not hasattr(window, "_title_save_timer"):
                        window._title_save_timer = QTimer(window)
                        window._title_save_timer.setSingleShot(True)
                        window._title_save_timer.setInterval(600)

                        def _title_autosave_fire():
                            try:
                                save_current_title(window)
                            except Exception:
                                pass

                        window._title_save_timer.timeout.connect(_title_autosave_fire)

                    def _on_title_changed(_text: str):
                        try:
                            sid = getattr(window, "_current_section_id", None)
                            pid = (
                                getattr(window, "_current_page_by_section", {}).get(int(sid))
                                if sid is not None
                                else None
                            )
                            if pid is not None:
                                window._title_save_timer.start()
                        except Exception:
                            pass

                    def _on_title_commit():
                        try:
                            save_current_title(window)
                        except Exception:
                            pass

                    title_le.textChanged.connect(_on_title_changed)
                    try:
                        title_le.editingFinished.connect(_on_title_commit)
                    except Exception:
                        pass
                    try:
                        title_le.returnPressed.connect(_on_title_commit)
                    except Exception:
                        pass
                    window._two_col_title_wired = True
        except Exception:
            pass
    except Exception:
        pass

    # Left tree interactions for two-pane
    tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
    if tree_widget is not None and not getattr(tree_widget, "_nb_left_signals_connected", False):

        def on_tree_item_clicked(item, column):
            if getattr(window, "_suppress_sync", False):
                return
            try:
                save_current_page(window)
            except Exception:
                pass
            kind = item.data(0, USER_ROLE_KIND)
            if item.parent() is None and kind not in ("section", "page"):
                nb_id = item.data(0, USER_ROLE_ID)
                if nb_id is None:
                    return
                window._current_notebook_id = int(nb_id)
                try:
                    from settings_manager import set_last_state

                    set_last_state(notebook_id=int(nb_id))
                except Exception:
                    pass
                ensure_left_tree_sections(window, int(nb_id))
                try:
                    _set_page_edit_html(window, "")
                    te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                    if te is not None:
                        te.setReadOnly(True)
                    title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
                    if title_le is not None:
                        title_le.blockSignals(True)
                        title_le.setEnabled(False)
                        title_le.setText("")
                        title_le.blockSignals(False)
                    if hasattr(window, "_autosave_timer") and window._autosave_timer.isActive():
                        window._autosave_timer.stop()
                        window._two_col_dirty = False
                except Exception:
                    pass
            elif kind == "section":
                sid = item.data(0, USER_ROLE_ID)
                if sid is None:
                    return
                window._current_section_id = int(sid)
                try:
                    if not item.isExpanded():
                        item.setExpanded(True)
                except Exception:
                    pass
                try:
                    _set_page_edit_html(window, "")
                    te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                    if te is not None:
                        te.setReadOnly(True)
                    title_le = window.findChild(QtWidgets.QLineEdit, "pageTitleEdit")
                    if title_le is not None:
                        title_le.blockSignals(True)
                        title_le.setEnabled(False)
                        title_le.setText("")
                        title_le.blockSignals(False)
                    if hasattr(window, "_autosave_timer") and window._autosave_timer.isActive():
                        window._autosave_timer.stop()
                        window._two_col_dirty = False
                except Exception:
                    pass
            elif kind == "page":
                pid = item.data(0, USER_ROLE_ID)
                parent_sid = item.data(0, USER_ROLE_PARENT_SECTION)
                if pid is None or parent_sid is None:
                    return
                try:
                    window._current_section_id = int(parent_sid)
                    _ensure_state(window)
                    window._current_page_by_section[int(parent_sid)] = int(pid)
                except Exception:
                    pass
                load_page(window, int(pid))
                try:
                    from settings_manager import set_last_state

                    set_last_state(section_id=int(parent_sid), page_id=int(pid))
                except Exception:
                    pass

        try:
            tree_widget.itemClicked.disconnect()
        except Exception:
            pass
        tree_widget.itemClicked.connect(on_tree_item_clicked)
        tree_widget._nb_left_signals_connected = True

    # Ctrl+S wired to save in two-pane as well (keep name parity)
    try:
        from PyQt5.QtGui import QKeySequence

        QtWidgets.QShortcut(QKeySequence.Save, window, activated=lambda: save_current_page(window))
    except Exception:
        pass

    try:
        def _on_item_expanded(item):
            if item is not None and item.parent() is None:
                nb_id = item.data(0, USER_ROLE_ID)
                if nb_id is not None:
                    ensure_left_tree_sections(window, int(nb_id))

        tw = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if tw is not None:
            try:
                tw.itemExpanded.disconnect()
            except Exception:
                pass
            tw.itemExpanded.connect(_on_item_expanded)
    except Exception:
        pass


def select_left_tree_page(window, section_id: int, page_id: int):
    """Select a page under the given section in the left binder tree.

    Expands the binder and section as needed so the page is visible.
    """
    try:
        tree_widget = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        if not tree_widget:
            return
        for i in range(tree_widget.topLevelItemCount()):
            top = tree_widget.topLevelItem(i)
            for j in range(top.childCount()):
                sec_item = top.child(j)
                if (
                    sec_item
                    and sec_item.data(0, USER_ROLE_KIND) == "section"
                    and int(sec_item.data(0, USER_ROLE_ID)) == int(section_id)
                ):
                    try:
                        if not top.isExpanded():
                            top.setExpanded(True)
                    except Exception:
                        pass
                    try:
                        if not sec_item.isExpanded():
                            sec_item.setExpanded(True)
                    except Exception:
                        pass
                    for k in range(sec_item.childCount()):
                        page_item = sec_item.child(k)
                        if (
                            page_item
                            and page_item.data(0, USER_ROLE_KIND) == "page"
                            and int(page_item.data(0, USER_ROLE_ID)) == int(page_id)
                        ):
                            tree_widget.setCurrentItem(page_item)
                            return
    except Exception:
        pass
