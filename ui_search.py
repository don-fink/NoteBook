"""Search box in the menu bar: find pages and step through them one at a time.

The search covers the binder selected in the left panel (or the binder of the selected section
or page), or the whole database, chosen from the magnifier button's menu. The scope shows next
to the counter, with a Search everywhere button that widens a binder search to the whole
database (Enter does the same when the binder has no matches).
Each result is opened, selected in the left panel and its matching words are highlighted.

Keys: Ctrl+F goes to the box. Enter / Shift+Enter in the box, or F3 / Shift+F3 anywhere, go
to the next / previous result. Esc clears the search.

Themes can style QWidget#nbSearchBar, QLineEdit#nbSearchBox, QLabel#nbSearchCount,
QLabel#nbSearchScopeLabel, QToolButton#nbSearchEverywhere and the highlight color: QWidget#nbSearchBar { qproperty-highlightColor: #ffd54f; }
"""

import search_engine
from qt_compat import QtCore, QtGui, QtWidgets

Qt = QtCore.Qt
USER_ROLE_ID = 1000
USER_ROLE_KIND = 1001


def _magnifier_icon(color: QtGui.QColor, size: int = 16) -> QtGui.QIcon:
    pm = QtGui.QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QtGui.QPainter(pm)
    try:
        p.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        pen = QtGui.QPen(color)
        pen.setWidthF(1.8)
        p.setPen(pen)
        r = size * 0.36
        c = QtCore.QPointF(size * 0.42, size * 0.42)
        p.drawEllipse(c, r, r)
        p.drawLine(
            QtCore.QPointF(size * 0.68, size * 0.68), QtCore.QPointF(size * 0.92, size * 0.92)
        )
    finally:
        p.end()
    return QtGui.QIcon(pm)


class SearchBar(QtWidgets.QWidget):
    """The search box, its scope menu, the result counter and previous/next buttons."""

    def __init__(self, window):
        super().__init__(window)
        self.setObjectName("nbSearchBar")
        self._window = window
        self._highlight = QtGui.QColor("#ffd54f")
        self._results = []
        self._index = -1
        self._scope = None  # (kind, id, name) frozen while a search is active
        self._whole_db = False
        self._navigating = False

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(4, 1, 6, 1)
        layout.setSpacing(3)

        self.scope_btn = QtWidgets.QToolButton(self)
        self.scope_btn.setObjectName("nbSearchScope")
        self.scope_btn.setToolTip("Search in: current selection or whole database")
        self.scope_btn.setPopupMode(QtWidgets.QToolButton.ToolButtonPopupMode.InstantPopup)
        self.scope_btn.setAutoRaise(True)
        menu = QtWidgets.QMenu(self.scope_btn)
        group = QtGui.QActionGroup(menu)
        self.act_selection = menu.addAction("Search in selected binder")
        self.act_whole = menu.addAction("Search whole database")
        for act in (self.act_selection, self.act_whole):
            act.setCheckable(True)
            group.addAction(act)
        # The scope choice is remembered in settings
        try:
            from settings_manager import get_search_whole_database

            self._whole_db = bool(get_search_whole_database())
        except Exception:
            self._whole_db = False
        (self.act_whole if self._whole_db else self.act_selection).setChecked(True)
        menu.addSeparator()
        self.act_exact = menu.addAction("Exact matches only")
        self.act_exact.setCheckable(True)
        self.scope_btn.setMenu(menu)
        layout.addWidget(self.scope_btn)

        self.box = QtWidgets.QLineEdit(self)
        self.box.setObjectName("nbSearchBox")
        self.box.setClearButtonEnabled(True)
        self.box.setMinimumWidth(220)
        self.box.installEventFilter(self)
        layout.addWidget(self.box)

        self.scope_label = QtWidgets.QLabel(self)
        self.scope_label.setObjectName("nbSearchScopeLabel")
        self.scope_label.setMaximumWidth(220)
        layout.addWidget(self.scope_label)

        self.everywhere_btn = QtWidgets.QToolButton(self)
        self.everywhere_btn.setObjectName("nbSearchEverywhere")
        self.everywhere_btn.setText("Search everywhere")
        self.everywhere_btn.setToolTip(
            "Search the whole database for this (Enter when nothing is found)"
        )
        self.everywhere_btn.hide()
        layout.addWidget(self.everywhere_btn)

        self.count = QtWidgets.QLabel(self)
        self.count.setObjectName("nbSearchCount")
        self.count.setMinimumWidth(56)
        self.count.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.count)

        self.prev_btn = QtWidgets.QToolButton(self)
        self.prev_btn.setText("▲")
        self.prev_btn.setToolTip("Previous result (Shift+F3)")
        self.next_btn = QtWidgets.QToolButton(self)
        self.next_btn.setText("▼")
        self.next_btn.setToolTip("Next result (F3)")
        for b in (self.prev_btn, self.next_btn):
            b.setAutoRaise(True)
            layout.addWidget(b)

        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(350)
        self._timer.timeout.connect(self.run)
        self.box.textChanged.connect(self._on_text_changed)
        self.prev_btn.clicked.connect(self.previous)
        self.next_btn.clicked.connect(self.next)
        self.everywhere_btn.clicked.connect(self.search_everywhere)
        self.act_whole.toggled.connect(self._on_scope_mode)
        self.act_exact.toggled.connect(lambda _on: self._restart())
        self._update_controls()

    # Themes set the highlight color with qproperty-highlightColor
    def _get_highlight(self):
        return self._highlight

    def _set_highlight(self, color):
        self._highlight = QtGui.QColor(color)

    highlightColor = QtCore.pyqtProperty(QtGui.QColor, _get_highlight, _set_highlight)

    # --- Scope ---

    def _tree(self):
        return self._window.findChild(QtWidgets.QTreeWidget, "notebookName")

    def _selection_scope(self):
        """(kind, id, name) for the binder of what's selected in the left panel.

        A selected section or page searches its whole binder; nothing selected searches the
        whole database."""
        tree = self._tree()
        item = tree.currentItem() if tree is not None else None
        while item is not None and item.parent() is not None:
            item = item.parent()
        if item is None:
            return ("database", None, "the whole database")
        try:
            return ("binder", int(item.data(0, USER_ROLE_ID)), item.text(0) or "Untitled")
        except Exception:
            return ("database", None, "the whole database")

    def current_scope(self):
        if self._whole_db:
            return ("database", None, "the whole database")
        return self._selection_scope()

    def refresh_placeholder(self):
        """Show the scope in the empty box. While a search is active the scope stays fixed."""
        if self.box.text().strip():
            return
        name = self.current_scope()[2]
        self.box.setPlaceholderText(f"Search in {name}")
        self.box.setToolTip(f"Search in {name} (Ctrl+F)")

    def _on_scope_mode(self, whole: bool):
        self._whole_db = bool(whole)
        try:
            from settings_manager import set_search_whole_database

            set_search_whole_database(self._whole_db)
        except Exception:
            pass
        self.refresh_placeholder()
        self._restart()

    # --- Searching ---

    def _on_text_changed(self, text):
        if not text.strip():
            self.clear_results()
            self.refresh_placeholder()
            return
        if self._scope is None:
            self._scope = self.current_scope()
        self._timer.start()

    def _restart(self):
        if self.box.text().strip():
            self._scope = self.current_scope()
            self.run()

    def run(self, keep_page_id=None):
        query = self.box.text().strip()
        if not query:
            self.clear_results()
            return
        db_path = getattr(self._window, "_db_path", None)
        if not db_path:
            return
        try:
            from two_pane_core import save_current_page

            save_current_page(self._window)  # Search what's on screen, not an older save
        except Exception:
            pass
        try:
            from settings_manager import get_show_deleted

            include_deleted = bool(get_show_deleted())
        except Exception:
            include_deleted = False
        scope = self._scope or self.current_scope()
        try:
            pages = search_engine.load_pages(db_path, scope[:2], include_deleted)
            self._results = search_engine.search(pages, query, self.act_exact.isChecked())
        except Exception:
            self._results = []
        self._index = -1
        if self._results:
            # Stay on the page being viewed if it's still a result (e.g. after widening)
            start = next((i for i, m in enumerate(self._results) if m.page_id == keep_page_id), 0)
            self.go(start)
        else:
            self._clear_highlights()
            self._update_controls()

    def search_everywhere(self):
        """Widen the current search to the whole database (the scope menu is left as it is)."""
        if self.box.text().strip():
            current = self._results[self._index] if 0 <= self._index < len(self._results) else None
            self._scope = ("database", None, "the whole database")
            self.run(keep_page_id=current.page_id if current else None)

    def clear_results(self):
        self._timer.stop()
        self._results = []
        self._index = -1
        self._scope = None
        self._clear_highlights()
        self._update_controls()

    def clear(self):
        """Empty the box and forget the results (Esc, or another database opened)."""
        self.box.blockSignals(True)
        self.box.clear()
        self.box.blockSignals(False)
        self.clear_results()
        self.refresh_placeholder()

    def _update_controls(self):
        n = len(self._results)
        has_query = bool(self.box.text().strip())
        scope = self._scope or self.current_scope()
        narrow = scope[0] != "database"
        if has_query:
            name = scope[2]
            elided = self.scope_label.fontMetrics().elidedText(
                name, Qt.TextElideMode.ElideRight, 200
            )
            self.scope_label.setText(f"in {elided}")
            self.scope_label.setToolTip(f"Searching in {name}")
        else:
            self.scope_label.setText("")
        show_everywhere = bool(has_query and narrow)
        relayout = (
            self.scope_label.isVisibleTo(self) != has_query
            or self.everywhere_btn.isVisibleTo(self) != show_everywhere
        )
        self.scope_label.setVisible(has_query)
        self.everywhere_btn.setVisible(show_everywhere)
        if relayout:
            self._relayout()
        if not has_query:
            self.count.setText("")
        elif n == 0:
            self.count.setText("No results")
        else:
            self.count.setText(f"{self._index + 1} of {n}" if self._index >= 0 else f"{n} found")
        self.prev_btn.setEnabled(n > 1)
        self.next_btn.setEnabled(n > 1)
        self.box.setProperty("noResults", bool(has_query and n == 0))
        self.box.style().unpolish(self.box)
        self.box.style().polish(self.box)

    def _relayout(self):
        """Make the menu bar resize its corner area to fit the bar's current contents.

        QMenuBar only recalculates the corner widget's space when it is resized."""
        try:
            self.updateGeometry()
            mb = self._window.menuBar()
            QtWidgets.QApplication.sendEvent(mb, QtGui.QResizeEvent(mb.size(), mb.size()))
        except Exception:
            pass

    # --- Stepping through results ---

    def next(self):
        if not self._results:
            if self.everywhere_btn.isVisible():
                self.search_everywhere()
            else:
                self.run()
            return
        self.go((self._index + 1) % len(self._results))

    def previous(self):
        if not self._results:
            self.run()
            return
        self.go((self._index - 1) % len(self._results))

    def go(self, index: int):
        if not (0 <= index < len(self._results)):
            return
        self._index = index
        match = self._results[index]
        self._navigating = True
        try:
            self._open_page(match)
        finally:
            self._navigating = False
        self._highlight_matches(match)
        self._update_controls()

    def _open_page(self, match):
        """Select the page in the left panel and open it, as a click on it would."""
        from two_pane_core import ensure_left_tree_sections, select_left_tree_page

        tree = self._tree()
        if tree is None:
            return
        select_left_tree_page(self._window, match.section_id, match.page_id)
        item = tree.currentItem()
        if not self._is_page_item(item, match.page_id):
            # The binder's sections may not have been built yet
            ensure_left_tree_sections(self._window, match.notebook_id)
            select_left_tree_page(self._window, match.section_id, match.page_id)
            item = tree.currentItem()
        if self._is_page_item(item, match.page_id):
            tree.scrollToItem(item)
            tree.itemClicked.emit(item, 0)

    @staticmethod
    def _is_page_item(item, page_id) -> bool:
        try:
            return (
                item is not None
                and item.data(0, USER_ROLE_KIND) == "page"
                and int(item.data(0, USER_ROLE_ID)) == int(page_id)
            )
        except Exception:
            return False

    # --- Highlighting on the page ---

    def _editor(self):
        return self._window.findChild(QtWidgets.QTextEdit, "pageEdit")

    def _clear_highlights(self):
        te = self._editor()
        if te is None:
            return
        te._search_selections = []
        try:
            te.setExtraSelections(list(getattr(te, "_spell_selections", [])))
        except Exception:
            pass

    def _highlight_matches(self, match):
        """Highlight the matched words on the open page and scroll to the first one.

        Uses the editor's extra selections, which are only drawn, never saved."""
        te = self._editor()
        if te is None:
            return
        doc = te.document()
        needles = sorted(
            {search_engine.fold(h) for h in match.highlight if h}, key=len, reverse=True
        )
        fmt = QtGui.QTextCharFormat()
        fmt.setBackground(self._highlight)
        fmt.setForeground(QtGui.QColor("#000000"))
        selections, first = [], None
        block = doc.begin()
        while block.isValid():
            text = block.text()
            # Map each folded letter back to its document position. Qt counts UTF-16 units,
            # so a character outside the BMP (such as an emoji) takes two positions.
            folded, starts, ends, pos = [], [], [], 0
            for ch in text:
                width = 2 if ord(ch) > 0xFFFF else 1
                f = search_engine.fold(ch)
                folded.append(f)
                starts.extend([pos] * len(f))
                ends.extend([pos + width] * len(f))
                pos += width
            folded = "".join(folded)
            for needle in needles:
                start = folded.find(needle)
                while start != -1:
                    a = block.position() + starts[start]
                    b = block.position() + ends[start + len(needle) - 1]
                    cur = QtGui.QTextCursor(doc)
                    cur.setPosition(a)
                    cur.setPosition(b, QtGui.QTextCursor.MoveMode.KeepAnchor)
                    sel = QtWidgets.QTextEdit.ExtraSelection()
                    sel.cursor = cur
                    sel.format = fmt
                    selections.append(sel)
                    if first is None or a < first:
                        first = a
                    start = folded.find(needle, start + len(needle))
            block = block.next()
        te._search_selections = selections
        try:
            te.setExtraSelections(list(getattr(te, "_spell_selections", [])) + selections)
        except Exception:
            pass
        if first is not None:
            cur = QtGui.QTextCursor(doc)
            cur.setPosition(first)
            te.setTextCursor(cur)
            te.ensureCursorVisible()

    # --- Keys in the box ---

    def eventFilter(self, obj, event):
        try:
            if obj is self.box and event.type() == QtCore.QEvent.Type.KeyPress:
                key = event.key()
                if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                    if self._timer.isActive():
                        self._timer.stop()
                        self.run()
                    elif event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                        self.previous()
                    else:
                        self.next()
                    return True
                if key == Qt.Key.Key_Escape:
                    self.clear()
                    te = self._editor()
                    if te is not None:
                        te.setFocus()
                    return True
        except Exception:
            return False
        return False

    def focus_box(self):
        self.refresh_placeholder()
        self.box.setFocus(Qt.FocusReason.ShortcutFocusReason)
        self.box.selectAll()


def install_search(window):
    """Put the search box in the menu bar's top-right corner and wire its menu items and keys."""
    if getattr(window, "_search", None) is not None:
        return window._search
    bar = SearchBar(window)
    window.menuBar().setCornerWidget(bar, Qt.Corner.TopRightCorner)
    window._search = bar

    tree = bar._tree()
    if tree is not None:
        tree.currentItemChanged.connect(lambda *_: bar.refresh_placeholder())

    def _bind(name, handler, fallback_keys):
        act = window.findChild(QtGui.QAction, name)
        if act is not None:
            act.triggered.connect(lambda _checked=False: handler())
        else:
            for keys in fallback_keys:
                QtGui.QShortcut(QtGui.QKeySequence(keys), window, activated=handler)

    _bind("actionFind", bar.focus_box, ["Ctrl+F"])
    _bind("actionFind_Next", bar.next, ["F3"])
    _bind("actionFind_Previous", bar.previous, ["Shift+F3"])
    bar.refresh_placeholder()

    # Draw the magnifier in the theme's text color, now and after theme changes
    def _paint_icon():
        try:
            bar.scope_btn.ensurePolished()
            color = bar.scope_btn.palette().color(QtGui.QPalette.ColorRole.ButtonText)
            bar.scope_btn.setIcon(_magnifier_icon(color))
        except Exception:
            pass

    QtCore.QTimer.singleShot(0, _paint_icon)
    try:
        from theme_manager import on_theme_applied

        on_theme_applied(_paint_icon)
    except Exception:
        pass
    return bar
