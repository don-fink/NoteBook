"""
left_tree.py

Stable surface for left‑panel (Binder/Sections/Pages) helpers while we migrate
implementation away from ui_tabs. Also hosts the LeftTreeDnDFilter used to
reorder top-level binders via drag-and-drop.
"""

from PyQt5 import QtWidgets
from PyQt5.QtCore import QEvent, QObject, QTimer
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QColor, QBrush

# Temporary re-exports while we migrate away from ui_tabs — now two-pane native
from two_pane_core import ensure_left_tree_sections, refresh_for_notebook  # noqa: F401

# Convenience re-exports for selection helpers (two-pane variants)
from two_pane_core import (  # noqa: F401
    _select_left_binder as select_left_binder,
    _select_tree_section as select_tree_section,
)

# Additional helpers used by callers (prefer core where available)
from two_pane_core import select_left_tree_page, update_left_tree_page_title  # noqa: F401


class LeftTreeDnDFilter(QObject):
    """Constrain drag/drop to top-level binder reordering and persist order.

    Installed on the left binder tree and its viewport. After a successful
    drop, it persists order to the DB, repopulates the tree, and restores the
    previous binder selection without auto-expanding other binders.
    """

    def __init__(self, window):
        super().__init__()
        self._window = window
        self._highlight_item = None
        self._highlight_kind = None  # 'child'
        self._indicator_line = None  # overlay line for before/after position

    def _top_level_item_at(self, tree: QtWidgets.QTreeWidget, pos_vp):
        item = tree.itemAt(pos_vp)
        if item is None:
            return None
        return item if item.parent() is None else item.parent()

    def eventFilter(self, obj, event):
        try:
            # Resolve tree and viewport position consistently
            if isinstance(obj, QtWidgets.QTreeWidget):
                tree = obj
                pos_vp = event.pos() if hasattr(event, "pos") else None
                if pos_vp is not None:
                    pos_vp = tree.viewport().mapFrom(tree, pos_vp)
            elif isinstance(obj, QtWidgets.QWidget) and isinstance(obj.parent(), QtWidgets.QTreeWidget):
                tree = obj.parent()
                pos_vp = event.pos() if hasattr(event, "pos") else None
            else:
                return False

            if event.type() in (QEvent.DragEnter, QEvent.DragMove):
                # Determine the dragged item
                sel = tree.selectedItems()
                drag_item = sel[0] if sel else tree.currentItem()
                if drag_item is None:
                    self._clear_child_highlight(); self._hide_indicator_line(tree)
                    event.ignore();
                    return True
                kind = drag_item.data(0, 1001)
                target_item = tree.itemAt(pos_vp) if pos_vp is not None else None
                # Page: allow only sibling moves (same parent)
                if kind == "page":
                    if target_item is None or target_item.data(0, 1001) != "page":
                        self._clear_child_highlight(); self._hide_indicator_line(tree); event.ignore(); return True
                    if target_item.parent() is not drag_item.parent():
                        self._clear_child_highlight(); self._hide_indicator_line(tree); event.ignore(); return True
                    # Show custom before/after indicator line
                    self._clear_child_highlight()
                    rect = tree.visualItemRect(target_item)
                    h = rect.height() or 1
                    rel_y = pos_vp.y() - rect.top()
                    position = "before" if rel_y < (h * 0.5) else "after"
                    self._show_indicator_line(tree, rect, position)
                    return False
                # Section: restrict to same binder (parent top-level item)
                if kind == "section":
                    if target_item is None or target_item.data(0, 1001) != "section":
                        self._clear_child_highlight(); self._hide_indicator_line(tree); event.ignore(); return True
                    if target_item.parent() is None or drag_item.parent() is None:
                        self._clear_child_highlight(); self._hide_indicator_line(tree); event.ignore(); return True
                    if target_item.parent() is not drag_item.parent():
                        self._clear_child_highlight(); self._hide_indicator_line(tree); event.ignore(); return True
                    self._clear_child_highlight(); self._hide_indicator_line(tree)
                    return False
                # Top-level binder: we don't customize visuals
                self._clear_child_highlight(); self._hide_indicator_line(tree)
                return False

            if event.type() == QEvent.Drop:
                self._clear_child_highlight(); self._hide_indicator_line(tree)
                sel = tree.selectedItems()
                drag_item = sel[0] if sel else tree.currentItem()
                if drag_item is None:
                    return False
                kind = drag_item.data(0, 1001)

                # Page: compute new order list; rebuild only (no manual item move) for consistency
                if kind == "page":
                    parent_item = drag_item.parent()
                    if parent_item is None:
                        return True
                    target_item = tree.itemAt(pos_vp) if pos_vp is not None else None
                    # No-op if dropping onto itself
                    if target_item is drag_item:
                        event.accept()
                        return True
                    # Gather current sibling pages excluding drag
                    siblings = []
                    drag_id = int(drag_item.data(0, 1000)) if drag_item.data(0, 1000) is not None else None
                    for j in range(parent_item.childCount()):
                        ch = parent_item.child(j)
                        if ch.data(0, 1001) == "page":
                            cid = int(ch.data(0, 1000))
                            if cid != drag_id:
                                siblings.append((cid, ch))
                    # Determine insertion index
                    insert_idx = len(siblings)  # default append
                    if target_item is not None and target_item.data(0, 1001) == "page" and target_item.parent() is parent_item:
                        rect = tree.visualItemRect(target_item)
                        h = rect.height() or 1
                        rel_y = pos_vp.y() - rect.top() if pos_vp is not None else 0
                        move_before = rel_y < (h * 0.5)
                        # Find target logical index among siblings list
                        t_cid = int(target_item.data(0, 1000))
                        for idx, (cid, _) in enumerate(siblings):
                            if cid == t_cid:
                                insert_idx = idx if move_before else idx + 1
                                break
                    # Build new ordered id list
                    new_order_ids = []
                    for idx, (cid, _) in enumerate(siblings):
                        if idx == insert_idx:
                            if drag_id is not None:
                                new_order_ids.append(drag_id)
                        new_order_ids.append(cid)
                    if insert_idx == len(siblings) and drag_id is not None:
                        new_order_ids.append(drag_id)
                    # Persist new order
                    def _persist_rebuild():
                        try:
                            from db_pages import set_pages_order
                            db_path = getattr(self._window, "_db_path", "notes.db")
                            # Determine section & parent_page_id
                            def _section_item_of(item):
                                cur = item
                                while cur is not None and cur.data(0, 1001) != "section":
                                    cur = cur.parent()
                                return cur
                            sec_item = _section_item_of(parent_item)
                            sec_id = int(sec_item.data(0, 1000)) if sec_item is not None else None
                            parent_pid = None if parent_item.data(0, 1001) == "section" else int(parent_item.data(0, 1000))
                            if new_order_ids and sec_id is not None:
                                set_pages_order(int(sec_id), new_order_ids, db_path, parent_page_id=parent_pid)
                            # Rebuild subtree from DB
                            self._refresh_page_subtree(parent_item, keep_expanded=True)
                            # Reselect moved page
                            if drag_id is not None:
                                found = self._find_descendant_by_id_kind(parent_item, "page", drag_id)
                                if found is not None:
                                    tree.setCurrentItem(found)
                        except Exception:
                            pass
                    QTimer.singleShot(0, _persist_rebuild)
                    event.accept()
                    return True

                # Section reordering within the same binder
                if kind == "section" and drag_item.parent() is not None:
                    src_binder_item = drag_item.parent()
                    tgt_binder_item = self._top_level_item_at(tree, pos_vp)
                    if tgt_binder_item is None or tgt_binder_item is not src_binder_item:
                        return True
                    try:
                        keep_expanded = bool(drag_item.isExpanded())
                    except Exception:
                        keep_expanded = False
                    QTimer.singleShot(
                        0,
                        lambda: self._persist_section_order(
                            tree, src_binder_item, drag_item, keep_expanded=keep_expanded
                        ),
                    )
                    return False

                # Top-level binder reordering
                if drag_item.parent() is None:
                    QTimer.singleShot(0, lambda: self._persist_binder_order(tree))
                    return False

            if event.type() == QEvent.DragLeave:
                self._clear_child_highlight(); self._hide_indicator_line(tree)
                return False
        except Exception:
            pass
        return False

    def _apply_child_highlight(self, item: QtWidgets.QTreeWidgetItem):
        try:
            if self._highlight_item is item:
                return
            self._clear_child_highlight()
            self._highlight_item = item
            self._highlight_kind = "child"
            item.setBackground(0, QBrush(QColor(255, 233, 179)))  # soft amber
        except Exception:
            pass

    def _clear_child_highlight(self):
        try:
            if self._highlight_item is not None:
                self._highlight_item.setBackground(0, QBrush())
            self._highlight_item = None
            self._highlight_kind = None
        except Exception:
            pass

    def _ensure_indicator_line(self, tree: QtWidgets.QTreeWidget):
        try:
            if self._indicator_line is None:
                self._indicator_line = QtWidgets.QFrame(tree.viewport())
                self._indicator_line.setObjectName("leftTreeDropIndicator")
                self._indicator_line.setFrameShape(QtWidgets.QFrame.NoFrame)
                self._indicator_line.setStyleSheet("background: rgba(0,120,215,0.9);")
                self._indicator_line.hide()
        except Exception:
            pass

    def _show_indicator_line(self, tree: QtWidgets.QTreeWidget, item_rect: QRect, position: str):
        try:
            self._ensure_indicator_line(tree)
            if self._indicator_line is None:
                return
            y = item_rect.top() - 1 if position == "before" else item_rect.bottom()
            y = max(0, y)
            w = tree.viewport().width()
            self._indicator_line.setGeometry(0, y, w, 3)
            self._indicator_line.show()
        except Exception:
            pass

    def _hide_indicator_line(self, tree: QtWidgets.QTreeWidget):
        try:
            if self._indicator_line is not None:
                self._indicator_line.hide()
        except Exception:
            pass

    def _refresh_page_subtree(self, parent_item: QtWidgets.QTreeWidgetItem, keep_expanded: bool = True):
        """Rebuild the page-children under the given parent (section or page) from DB.

        Ensures the tree view matches persisted order/parentage after DnD.
        """
        if parent_item is None:
            return
        try:
            kind = parent_item.data(0, 1001)
            db_path = getattr(self._window, "_db_path", "notes.db")
            tree = parent_item.treeWidget()
            # Preserve expansion state
            expanded = bool(parent_item.isExpanded()) if keep_expanded else False

            # Clear existing children
            try:
                parent_item.takeChildren()
            except Exception:
                pass

            # Fetch and rebuild children
            if kind == "section":
                try:
                    sec_id = int(parent_item.data(0, 1000))
                except Exception:
                    return
                try:
                    from db_pages import get_root_pages_by_section_id
                    pages_root = get_root_pages_by_section_id(sec_id, db_path)
                except Exception:
                    pages_root = []
                try:
                    pages_sorted = sorted(pages_root, key=lambda p: (p[6], p[0]))
                except Exception:
                    pages_sorted = pages_root
                for p in pages_sorted:
                    page_id = p[0]
                    page_title = str(p[2])
                    page_item = QtWidgets.QTreeWidgetItem([page_title])
                    page_item.setData(0, 1000, page_id)
                    try:
                        page_item.setData(0, 1001, "page")
                        page_item.setData(0, 1002, sec_id)
                    except Exception:
                        pass
                    try:
                        pflags = page_item.flags()
                        from page_editor import is_two_column_ui as _is_two_col
                        if _is_two_col(tree.window()):
                            pflags = pflags | Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled
                        else:
                            pflags = (pflags | Qt.ItemIsEnabled) & ~(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
                        page_item.setFlags(pflags)
                    except Exception:
                        pass
                    parent_item.addChild(page_item)
                    # Recurse children
                    try:
                        from ui_sections import _add_child_pages_recursively
                        _add_child_pages_recursively(sec_id, int(page_id), page_item, db_path)
                    except Exception:
                        pass
            elif kind == "page":
                try:
                    sec_id = int(parent_item.data(0, 1002))
                    parent_page_id = int(parent_item.data(0, 1000))
                except Exception:
                    return
                try:
                    from db_pages import get_child_pages
                    children = get_child_pages(sec_id, parent_page_id, db_path)
                except Exception:
                    children = []
                try:
                    children_sorted = sorted(children, key=lambda p: (p[6], p[0]))
                except Exception:
                    children_sorted = children
                for p in children_sorted:
                    page_id = p[0]
                    page_title = str(p[2])
                    page_item = QtWidgets.QTreeWidgetItem([page_title])
                    page_item.setData(0, 1000, page_id)
                    try:
                        page_item.setData(0, 1001, "page")
                        page_item.setData(0, 1002, sec_id)
                    except Exception:
                        pass
                    try:
                        pflags = page_item.flags()
                        from page_editor import is_two_column_ui as _is_two_col
                        if _is_two_col(tree.window()):
                            # Subpages: enable DnD like root pages for sibling-only reorder
                            pflags = pflags | Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled
                        else:
                            pflags = (pflags | Qt.ItemIsEnabled) & ~(Qt.ItemIsSelectable)
                        page_item.setFlags(pflags)
                    except Exception:
                        pass
                    parent_item.addChild(page_item)
                    try:
                        from ui_sections import _add_child_pages_recursively
                        _add_child_pages_recursively(sec_id, int(page_id), page_item, db_path)
                    except Exception:
                        pass

            # Restore expansion
            try:
                parent_item.setExpanded(expanded)
            except Exception:
                pass
        except Exception:
            pass

    def _find_descendant_by_id_kind(self, parent_item: QtWidgets.QTreeWidgetItem, kind: str, id_value: int):
        if parent_item is None:
            return None
        try:
            for j in range(parent_item.childCount()):
                ch = parent_item.child(j)
                try:
                    if ch.data(0, 1001) == kind and int(ch.data(0, 1000)) == int(id_value):
                        return ch
                except Exception:
                    pass
                found = self._find_descendant_by_id_kind(ch, kind, id_value)
                if found is not None:
                    return found
        except Exception:
            pass
        return None

    def _persist_binder_order(self, tree: QtWidgets.QTreeWidget):
        try:
            # Capture current selection id (to restore after repopulate)
            try:
                current = tree.currentItem()
                cur_id = current.data(0, 1000) if current is not None else None
            except Exception:
                cur_id = None
            # Build top-level id order
            ordered_ids = []
            for i in range(tree.topLevelItemCount()):
                top = tree.topLevelItem(i)
                nid = top.data(0, 1000)
                if nid is not None:
                    ordered_ids.append(int(nid))
            if not ordered_ids:
                return
            # Persist order
            db_path = getattr(self._window, "_db_path", "notes.db")
            try:
                from db_access import set_notebooks_order

                set_notebooks_order(ordered_ids, db_path)
            except Exception:
                pass
            # Rebuild and restore selection without expanding others
            try:
                from ui_logic import populate_notebook_names

                populate_notebook_names(self._window, db_path)
                if cur_id is not None:
                    for i in range(tree.topLevelItemCount()):
                        top = tree.topLevelItem(i)
                        if int(top.data(0, 1000)) == int(cur_id):
                            tree.setCurrentItem(top)
                            break
                # Update current notebook context, persist last db state
                if cur_id is not None:
                    try:
                        self._window._current_notebook_id = int(cur_id)
                    except Exception:
                        pass
                    try:
                        from settings_manager import set_last_state

                        set_last_state(notebook_id=int(cur_id))
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            pass

    def _persist_section_order(self, tree: QtWidgets.QTreeWidget, binder_item: QtWidgets.QTreeWidgetItem, moved_section_item: QtWidgets.QTreeWidgetItem, keep_expanded: bool = False):
        """Persist the order of sections under the given binder and refresh its children.

        Keeps the moved section selected and avoids expanding unrelated binders.
        """
        try:
            # Resolve ids up front
            try:
                nb_id = binder_item.data(0, 1000)
                moved_sid = moved_section_item.data(0, 1000)
            except Exception:
                nb_id = None
                moved_sid = None
            # Build ordered section ids for this binder
            section_ids = []
            for j in range(binder_item.childCount()):
                ch = binder_item.child(j)
                if ch and ch.data(0, 1001) == "section":
                    sid = ch.data(0, 1000)
                    if sid is not None:
                        section_ids.append(int(sid))
            if not section_ids:
                return
            # Persist order
            db_path = getattr(self._window, "_db_path", "notes.db")
            try:
                from db_sections import set_sections_order

                if nb_id is not None:
                    set_sections_order(int(nb_id), section_ids, db_path)
            except Exception:
                pass
            # Refresh only this binder’s children and reselect the moved section
            if nb_id is None:
                return
            try:
                # If the section was expanded before the move, expand and select it.
                # If it was collapsed, refresh children without auto-expanding, then reselect and keep it collapsed.
                if keep_expanded and moved_sid is not None:
                    ensure_left_tree_sections(self._window, int(nb_id), select_section_id=int(moved_sid))
                else:
                    ensure_left_tree_sections(self._window, int(nb_id))
                    # Find the freshly rebuilt binder item and the moved section under it
                    target_binder = None
                    for i in range(tree.topLevelItemCount()):
                        top = tree.topLevelItem(i)
                        try:
                            if int(top.data(0, 1000)) == int(nb_id):
                                target_binder = top
                                break
                        except Exception:
                            pass
                    if target_binder is not None and moved_sid is not None:
                        for j in range(target_binder.childCount()):
                            sec = target_binder.child(j)
                            try:
                                if int(sec.data(0, 1000)) == int(moved_sid):
                                    # Ensure it remains collapsed and selected
                                    try:
                                        if sec.isExpanded():
                                            sec.setExpanded(False)
                                    except Exception:
                                        pass
                                    try:
                                        tree.setCurrentItem(sec)
                                    except Exception:
                                        pass
                                    break
                            except Exception:
                                pass
            except Exception:
                pass
        except Exception:
            pass
