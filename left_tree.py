"""
left_tree.py

Stable surface for left‑panel (Binder/Sections/Pages) helpers while we migrate
implementation away from ui_tabs. Also hosts the LeftTreeDnDFilter used to
reorder top-level binders via drag-and-drop.
"""

from PyQt5 import QtWidgets
from PyQt5.QtCore import QEvent, QObject, QTimer

# Temporary re-exports while we still rely on ui_tabs under the hood
from ui_tabs import (  # noqa: F401
    ensure_left_tree_sections,  # Expand binder and populate sections/pages
    refresh_for_notebook,  # Rebuild center/right panes for a notebook
)

# Convenience re-exports for selection helpers
from ui_tabs import (  # noqa: F401
    _select_left_binder as select_left_binder,
    _select_tree_section as select_tree_section,
)

# Additional helpers used by callers
from ui_tabs import (  # noqa: F401
    _select_left_tree_page as select_left_tree_page,
    _update_left_tree_page_title as update_left_tree_page_title,
)


class LeftTreeDnDFilter(QObject):
    """Constrain drag/drop to top-level binder reordering and persist order.

    Installed on the left binder tree and its viewport. After a successful
    drop, it persists order to the DB, repopulates the tree, and restores the
    previous binder selection without auto-expanding other binders.
    """

    def __init__(self, window):
        super().__init__()
        self._window = window

    def _top_level_item_at(self, tree: QtWidgets.QTreeWidget, pos_vp):
        item = tree.itemAt(pos_vp)
        if item is None:
            return None
        return item if item.parent() is None else item.parent()

    def eventFilter(self, obj, event):
        try:
            # Resolve tree and viewport consistently
            if isinstance(obj, QtWidgets.QTreeWidget):
                tree = obj
                viewport = tree.viewport()
                pos_vp = viewport.mapFrom(tree, event.pos())
            elif isinstance(obj, QtWidgets.QWidget) and isinstance(
                obj.parent(), QtWidgets.QTreeWidget
            ):
                tree = obj.parent()
                viewport = obj
                pos_vp = event.pos()
            else:
                return False

            if event.type() in (QEvent.DragEnter, QEvent.DragMove):
                # Stricter drop zones: only allow indicators over valid targets
                # Determine dragged item kind via current/selection
                try:
                    sel = tree.selectedItems()
                    drag_item = sel[0] if sel else tree.currentItem()
                except Exception:
                    drag_item = None
                target = tree.itemAt(pos_vp)
                if drag_item is None:
                    return False
                kind = drag_item.data(0, 1001)
                if kind == "page":
                    # Only allow when hovering over a page within the SAME section
                    if target is None or target.data(0, 1001) != "page":
                        event.ignore()
                        return True
                    if target.parent() is not drag_item.parent():
                        event.ignore()
                        return True
                    # Valid: let Qt paint indicator
                    return False
                if kind == "section":
                    # Only allow when hovering over a section within the SAME binder
                    if target is None or target.data(0, 1001) != "section":
                        event.ignore()
                        return True
                    if target.parent() is None or drag_item.parent() is None:
                        event.ignore()
                        return True
                    if target.parent() is not drag_item.parent():
                        event.ignore()
                        return True
                    return False
                # Top-level binder: allow default handling (works well)
                return False
            if event.type() == QEvent.Drop:
                # Determine the dragged item from selection (more reliable than currentItem on drop)
                sel = tree.selectedItems()
                drag_item = sel[0] if sel else tree.currentItem()
                if drag_item is None:
                    return False
                kind = drag_item.data(0, 1001)
                # Page reordering within the same section
                if kind == "page" and drag_item.parent() is not None:
                    src_section_item = drag_item.parent()
                    target_item = tree.itemAt(pos_vp)
                    # Only allow within the same section
                    if target_item is None:
                        return True
                    if target_item.data(0, 1001) == "page":
                        if target_item.parent() is not src_section_item:
                            return True
                    elif target_item.data(0, 1001) == "section":
                        if target_item is not src_section_item:
                            return True
                    else:
                        return True
                    # Let Qt move the row visually, then persist the new order without rebuilding the tree
                    moved_pid = drag_item.data(0, 1000)
                    try:
                        moved_pid = int(moved_pid)
                    except Exception:
                        pass
                    def _persist_pages_after_move():
                        try:
                            # Build new order directly from the section's children
                            sid = src_section_item.data(0, 1000)
                            try:
                                sid = int(sid)
                            except Exception:
                                pass
                            page_ids = []
                            for j in range(src_section_item.childCount()):
                                ch = src_section_item.child(j)
                                if ch and ch.data(0, 1001) == "page":
                                    pid = ch.data(0, 1000)
                                    if pid is not None:
                                        page_ids.append(int(pid))
                            if page_ids:
                                from db_pages import set_pages_order
                                db_path = getattr(self._window, "_db_path", "notes.db")
                                set_pages_order(int(sid), page_ids, db_path)
                            # Ensure the moved page stays selected
                            try:
                                tree.setCurrentItem(drag_item)
                            except Exception:
                                pass
                        except Exception:
                            pass

                    QTimer.singleShot(0, _persist_pages_after_move)
                    return False
                # Section reordering within the same binder
                if kind == "section" and drag_item.parent() is not None:
                    src_binder_item = drag_item.parent()
                    tgt_binder_item = self._top_level_item_at(tree, pos_vp)
                    # If drop target is not within the same binder, cancel the drop
                    if tgt_binder_item is None or tgt_binder_item is not src_binder_item:
                        return True
                    # Capture expanded state of the moved section so we can restore it
                    try:
                        keep_expanded = bool(drag_item.isExpanded())
                    except Exception:
                        keep_expanded = False
                    # Let Qt perform the move, then persist order for this binder and restore expansion state
                    QTimer.singleShot(
                        0,
                        lambda: self._persist_section_order(
                            tree, src_binder_item, drag_item, keep_expanded=keep_expanded
                        ),
                    )
                    return False
                # Top-level binder reordering
                if drag_item.parent() is None:
                    # Persist after Qt completes the internal move
                    QTimer.singleShot(0, lambda: self._persist_binder_order(tree))
                    return False
        except Exception:
            pass
        return False

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
