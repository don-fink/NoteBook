"""
services.selection

Selection/context helpers centralizing how we determine the current notebook,
section, and page in two‑pane mode. For now, this mirrors existing logic and
delegates to window attributes populated by ui_tabs/main.
"""

from typing import Optional, Tuple
from PyQt5 import QtWidgets


def current_context(window) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """Return (notebook_id, section_id, page_id) from window state and selection.

    - Prefers window._current_notebook_id / _current_section_id / _current_page_by_section
    - Falls back to left tree selection if available
    """

    nb_id = getattr(window, "_current_notebook_id", None)
    sid = getattr(window, "_current_section_id", None)
    pid = None
    try:
        if sid is not None:
            pid = getattr(window, "_current_page_by_section", {}).get(int(sid))
    except Exception:
        pid = getattr(window, "_current_page_by_section", {}).get(sid)

    if sid is None or pid is None:
        tree = window.findChild(QtWidgets.QTreeWidget, "notebookName")
        cur = tree.currentItem() if tree is not None else None
        if cur is not None:
            kind = cur.data(0, 1001)
            if kind == "page":
                pid = cur.data(0, 1000)
                sid = cur.data(0, 1002)
                parent = cur.parent()
                if parent is not None and nb_id is None:
                    p2 = parent.parent()
                    if p2 is None:
                        nb_id = parent.data(0, 1000)
            elif kind == "section" and sid is None:
                sid = cur.data(0, 1000)
                parent = cur.parent()
                if parent is not None and nb_id is None:
                    nb_id = parent.data(0, 1000)

    return (nb_id, sid, pid)
