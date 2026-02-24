"""
db_pages.py
Provides functions to retrieve pages for a given section from the database.
"""

import sqlite3


def get_pages_by_section_id(section_id, db_path, include_deleted: bool = False):
    """Return all pages for a section (flat). Legacy helper.

    Note: With hierarchical pages, prefer get_root_pages_by_section_id() and get_child_pages().
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    if include_deleted:
        cursor.execute("SELECT * FROM pages WHERE section_id = ?", (section_id,))
    else:
        cursor.execute("SELECT * FROM pages WHERE section_id = ? AND deleted_at IS NULL", (section_id,))
    rows = cursor.fetchall()
    conn.close()
    return rows


def get_root_pages_by_section_id(section_id: int, db_path: str, include_deleted: bool = False):
    """Return top-level pages in a section (parent_page_id IS NULL)."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    if include_deleted:
        cur.execute(
            "SELECT * FROM pages WHERE section_id = ? AND (parent_page_id IS NULL OR parent_page_id = '') ORDER BY order_index ASC, id ASC",
            (int(section_id),),
        )
    else:
        cur.execute(
            "SELECT * FROM pages WHERE section_id = ? AND (parent_page_id IS NULL OR parent_page_id = '') AND deleted_at IS NULL ORDER BY order_index ASC, id ASC",
            (int(section_id),),
        )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_child_pages(section_id: int, parent_page_id: int, db_path: str, include_deleted: bool = False):
    """Return direct child pages under a parent page within the same section."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    if include_deleted:
        cur.execute(
            "SELECT * FROM pages WHERE section_id = ? AND parent_page_id = ? ORDER BY order_index ASC, id ASC",
            (int(section_id), int(parent_page_id)),
        )
    else:
        cur.execute(
            "SELECT * FROM pages WHERE section_id = ? AND parent_page_id = ? AND deleted_at IS NULL ORDER BY order_index ASC, id ASC",
            (int(section_id), int(parent_page_id)),
        )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_page_by_id(page_id: int, db_path: str):
    """Return a single page row by id, or None if not found."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT * FROM pages WHERE id = ?", (int(page_id),))
    row = cur.fetchone()
    conn.close()
    return row


def _get_next_page_order_index(section_id: int, db_path: str, parent_page_id=None) -> int:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    if parent_page_id is None:
        cur.execute(
            "SELECT COALESCE(MAX(order_index), 0) FROM pages WHERE section_id = ? AND (parent_page_id IS NULL OR parent_page_id = '') AND deleted_at IS NULL",
            (section_id,),
        )
    else:
        cur.execute(
            "SELECT COALESCE(MAX(order_index), 0) FROM pages WHERE section_id = ? AND parent_page_id = ? AND deleted_at IS NULL",
            (section_id, int(parent_page_id)),
        )
    max_idx = cur.fetchone()[0] or 0
    conn.close()
    return int(max_idx) + 1


def create_page(section_id: int, title: str, db_path: str, parent_page_id: int = None) -> int:
    """Create a new page in a section (optionally as a child of another page) and return its id."""
    order_index = _get_next_page_order_index(section_id, db_path, parent_page_id=parent_page_id)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO pages (section_id, title, content_html, order_index, parent_page_id)
        VALUES (?, ?, ?, ?, ?)
        """,
        (section_id, title, "", order_index, parent_page_id),
    )
    conn.commit()
    page_id = cur.lastrowid
    conn.close()
    return page_id


def update_page_title(page_id: int, title: str, db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "UPDATE pages SET title = ?, modified_at = datetime('now') WHERE id = ?",
        (title, page_id),
    )
    conn.commit()
    conn.close()


def update_page_content(page_id: int, content_html: str, db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "UPDATE pages SET content_html = ?, modified_at = datetime('now') WHERE id = ?",
        (content_html, page_id),
    )
    conn.commit()
    conn.close()


def delete_page(page_id: int, db_path: str):
    """Soft-delete a page and all its descendants by setting deleted_at timestamp."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Collect descendants with iterative BFS to avoid Python recursion limits
        to_delete = [int(page_id)]
        idx = 0
        while idx < len(to_delete):
            pid = to_delete[idx]
            cur.execute("SELECT id FROM pages WHERE parent_page_id = ?", (int(pid),))
            children = [r[0] for r in cur.fetchall()]
            to_delete.extend(int(c) for c in children)
            idx += 1
        # Soft-delete all pages (parent and descendants)
        for pid in to_delete:
            cur.execute(
                "UPDATE pages SET deleted_at = datetime('now') WHERE id = ? AND deleted_at IS NULL",
                (int(pid),)
            )
        conn.commit()
    finally:
        conn.close()


def restore_page(page_id: int, db_path: str):
    """Restore a soft-deleted page and all its descendants."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Collect descendants with iterative BFS
        to_restore = [int(page_id)]
        idx = 0
        while idx < len(to_restore):
            pid = to_restore[idx]
            cur.execute("SELECT id FROM pages WHERE parent_page_id = ?", (int(pid),))
            children = [r[0] for r in cur.fetchall()]
            to_restore.extend(int(c) for c in children)
            idx += 1
        # Restore all pages (parent and descendants)
        for pid in to_restore:
            cur.execute(
                "UPDATE pages SET deleted_at = NULL WHERE id = ?",
                (int(pid),)
            )
        conn.commit()
    finally:
        conn.close()


def permanently_delete_page(page_id: int, db_path: str):
    """Permanently delete a page and all its descendants from the database."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Collect descendants with iterative BFS to avoid Python recursion limits
        to_delete = [int(page_id)]
        idx = 0
        while idx < len(to_delete):
            pid = to_delete[idx]
            cur.execute("SELECT id FROM pages WHERE parent_page_id = ?", (int(pid),))
            children = [r[0] for r in cur.fetchall()]
            to_delete.extend(int(c) for c in children)
            idx += 1
        # Delete from leaves up (reverse order)
        for pid in reversed(to_delete):
            cur.execute("DELETE FROM pages WHERE id = ?", (int(pid),))
        conn.commit()
    finally:
        conn.close()


def set_pages_order(section_id: int, ordered_page_ids: list, db_path: str, parent_page_id: int = None):
    """Update order_index for sibling pages (same section and same parent_page_id)."""
    if not isinstance(ordered_page_ids, (list, tuple)):
        return
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        order_val = 1
        for pid in ordered_page_ids:
            try:
                pid_int = int(pid)
            except Exception:
                continue
            if parent_page_id is None:
                cur.execute(
                    "UPDATE pages SET order_index = ? WHERE id = ? AND section_id = ? AND (parent_page_id IS NULL OR parent_page_id = '')",
                    (order_val, pid_int, section_id),
                )
            else:
                cur.execute(
                    "UPDATE pages SET order_index = ? WHERE id = ? AND section_id = ? AND parent_page_id = ?",
                    (order_val, pid_int, section_id, int(parent_page_id)),
                )
            order_val += 1
        conn.commit()
    finally:
        conn.close()


def set_pages_parent_and_order(target_section_id: int, ordered_page_ids: list, db_path: str, parent_page_id: int = None):
    """Move pages to target_section_id and assign sequential order_index among siblings under parent_page_id.
    Supports cross-section and cross-parent reparenting.
    """
    if not isinstance(ordered_page_ids, (list, tuple)) or not ordered_page_ids:
        return
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        order_val = 1
        for pid in ordered_page_ids:
            try:
                pid_int = int(pid)
            except Exception:
                continue
            cur.execute(
                "UPDATE pages SET section_id = ?, order_index = ?, parent_page_id = ? WHERE id = ?",
                (int(target_section_id), order_val, parent_page_id, pid_int),
            )
            order_val += 1
        conn.commit()
    finally:
        conn.close()
