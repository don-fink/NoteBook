"""
db_sections.py
Provides functions to retrieve sections for a given notebook from the database.
"""

import sqlite3


def get_sections_by_notebook_id(notebook_id, db_path, include_deleted: bool = False):
    """Get all sections for a notebook, optionally including soft-deleted ones."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    if include_deleted:
        cursor.execute(
            "SELECT * FROM sections WHERE notebook_id = ? ORDER BY order_index, id", (notebook_id,)
        )
    else:
        cursor.execute(
            "SELECT * FROM sections WHERE notebook_id = ? AND deleted_at IS NULL ORDER BY order_index, id", (notebook_id,)
        )
    rows = cursor.fetchall()
    conn.close()
    return rows


def update_section_color(section_id: int, color_hex: str, db_path: str):
    """Set or clear the color for a section. Pass None or '' to clear."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    if color_hex:
        cur.execute(
            "UPDATE sections SET color_hex = ?, modified_at = datetime('now') WHERE id = ?",
            (color_hex, section_id),
        )
    else:
        cur.execute(
            "UPDATE sections SET color_hex = NULL, modified_at = datetime('now') WHERE id = ?",
            (section_id,),
        )
    conn.commit()
    conn.close()


def get_section_color_map(notebook_id: int, db_path: str):
    """Return a dict {section_id: color_hex or None} for a notebook."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, color_hex FROM sections WHERE notebook_id = ?", (notebook_id,))
    rows = cur.fetchall()
    conn.close()
    return {row[0]: row[1] for row in rows}


def _get_next_section_order_index(notebook_id: int, db_path: str) -> int:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "SELECT COALESCE(MAX(order_index), 0) FROM sections WHERE notebook_id = ? AND deleted_at IS NULL", (notebook_id,)
    )
    max_idx = cur.fetchone()[0] or 0
    conn.close()
    return int(max_idx) + 1


def create_section(notebook_id: int, title: str, db_path: str) -> int:
    order_index = _get_next_section_order_index(notebook_id, db_path)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO sections (notebook_id, title, order_index) VALUES (?, ?, ?)",
        (notebook_id, title, order_index),
    )
    conn.commit()
    sid = cur.lastrowid
    conn.close()
    return sid


def rename_section(section_id: int, title: str, db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "UPDATE sections SET title = ?, modified_at = datetime('now') WHERE id = ?",
        (title, section_id),
    )
    conn.commit()
    conn.close()


def delete_section(section_id: int, db_path: str):
    """Soft-delete a section by setting deleted_at timestamp.
    
    Also soft-deletes all pages within the section (cascade).
    """
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Soft-delete all pages in this section
        cur.execute(
            "UPDATE pages SET deleted_at = datetime('now') WHERE section_id = ? AND deleted_at IS NULL",
            (section_id,)
        )
        # Soft-delete the section itself
        cur.execute(
            "UPDATE sections SET deleted_at = datetime('now') WHERE id = ?",
            (section_id,)
        )
        conn.commit()
    finally:
        conn.close()


def restore_section(section_id: int, db_path: str):
    """Restore a soft-deleted section and all its pages."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Restore the section
        cur.execute(
            "UPDATE sections SET deleted_at = NULL WHERE id = ?",
            (section_id,)
        )
        # Restore all pages in this section
        cur.execute(
            "UPDATE pages SET deleted_at = NULL WHERE section_id = ?",
            (section_id,)
        )
        conn.commit()
    finally:
        conn.close()


def permanently_delete_section(section_id: int, db_path: str):
    """Permanently delete a section and all its pages from the database."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Delete all pages in this section
        cur.execute("DELETE FROM pages WHERE section_id = ?", (section_id,))
        # Delete the section itself
        cur.execute("DELETE FROM sections WHERE id = ?", (section_id,))
        conn.commit()
    finally:
        conn.close()


def set_sections_order(notebook_id: int, ordered_section_ids: list, db_path: str):
    """Update order_index for all sections in a notebook based on the given ordered list of ids.
    Missing ids are ignored; extra ids are ignored.
    """
    if not isinstance(ordered_section_ids, (list, tuple)):
        return
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        order_val = 1
        for sid in ordered_section_ids:
            try:
                sid_int = int(sid)
            except Exception:
                continue
            cur.execute(
                "UPDATE sections SET order_index = ? WHERE id = ? AND notebook_id = ?",
                (order_val, sid_int, notebook_id),
            )
            order_val += 1
        conn.commit()
    finally:
        conn.close()


def move_section_up(section_id: int, db_path: str) -> bool:
    """Move a section up (to a smaller order_index) within its notebook. Returns True if moved."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT notebook_id, COALESCE(order_index, 0) FROM sections WHERE id = ?", (section_id,)
        )
        row = cur.fetchone()
        if not row:
            return False
        notebook_id, current_order = row
        # Find previous (smaller) neighbor
        cur.execute(
            """
            SELECT id, COALESCE(order_index, 0) FROM sections
            WHERE notebook_id = ? AND COALESCE(order_index, 0) < ?
            ORDER BY COALESCE(order_index, 0) DESC, id DESC
            LIMIT 1
            """,
            (notebook_id, current_order),
        )
        neighbor = cur.fetchone()
        if not neighbor:
            return False
        neighbor_id, neighbor_order = neighbor
        # Swap order_index values in a transaction
        cur.execute(
            "UPDATE sections SET order_index = ? WHERE id = ?", (neighbor_order, section_id)
        )
        cur.execute(
            "UPDATE sections SET order_index = ? WHERE id = ?", (current_order, neighbor_id)
        )
        conn.commit()
        return True
    finally:
        conn.close()


def move_section_down(section_id: int, db_path: str) -> bool:
    """Move a section down (to a larger order_index) within its notebook. Returns True if moved."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT notebook_id, COALESCE(order_index, 0) FROM sections WHERE id = ?", (section_id,)
        )
        row = cur.fetchone()
        if not row:
            return False
        notebook_id, current_order = row
        # Find next (larger) neighbor
        cur.execute(
            """
            SELECT id, COALESCE(order_index, 0) FROM sections
            WHERE notebook_id = ? AND COALESCE(order_index, 0) > ?
            ORDER BY COALESCE(order_index, 0) ASC, id ASC
            LIMIT 1
            """,
            (notebook_id, current_order),
        )
        neighbor = cur.fetchone()
        if not neighbor:
            return False
        neighbor_id, neighbor_order = neighbor
        # Swap order_index values in a transaction
        cur.execute(
            "UPDATE sections SET order_index = ? WHERE id = ?", (neighbor_order, section_id)
        )
        cur.execute(
            "UPDATE sections SET order_index = ? WHERE id = ?", (current_order, neighbor_id)
        )
        conn.commit()
        return True
    finally:
        conn.close()
