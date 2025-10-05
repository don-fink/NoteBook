"""
db_pages.py
Provides functions to retrieve pages for a given section from the database.
"""
import sqlite3

def get_pages_by_section_id(section_id, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM pages WHERE section_id = ?", (section_id,))
    rows = cursor.fetchall()
    conn.close()
    return rows

def _get_next_page_order_index(section_id: int, db_path: str) -> int:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(MAX(order_index), 0) FROM pages WHERE section_id = ?", (section_id,))
    max_idx = cur.fetchone()[0] or 0
    conn.close()
    return int(max_idx) + 1

def create_page(section_id: int, title: str, db_path: str) -> int:
    """Create a new page in a section and return its id."""
    order_index = _get_next_page_order_index(section_id, db_path)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO pages (section_id, title, content_html, order_index)
        VALUES (?, ?, ?, ?)
        """,
        (section_id, title, "", order_index),
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
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM pages WHERE id = ?", (page_id,))
    conn.commit()
    conn.close()


def set_pages_order(section_id: int, ordered_page_ids: list, db_path: str):
    """Update order_index for all pages in a section based on the given ordered list of ids."""
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
            cur.execute(
                "UPDATE pages SET order_index = ? WHERE id = ? AND section_id = ?",
                (order_val, pid_int, section_id),
            )
            order_val += 1
        conn.commit()
    finally:
        conn.close()


def set_pages_parent_and_order(target_section_id: int, ordered_page_ids: list, db_path: str):
    """Move pages to target_section_id (if not already) and assign sequential order_index.
    This supports cross-section reparenting via drag-and-drop.
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
                "UPDATE pages SET section_id = ?, order_index = ? WHERE id = ?",
                (target_section_id, order_val, pid_int),
            )
            order_val += 1
        conn.commit()
    finally:
        conn.close()
