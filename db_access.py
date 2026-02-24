"""
db_access.py
Provides functions to retrieve notebooks from the database.
"""

import sqlite3


def get_notebooks(db_path, include_deleted: bool = False):
    """Get all notebooks, optionally including soft-deleted ones."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    if include_deleted:
        cursor.execute("SELECT * FROM notebooks ORDER BY order_index, id")
    else:
        cursor.execute("SELECT * FROM notebooks WHERE deleted_at IS NULL ORDER BY order_index, id")
    rows = cursor.fetchall()
    conn.close()
    return rows


def _get_next_notebook_order_index(db_path: str) -> int:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(MAX(order_index), 0) FROM notebooks WHERE deleted_at IS NULL")
    max_idx = cur.fetchone()[0] or 0
    conn.close()
    return int(max_idx) + 1


def create_notebook(title: str, db_path: str) -> int:
    order_index = _get_next_notebook_order_index(db_path)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO notebooks (title, order_index) VALUES (?, ?)",
        (title, order_index),
    )
    conn.commit()
    nid = cur.lastrowid
    conn.close()
    return nid


def rename_notebook(notebook_id: int, title: str, db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "UPDATE notebooks SET title = ?, modified_at = datetime('now') WHERE id = ?",
        (title, notebook_id),
    )
    conn.commit()
    conn.close()


def delete_notebook(notebook_id: int, db_path: str):
    """Soft-delete a notebook by setting deleted_at timestamp.
    
    Also soft-deletes all sections and pages within the notebook (cascade).
    """
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Soft-delete all pages in all sections of this notebook
        cur.execute(
            """UPDATE pages SET deleted_at = datetime('now')
               WHERE section_id IN (SELECT id FROM sections WHERE notebook_id = ?)
               AND deleted_at IS NULL""",
            (notebook_id,)
        )
        # Soft-delete all sections in this notebook
        cur.execute(
            "UPDATE sections SET deleted_at = datetime('now') WHERE notebook_id = ? AND deleted_at IS NULL",
            (notebook_id,)
        )
        # Soft-delete the notebook itself
        cur.execute(
            "UPDATE notebooks SET deleted_at = datetime('now') WHERE id = ?",
            (notebook_id,)
        )
        conn.commit()
    finally:
        conn.close()


def restore_notebook(notebook_id: int, db_path: str):
    """Restore a soft-deleted notebook and all its sections and pages."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Restore the notebook
        cur.execute(
            "UPDATE notebooks SET deleted_at = NULL WHERE id = ?",
            (notebook_id,)
        )
        # Restore all sections in this notebook
        cur.execute(
            "UPDATE sections SET deleted_at = NULL WHERE notebook_id = ?",
            (notebook_id,)
        )
        # Restore all pages in all sections of this notebook
        cur.execute(
            """UPDATE pages SET deleted_at = NULL
               WHERE section_id IN (SELECT id FROM sections WHERE notebook_id = ?)""",
            (notebook_id,)
        )
        conn.commit()
    finally:
        conn.close()


def permanently_delete_notebook(notebook_id: int, db_path: str):
    """Permanently delete a notebook and all its sections/pages from the database."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        # Delete all pages in all sections of this notebook
        cur.execute(
            "DELETE FROM pages WHERE section_id IN (SELECT id FROM sections WHERE notebook_id = ?)",
            (notebook_id,)
        )
        # Delete all sections in this notebook
        cur.execute(
            "DELETE FROM sections WHERE notebook_id = ?",
            (notebook_id,)
        )
        # Delete the notebook itself
        cur.execute("DELETE FROM notebooks WHERE id = ?", (notebook_id,))
        conn.commit()
    finally:
        conn.close()


def set_notebooks_order(ordered_ids, db_path: str):
    """Update order_index for all notebooks based on the given ordered list of ids.
    Any ids not present are ignored; unknown ids are skipped.
    """
    if not ordered_ids:
        return
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        # Assign sequential order_index starting at 0 to maintain stable ordering
        for idx, nid in enumerate(ordered_ids):
            try:
                cur.execute(
                    "UPDATE notebooks SET order_index = ?, modified_at = datetime('now') WHERE id = ?",
                    (int(idx), int(nid)),
                )
            except Exception:
                # Skip bad ids; continue others
                pass
        conn.commit()
    finally:
        conn.close()


def empty_all_deleted(db_path: str) -> dict:
    """Permanently delete all soft-deleted items from the database.
    
    Returns a dict with counts: {'notebooks': N, 'sections': N, 'pages': N}
    """
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    counts = {'notebooks': 0, 'sections': 0, 'pages': 0}
    try:
        # Count before deletion
        cur.execute("SELECT COUNT(*) FROM pages WHERE deleted_at IS NOT NULL")
        counts['pages'] = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM sections WHERE deleted_at IS NOT NULL")
        counts['sections'] = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM notebooks WHERE deleted_at IS NOT NULL")
        counts['notebooks'] = cur.fetchone()[0]
        
        # Delete in order: pages first, then sections, then notebooks
        cur.execute("DELETE FROM pages WHERE deleted_at IS NOT NULL")
        cur.execute("DELETE FROM sections WHERE deleted_at IS NOT NULL")
        cur.execute("DELETE FROM notebooks WHERE deleted_at IS NOT NULL")
        conn.commit()
    finally:
        conn.close()
    return counts


def get_deleted_counts(db_path: str) -> dict:
    """Get counts of soft-deleted items.
    
    Returns a dict with counts: {'notebooks': N, 'sections': N, 'pages': N, 'total': N}
    """
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        cur.execute("SELECT COUNT(*) FROM pages WHERE deleted_at IS NOT NULL")
        pages = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM sections WHERE deleted_at IS NOT NULL")
        sections = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM notebooks WHERE deleted_at IS NOT NULL")
        notebooks = cur.fetchone()[0]
        return {
            'notebooks': notebooks,
            'sections': sections,
            'pages': pages,
            'total': notebooks + sections + pages
        }
    finally:
        conn.close()
