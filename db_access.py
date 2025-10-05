"""
db_access.py
Provides functions to retrieve notebooks from the database.
"""
import sqlite3

def get_notebooks(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM notebooks ORDER BY order_index, id")
    rows = cursor.fetchall()
    conn.close()
    return rows

def _get_next_notebook_order_index(db_path: str) -> int:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(MAX(order_index), 0) FROM notebooks")
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
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM notebooks WHERE id = ?", (notebook_id,))
    conn.commit()
    conn.close()
