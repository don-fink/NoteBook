"""maintenance_order.py

Programmatic order_index normalization utilities for notebooks, sections, and pages.
Shared by the UI menu action so users can normalize without running the script.

Functions:
  collect_changes(db_path) -> dict: returns mapping with keys notebooks/sections/pages each
      holding list of (id, new_order_index) tuples (empty list if no changes for that entity).
  apply_changes(db_path, changes) -> None: applies updates in a transaction.
  summarize(changes) -> str: human-readable summary line per entity.

Resequencing preserves relative order defined by (order_index, id) and produces
compact 1..N sequences per sibling group.
"""
from __future__ import annotations
import sqlite3
from typing import Dict, List, Tuple


def _fetchall(con: sqlite3.Connection, sql: str, params=()):
    cur = con.cursor()
    cur.execute(sql, params)
    return cur.fetchall()


def _normalize_sequence(rows: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    if not rows:
        return []
    ordered = sorted(rows, key=lambda r: (r[1], r[0]))
    changes = []
    for new_idx, (rid, cur_idx) in enumerate(ordered, start=1):
        if cur_idx != new_idx:
            changes.append((rid, new_idx))
    return changes


def _has_parent_column(con: sqlite3.Connection) -> bool:
    cur = con.cursor()
    cur.execute("PRAGMA table_info(pages)")
    return any(r[1] == "parent_page_id" for r in cur.fetchall())


def collect_changes(db_path: str) -> Dict[str, List[Tuple[int, int]]]:
    con = sqlite3.connect(db_path)
    try:
        changes: Dict[str, List[Tuple[int, int]]] = {"notebooks": [], "sections": [], "pages": []}
        # Notebooks
        n_rows = _fetchall(con, "SELECT id, order_index FROM notebooks ORDER BY order_index, id")
        changes["notebooks"] = _normalize_sequence(n_rows)
        # Sections per notebook
        nb_ids = [r[0] for r in _fetchall(con, "SELECT id FROM notebooks ORDER BY id")]
        all_sec_changes: List[Tuple[int, int]] = []
        for nb_id in nb_ids:
            rows = _fetchall(con, "SELECT id, order_index FROM sections WHERE notebook_id = ? ORDER BY order_index, id", (nb_id,))
            all_sec_changes.extend(_normalize_sequence(rows))
        changes["sections"] = all_sec_changes
        # Pages groups
        if _has_parent_column(con):
            groups = _fetchall(con, "SELECT DISTINCT section_id, parent_page_id FROM pages ORDER BY section_id, parent_page_id")
            all_page_changes: List[Tuple[int, int]] = []
            for section_id, parent_page_id in groups:
                if parent_page_id is None:
                    rows = _fetchall(con, "SELECT id, order_index FROM pages WHERE section_id = ? AND parent_page_id IS NULL ORDER BY order_index, id", (section_id,))
                else:
                    rows = _fetchall(con, "SELECT id, order_index FROM pages WHERE section_id = ? AND parent_page_id = ? ORDER BY order_index, id", (section_id, parent_page_id))
                all_page_changes.extend(_normalize_sequence(rows))
            changes["pages"] = all_page_changes
        else:
            groups = _fetchall(con, "SELECT DISTINCT section_id FROM pages ORDER BY section_id")
            all_page_changes: List[Tuple[int, int]] = []
            for (section_id,) in groups:
                rows = _fetchall(con, "SELECT id, order_index FROM pages WHERE section_id = ? ORDER BY order_index, id", (section_id,))
                all_page_changes.extend(_normalize_sequence(rows))
            changes["pages"] = all_page_changes
        return changes
    finally:
        con.close()


def apply_changes(db_path: str, changes: Dict[str, List[Tuple[int, int]]]):
    con = sqlite3.connect(db_path)
    try:
        cur = con.cursor()
        cur.execute("BEGIN")
        for nid, new_ord in changes.get("notebooks", []):
            cur.execute("UPDATE notebooks SET order_index = ? WHERE id = ?", (new_ord, nid))
        for sid, new_ord in changes.get("sections", []):
            cur.execute("UPDATE sections SET order_index = ? WHERE id = ?", (new_ord, sid))
        for pid, new_ord in changes.get("pages", []):
            cur.execute("UPDATE pages SET order_index = ? WHERE id = ?", (new_ord, pid))
        cur.execute("COMMIT")
    except Exception:
        try:
            cur.execute("ROLLBACK")
        except Exception:
            pass
        raise
    finally:
        con.close()


def summarize(changes: Dict[str, List[Tuple[int, int]]]) -> str:
    return "\n".join(
        [
            f"notebooks: {('no changes' if not changes.get('notebooks') else str(len(changes['notebooks'])) + ' updates')}",
            f"sections: {('no changes' if not changes.get('sections') else str(len(changes['sections'])) + ' updates')}",
            f"pages: {('no changes' if not changes.get('pages') else str(len(changes['pages'])) + ' updates')}",
        ]
    )
