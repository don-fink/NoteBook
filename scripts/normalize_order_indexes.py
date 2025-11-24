"""
normalize_order_indexes.py

Utility script to normalize (re-sequence) order_index values for:
  - notebooks (top-level binders)
  - sections per notebook
  - pages per (section_id, parent_page_id) sibling group

Why: Over time (inserts, deletes, migrations) order_index can develop gaps or
duplicates. While the app sorts by (order_index, id) deterministically, a compact
sequential range (1..N) is simpler to reason about and prevents edge cases in
relative ordering features.

Usage (PowerShell examples on Windows):
  python scripts/normalize_order_indexes.py notes_dev.db --dry-run
  python scripts/normalize_order_indexes.py notes.db --apply

Flags:
  --dry-run  (default) Show planned changes without modifying the database.
  --apply    Perform updates inside a single transaction.
  --verbose  Print each group details.

Exit codes:
  0 success (or no changes needed)
  1 error / database not found

Idempotent: running again after --apply should result in 'No changes needed'.

Safety: Uses a single transaction for updates; rolls back automatically
on exceptions.
"""
from __future__ import annotations
import argparse
import os
import sqlite3
import sys
from typing import List, Tuple, Dict

# --- Helpers -----------------------------------------------------------------

def _fetchall(con: sqlite3.Connection, sql: str, params: Tuple = ()):
    cur = con.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    return rows


def _normalize_sequence(values: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    """Given list of (id, current_order_index) return list of (id, new_order_index).
    New order_index is 1..N preserving existing relative order by current_order_index then id.
    If already sequential (1..N) returns empty list (no changes).
    """
    if not values:
        return []
    # Sort by (current_order_index, id) for deterministic ordering
    ordered = sorted(values, key=lambda r: (r[1], r[0]))
    changed = []
    for new_idx, (rid, cur_idx) in enumerate(ordered, start=1):
        if cur_idx != new_idx:
            changed.append((rid, new_idx))
    return changed


# --- Normalizers --------------------------------------------------------------

def normalize_notebooks(con: sqlite3.Connection) -> Dict[str, List[Tuple[int, int]]]:
    rows = _fetchall(con, "SELECT id, order_index FROM notebooks ORDER BY order_index ASC, id ASC")
    changes = _normalize_sequence(rows)
    return {"notebooks": changes} if changes else {"notebooks": []}


def normalize_sections(con: sqlite3.Connection) -> Dict[str, List[Tuple[int, int]]]:
    notebook_ids = [r[0] for r in _fetchall(con, "SELECT id FROM notebooks ORDER BY id ASC")]
    all_changes: List[Tuple[int, int]] = []
    for nb_id in notebook_ids:
        rows = _fetchall(con, "SELECT id, order_index FROM sections WHERE notebook_id = ? ORDER BY order_index, id", (nb_id,))
        changes = _normalize_sequence(rows)
        all_changes.extend(changes)
    return {"sections": all_changes}


def normalize_pages(con: sqlite3.Connection) -> Dict[str, List[Tuple[int, int]]]:
    """Normalize pages order_index per sibling group.

    Backward compatible: if legacy databases lack parent_page_id column all pages in a section
    are treated as one group.
    """
    cur = con.cursor()
    cur.execute("PRAGMA table_info(pages)")
    cols = [r[1] for r in cur.fetchall()]
    has_parent = "parent_page_id" in cols

    all_changes: List[Tuple[int, int]] = []
    if has_parent:
        # Distinct groups (section_id, parent_page_id) - SQLite lacks NULLS FIRST keyword
        groups = _fetchall(con, "SELECT DISTINCT section_id, parent_page_id FROM pages ORDER BY section_id ASC, parent_page_id ASC")
        for section_id, parent_page_id in groups:
            if parent_page_id is None:
                rows = _fetchall(con, "SELECT id, order_index FROM pages WHERE section_id = ? AND parent_page_id IS NULL ORDER BY order_index, id", (section_id,))
            else:
                rows = _fetchall(con, "SELECT id, order_index FROM pages WHERE section_id = ? AND parent_page_id = ? ORDER BY order_index, id", (section_id, parent_page_id))
            changes = _normalize_sequence(rows)
            all_changes.extend(changes)
    else:
        # Legacy: group only by section
        sections = _fetchall(con, "SELECT DISTINCT section_id FROM pages ORDER BY section_id ASC")
        for (section_id,) in sections:
            rows = _fetchall(con, "SELECT id, order_index FROM pages WHERE section_id = ? ORDER BY order_index, id", (section_id,))
            changes = _normalize_sequence(rows)
            all_changes.extend(changes)
    return {"pages": all_changes}


# --- Apply --------------------------------------------------------------------

def apply_changes(con: sqlite3.Connection, changes: Dict[str, List[Tuple[int, int]]]):
    cur = con.cursor()
    # Notebooks
    for nid, new_order in changes.get("notebooks", []):
        cur.execute("UPDATE notebooks SET order_index = ? WHERE id = ?", (new_order, nid))
    # Sections
    for sid, new_order in changes.get("sections", []):
        cur.execute("UPDATE sections SET order_index = ? WHERE id = ?", (new_order, sid))
    # Pages
    for pid, new_order in changes.get("pages", []):
        cur.execute("UPDATE pages SET order_index = ? WHERE id = ?", (new_order, pid))
    con.commit()


# --- Reporting ----------------------------------------------------------------

def summarize(changes: Dict[str, List[Tuple[int, int]]]) -> str:
    lines = []
    for key in ("notebooks", "sections", "pages"):
        items = changes.get(key, [])
        if not items:
            lines.append(f"{key}: no changes")
        else:
            lines.append(f"{key}: {len(items)} updates")
    return "\n".join(lines)


def verbose_dump(con: sqlite3.Connection, changes: Dict[str, List[Tuple[int, int]]]):
    def dump(label: str, rows: List[Tuple[int, int]]):
        if not rows:
            print(f"{label}: (none)")
            return
        print(f"{label} changes (id -> new_order_index):")
        for rid, new_ord in rows:
            print(f"  {rid} -> {new_ord}")
    dump("Notebooks", changes.get("notebooks", []))
    dump("Sections", changes.get("sections", []))
    dump("Pages", changes.get("pages", []))


# --- Main ---------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(description="Normalize order_index values for notebooks, sections, and pages.")
    parser.add_argument("db_path", help="Path to SQLite database file")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--dry-run", dest="dry_run", action="store_true", help="Show planned changes only (default)")
    group.add_argument("--apply", dest="apply", action="store_true", help="Apply changes")
    parser.add_argument("--verbose", action="store_true", help="Show detailed per-id changes")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.db_path):
        print(f"Error: database '{args.db_path}' not found", file=sys.stderr)
        return 1

    con = sqlite3.connect(args.db_path)
    try:
        changes = {}
        changes.update(normalize_notebooks(con))
        changes.update(normalize_sections(con))
        changes.update(normalize_pages(con))
        if args.verbose:
            verbose_dump(con, changes)
        print(summarize(changes))
        total = sum(len(changes[k]) for k in changes)
        if total == 0:
            print("Already normalized.")
            return 0
        if args.apply:
            try:
                con.execute("BEGIN")
                apply_changes(con, changes)
                print("Applied normalization.")
            except Exception as e:
                con.execute("ROLLBACK")
                print(f"Failed to apply changes: {e}", file=sys.stderr)
                return 1
        else:
            print("Dry run; re-run with --apply to persist.")
        return 0
    finally:
        con.close()


if __name__ == "__main__":
    raise SystemExit(main())
