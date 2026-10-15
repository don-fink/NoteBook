"""
spreadsheet_store.py
Helpers for managing mutable spreadsheet files stored alongside other media.
"""

from __future__ import annotations

import csv
import importlib
import os
import re
import shutil
import sqlite3

from media_store import ensure_dir, media_root_for_db

SPREADSHEET_DIR = "media/spreadsheets"
SPREADSHEET_TEMPLATE_DIR = "media/spreadsheet_templates"

SPREADSHEET_TABLES_SQL = """
CREATE TABLE IF NOT EXISTS spreadsheets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    page_id INTEGER NOT NULL,
    title TEXT NOT NULL,
    rel_path TEXT NOT NULL UNIQUE,
    file_ext TEXT NOT NULL DEFAULT 'csv',
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    modified_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (page_id) REFERENCES pages(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_spreadsheets_page ON spreadsheets(page_id);
"""


def _conn(db_path: str):
    return sqlite3.connect(db_path)


def spreadsheet_root_for_db(db_path: str) -> str:
    return os.path.join(media_root_for_db(db_path), "media", "spreadsheets")


def spreadsheet_templates_root_for_db(db_path: str) -> str:
    return os.path.join(media_root_for_db(db_path), "media", "spreadsheet_templates")


def resolve_spreadsheet_path(db_path: str, rel_path: str) -> str:
    rel = (rel_path or "").replace("/", os.sep).replace("\\", os.sep)
    return os.path.join(media_root_for_db(db_path), rel)


def ensure_spreadsheet_support(db_path: str):
    con = _conn(db_path)
    try:
        cur = con.cursor()
        cur.executescript(SPREADSHEET_TABLES_SQL)
        con.commit()
    finally:
        con.close()
    ensure_dir(spreadsheet_root_for_db(db_path))
    ensure_dir(spreadsheet_templates_root_for_db(db_path))


def sanitize_spreadsheet_title(title: str) -> str:
    text = re.sub(r"\s+", " ", str(title or "Spreadsheet").strip())
    text = re.sub(r"[^A-Za-z0-9 _.-]", "", text)
    text = text.strip(" .")
    return text or "Spreadsheet"


def build_spreadsheet_rel_path(filename: str) -> str:
    return f"{SPREADSHEET_DIR}/{filename}".replace("\\", "/")


def ensure_unique_spreadsheet_rel_path(db_path: str, rel_path: str) -> str:
    rel = (rel_path or "").replace("\\", "/").lstrip("/")
    if not rel.startswith(f"{SPREADSHEET_DIR}/"):
        rel = build_spreadsheet_rel_path(os.path.basename(rel) or "Spreadsheet.csv")
    abs_path = resolve_spreadsheet_path(db_path, rel)
    if not os.path.exists(abs_path):
        return rel

    folder = os.path.dirname(rel).replace("\\", "/")
    filename = os.path.basename(rel)
    stem, ext = os.path.splitext(filename)
    counter = 2
    while True:
        candidate = f"{folder}/{stem}-{counter}{ext}"
        if not os.path.exists(resolve_spreadsheet_path(db_path, candidate)):
            return candidate
        counter += 1


def allocate_new_spreadsheet_rel_path(db_path: str, title: str, file_ext: str = "csv") -> str:
    ext = (file_ext or "csv").lstrip(".").lower() or "csv"
    safe_title = sanitize_spreadsheet_title(title)
    rel_path = build_spreadsheet_rel_path(f"{safe_title}.{ext}")
    return ensure_unique_spreadsheet_rel_path(db_path, rel_path)


def create_spreadsheet_record(
    db_path: str,
    page_id: int,
    title: str,
    rel_path: str,
    *,
    file_ext: str = "csv",
    created_at: str | None = None,
    modified_at: str | None = None,
) -> int:
    ensure_spreadsheet_support(db_path)
    con = _conn(db_path)
    try:
        cur = con.cursor()
        if created_at or modified_at:
            cur.execute(
                """
                INSERT INTO spreadsheets(page_id, title, rel_path, file_ext, created_at, modified_at)
                VALUES (?, ?, ?, ?, COALESCE(?, datetime('now')), COALESCE(?, datetime('now')))
                """,
                (
                    int(page_id),
                    sanitize_spreadsheet_title(title),
                    rel_path.replace("\\", "/"),
                    (file_ext or "csv").lstrip(".").lower() or "csv",
                    created_at,
                    modified_at,
                ),
            )
        else:
            cur.execute(
                "INSERT INTO spreadsheets(page_id, title, rel_path, file_ext) VALUES (?, ?, ?, ?)",
                (
                    int(page_id),
                    sanitize_spreadsheet_title(title),
                    rel_path.replace("\\", "/"),
                    (file_ext or "csv").lstrip(".").lower() or "csv",
                ),
            )
        con.commit()
        return int(cur.lastrowid)
    finally:
        con.close()


def get_spreadsheets_for_page(db_path: str, page_id: int) -> list[dict]:
    ensure_spreadsheet_support(db_path)
    con = _conn(db_path)
    try:
        cur = con.cursor()
        cur.execute(
            """
            SELECT id, page_id, title, rel_path, file_ext, created_at, modified_at
            FROM spreadsheets
            WHERE page_id = ?
            ORDER BY id
            """,
            (int(page_id),),
        )
        rows = cur.fetchall()
        return [
            {
                "id": int(row[0]),
                "page_id": int(row[1]),
                "title": row[2],
                "rel_path": row[3],
                "file_ext": row[4],
                "created_at": row[5],
                "modified_at": row[6],
            }
            for row in rows
        ]
    finally:
        con.close()


def create_blank_csv_spreadsheet(db_path: str, page_id: int, title: str) -> tuple[int, str, str]:
    ensure_spreadsheet_support(db_path)
    display_title = sanitize_spreadsheet_title(title)
    rel_path = allocate_new_spreadsheet_rel_path(db_path, display_title, "csv")
    abs_path = resolve_spreadsheet_path(db_path, rel_path)
    ensure_dir(os.path.dirname(abs_path))

    with open(abs_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["A", "B", "C", "D", "E"])

    spreadsheet_id = create_spreadsheet_record(
        db_path,
        int(page_id),
        display_title,
        rel_path,
        file_ext="csv",
    )
    return spreadsheet_id, rel_path, abs_path


def create_blank_xlsx_spreadsheet(db_path: str, page_id: int, title: str) -> tuple[int, str, str]:
    ensure_spreadsheet_support(db_path)
    display_title = sanitize_spreadsheet_title(title)
    rel_path = allocate_new_spreadsheet_rel_path(db_path, display_title, "xlsx")
    abs_path = resolve_spreadsheet_path(db_path, rel_path)
    ensure_dir(os.path.dirname(abs_path))

    openpyxl = importlib.import_module("openpyxl")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for col_index, header in enumerate(["A", "B", "C", "D", "E"], start=1):
        sheet.cell(row=1, column=col_index).value = header
    workbook.save(abs_path)

    spreadsheet_id = create_spreadsheet_record(
        db_path,
        int(page_id),
        display_title,
        rel_path,
        file_ext="xlsx",
    )
    return spreadsheet_id, rel_path, abs_path


def list_spreadsheet_templates(db_path: str) -> list[str]:
    ensure_spreadsheet_support(db_path)
    root = spreadsheet_templates_root_for_db(db_path)
    try:
        names = [name for name in os.listdir(root) if os.path.isfile(os.path.join(root, name))]
    except Exception:
        return []
    return sorted(names, key=str.casefold)


def create_spreadsheet_from_template(
    db_path: str,
    page_id: int,
    title: str,
    template_name: str,
) -> tuple[int, str, str]:
    ensure_spreadsheet_support(db_path)
    src_name = os.path.basename(template_name or "")
    src_path = os.path.join(spreadsheet_templates_root_for_db(db_path), src_name)
    if not os.path.isfile(src_path):
        raise FileNotFoundError(f"Template not found: {src_name}")

    ext = os.path.splitext(src_name)[1].lstrip(".").lower() or "csv"
    display_title = sanitize_spreadsheet_title(title or os.path.splitext(src_name)[0])
    rel_path = allocate_new_spreadsheet_rel_path(db_path, display_title, ext)
    abs_path = resolve_spreadsheet_path(db_path, rel_path)
    ensure_dir(os.path.dirname(abs_path))
    shutil.copy2(src_path, abs_path)

    spreadsheet_id = create_spreadsheet_record(
        db_path,
        int(page_id),
        display_title,
        rel_path,
        file_ext=ext,
    )
    return spreadsheet_id, rel_path, abs_path


def import_existing_spreadsheet(
    db_path: str,
    page_id: int,
    title: str,
    src_path: str,
) -> tuple[int, str, str]:
    ensure_spreadsheet_support(db_path)
    if not os.path.isfile(src_path):
        raise FileNotFoundError(src_path)

    ext = os.path.splitext(src_path)[1].lstrip(".").lower() or "csv"
    default_title = title or os.path.splitext(os.path.basename(src_path))[0]
    display_title = sanitize_spreadsheet_title(default_title)
    rel_path = allocate_new_spreadsheet_rel_path(db_path, display_title, ext)
    abs_path = resolve_spreadsheet_path(db_path, rel_path)
    ensure_dir(os.path.dirname(abs_path))
    shutil.copy2(src_path, abs_path)

    spreadsheet_id = create_spreadsheet_record(
        db_path,
        int(page_id),
        display_title,
        rel_path,
        file_ext=ext,
    )
    return spreadsheet_id, rel_path, abs_path
