"""
media_store.py
Helpers for managing per-database media storage.
- Build media root path next to the DB (e.g., C:\\path\\notes.db.media\\)
- Save files by SHA-256 with fanout directories
- Create and query media/media_refs records
- Garbage collect unreferenced media

Note: UI wiring (normalizing HTML, hooks, etc.) will be added later.
"""

import hashlib
import mimetypes
import os
import shutil
import sqlite3
from typing import Optional, Tuple

FANOUT1 = 2
FANOUT2 = 2


def media_root_for_db(db_path: str) -> str:
    base = os.path.abspath(db_path)
    return base + ".media"


def build_rel_path(sha256_hex: str, ext: str) -> str:
    """Return a POSIX-style relative path for HTML/URLs (forward slashes).

    Filesystem operations use os.path.join(base, rel_path), which accepts '/'.
    """
    a = sha256_hex[:FANOUT1]
    b = sha256_hex[FANOUT1 : FANOUT1 + FANOUT2]
    return f"media/{a}/{b}/{sha256_hex}.{ext}"


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def sha256_file(src_path: str) -> str:
    h = hashlib.sha256()
    with open(src_path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def guess_mime_and_ext(src_path: str) -> Tuple[str, str]:
    mime, _ = mimetypes.guess_type(src_path)
    if not mime:
        mime = "application/octet-stream"
    ext = os.path.splitext(src_path)[1].lstrip(".").lower() or "bin"
    return mime, ext


def _conn(db_path: str):
    return sqlite3.connect(db_path)


def ensure_media_tables(db_path: str):
    """Ensure media and media_refs tables (and minimal indexes) exist.

    This is a safety net in case a database hasn't been migrated yet. Idempotent.
    """
    con = _conn(db_path)
    try:
        cur = con.cursor()
        cur.executescript(
            """
            CREATE TABLE IF NOT EXISTS media (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sha256 TEXT NOT NULL UNIQUE,
                mime_type TEXT NOT NULL,
                ext TEXT NOT NULL,
                original_filename TEXT,
                size_bytes INTEGER,
                created_at TEXT NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS media_refs (
                media_id INTEGER NOT NULL REFERENCES media(id) ON DELETE CASCADE,
                page_id INTEGER REFERENCES pages(id) ON DELETE CASCADE,
                section_id INTEGER REFERENCES sections(id) ON DELETE CASCADE,
                notebook_id INTEGER REFERENCES notebooks(id) ON DELETE CASCADE,
                role TEXT NOT NULL,
                CHECK (
                    (page_id IS NOT NULL AND section_id IS NULL AND notebook_id IS NULL) OR
                    (page_id IS NULL AND section_id IS NOT NULL AND notebook_id IS NULL) OR
                    (page_id IS NULL AND section_id IS NULL AND notebook_id IS NOT NULL)
                )
            );
            CREATE INDEX IF NOT EXISTS idx_media_refs_media ON media_refs(media_id);
            CREATE INDEX IF NOT EXISTS idx_media_refs_page ON media_refs(page_id);
            CREATE INDEX IF NOT EXISTS idx_media_refs_section ON media_refs(section_id);
            CREATE INDEX IF NOT EXISTS idx_media_refs_notebook ON media_refs(notebook_id);
            """
        )
        con.commit()
    finally:
        try:
            con.close()
        except Exception:
            pass


def upsert_media_record(
    db_path: str,
    sha256_hex: str,
    mime_type: str,
    ext: str,
    original_filename: Optional[str],
    size_bytes: int,
) -> int:
    # Ensure tables exist (in case migration hasn't run yet)
    try:
        ensure_media_tables(db_path)
    except Exception:
        pass
    con = _conn(db_path)
    try:
        cur = con.cursor()
        cur.execute("SELECT id FROM media WHERE sha256=?", (sha256_hex,))
        row = cur.fetchone()
        if row:
            return int(row[0])
        cur.execute(
            "INSERT INTO media(sha256, mime_type, ext, original_filename, size_bytes) VALUES (?,?,?,?,?)",
            (sha256_hex, mime_type, ext, original_filename or None, int(size_bytes)),
        )
        con.commit()
        return int(cur.lastrowid)
    finally:
        con.close()


def add_media_ref(
    db_path: str,
    media_id: int,
    *,
    page_id: int = None,
    section_id: int = None,
    notebook_id: int = None,
    role: str = "attachment",
):
    if sum(x is not None for x in (page_id, section_id, notebook_id)) != 1:
        raise ValueError("Exactly one of page_id, section_id, notebook_id must be provided")
    try:
        ensure_media_tables(db_path)
    except Exception:
        pass
    con = _conn(db_path)
    try:
        cur = con.cursor()
        cur.execute(
            "INSERT INTO media_refs(media_id, page_id, section_id, notebook_id, role) VALUES (?,?,?,?,?)",
            (int(media_id), page_id, section_id, notebook_id, role),
        )
        con.commit()
    finally:
        con.close()


def remove_media_ref(
    db_path: str,
    media_id: int,
    *,
    page_id: int = None,
    section_id: int = None,
    notebook_id: int = None,
):
    if sum(x is not None for x in (page_id, section_id, notebook_id)) != 1:
        raise ValueError("Exactly one of page_id, section_id, notebook_id must be provided")
    con = _conn(db_path)
    try:
        cur = con.cursor()
        cur.execute(
            "DELETE FROM media_refs WHERE media_id=? AND IFNULL(page_id,0)=IFNULL(?,0) AND IFNULL(section_id,0)=IFNULL(?,0) AND IFNULL(notebook_id,0)=IFNULL(?,0)",
            (int(media_id), page_id, section_id, notebook_id),
        )
        con.commit()
    finally:
        con.close()


def save_file_into_store(
    db_path: str, src_path: str, *, original_filename: Optional[str] = None
) -> Tuple[int, str]:
    """Copy a file into the DB's media store (content-addressed). Returns (media_id, relative_path)."""
    try:
        ensure_media_tables(db_path)
    except Exception:
        pass
    sha_hex = sha256_file(src_path)
    mime_type, ext = guess_mime_and_ext(src_path)
    rel_path = build_rel_path(sha_hex, ext)
    base = media_root_for_db(db_path)
    abs_path = os.path.join(base, rel_path)
    ensure_dir(os.path.dirname(abs_path))
    if not os.path.exists(abs_path):
        shutil.copy2(src_path, abs_path)
    size = os.path.getsize(abs_path)
    media_id = upsert_media_record(
        db_path, sha_hex, mime_type, ext, original_filename or os.path.basename(src_path), size
    )
    return media_id, rel_path


def resolve_media_path(db_path: str, rel_path: str) -> str:
    return os.path.join(media_root_for_db(db_path), rel_path)


def garbage_collect_unused_media(db_path: str) -> int:
    """Remove unreferenced media files from disk and DB. Returns count removed."""
    con = _conn(db_path)
    removed = 0
    try:
        cur = con.cursor()
        cur.execute(
            "SELECT id, sha256, ext FROM media WHERE id NOT IN (SELECT DISTINCT media_id FROM media_refs)"
        )
        rows = cur.fetchall()
        for mid, sha_hex, ext in rows:
            rel = build_rel_path(sha_hex, ext)
            abs_p = resolve_media_path(db_path, rel)
            try:
                if os.path.exists(abs_p):
                    os.remove(abs_p)
            except Exception:
                pass
            cur.execute("DELETE FROM media WHERE id=?", (int(mid),))
            removed += 1
        con.commit()
    finally:
        con.close()
    return removed
