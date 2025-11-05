"""
backup.py
Create timestamped backups of the current database, optionally including its media folder,
and enforce a keep-N retention policy.

Output format:
- A single .bundle file (ZIP format) named: <db_stem>-YYYYmmdd-HHMMSS.bundle
- Contents:
  - <db_filename> (the SQLite database file)
  - media/ ... (if media folder exists, full folder tree under 'media/')
"""

from __future__ import annotations

import os
import time
import zipfile
from typing import Optional
import json
import sqlite3


def _timestamp() -> str:
    return time.strftime("%Y%m%d-%H%M%S")


def _ensure_dir(path: str):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass


def _list_existing_bundles(dest_dir: str, stem: str) -> list[str]:
    try:
        items = []
        for name in os.listdir(dest_dir):
            if name.startswith(stem + "-") and name.endswith(".bundle"):
                items.append(os.path.join(dest_dir, name))
        return items
    except Exception:
        return []


def _retention_prune(dest_dir: str, stem: str, keep: int):
    if keep is None or keep <= 0:
        return
    bundles = _list_existing_bundles(dest_dir, stem)
    # Sort by modified time descending (newest first)
    try:
        bundles.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    except Exception:
        bundles.sort(reverse=True)
    # Delete all beyond keep
    for old in bundles[keep:]:
        try:
            os.remove(old)
        except Exception:
            pass

def _cleanup_stale_tmp_backups(dest_dir: str, *, min_age_seconds: int = 24 * 60 * 60):
    """Remove stale temporary backup/export files left from interrupted runs.

    We write bundles and binder exports to a temporary file (e.g., .bundle.tmp, .binder.tmp)
    and atomically rename at the end. If the process is interrupted before rename, the
    .tmp file can be left behind. It's safe to remove .tmp files older than a threshold.
    """
    try:
        now = time.time()
        for name in os.listdir(dest_dir or "."):
            lower = name.lower()
            if lower.endswith(".bundle.tmp") or lower.endswith(".binder.tmp"):
                p = os.path.join(dest_dir, name)
                try:
                    mtime = os.path.getmtime(p)
                    if (now - float(mtime)) >= float(min_age_seconds):
                        os.remove(p)
                except Exception:
                    pass
    except Exception:
        pass


def _zip_add_file(zf: zipfile.ZipFile, src_path: str, arcname: str):
    try:
        zf.write(src_path, arcname)
    except Exception:
        # Ignore missing or unreadable files
        pass


def _zip_add_dir(zf: zipfile.ZipFile, dir_path: str, arcbase: str):
    if not os.path.isdir(dir_path):
        return
    # Walk and add files; skip empty dirs
    for root, _dirs, files in os.walk(dir_path):
        for f in files:
            sp = os.path.join(root, f)
            try:
                rel = os.path.relpath(sp, dir_path)
            except Exception:
                rel = f
            arc = os.path.join(arcbase, rel).replace("\\", "/")
            _zip_add_file(zf, sp, arc)


def make_exit_backup(db_path: str, dest_dir: str, keep: int = 5, include_media: bool = True) -> Optional[str]:
    """Create a timestamped backup bundle for the given database.

    Returns the full path to the created .bundle, or None if it failed.
    """
    if not db_path or not os.path.isfile(db_path):
        return None
    if not dest_dir:
        return None
    try:
        _ensure_dir(dest_dir)
    except Exception:
        return None

    # Opportunistically clean up old temp files from prior interrupted backups/exports
    try:
        _cleanup_stale_tmp_backups(dest_dir, min_age_seconds=24 * 60 * 60)
    except Exception:
        pass

    base = os.path.basename(db_path)
    stem, _ext = os.path.splitext(base)
    ts = _timestamp()
    bundle_name = f"{stem}-{ts}.bundle"
    bundle_path = os.path.join(dest_dir, bundle_name)
    # Write to a temp file first, then atomically replace so partially-written bundles aren't visible
    tmp_path = bundle_path + ".tmp"

    # Choose compression if available
    try:
        compression = zipfile.ZIP_DEFLATED
    except Exception:
        compression = zipfile.ZIP_STORED

    # Ensure any stale temp file doesn't interfere
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass

    try:
        with zipfile.ZipFile(tmp_path, mode="w", compression=compression) as zf:
            # Add the database file at archive root using its filename
            _zip_add_file(zf, db_path, base)

            if include_media:
                try:
                    from media_store import media_root_for_db

                    media_dir = media_root_for_db(db_path)
                except Exception:
                    media_dir = None
                if media_dir and os.path.isdir(media_dir):
                    _zip_add_dir(zf, media_dir, "media")
        # Atomic-ish move into place (on Windows, this replaces if exists)
        try:
            os.replace(tmp_path, bundle_path)
        except Exception:
            # Fallback: try rename then remove tmp
            try:
                os.rename(tmp_path, bundle_path)
            except Exception:
                # Cleanup temp on failure
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
                return None
    except Exception:
        # Best effort: remove incomplete bundle
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        return None

    # Enforce retention
    try:
        _retention_prune(dest_dir, stem, int(keep) if keep is not None else 5)
    except Exception:
        pass

    return bundle_path


# -------------------- Manual operations UI helpers --------------------
def _sanitize_db_filename(name: str) -> Optional[str]:
    """Validate and normalize a Windows filename for the database.

    - Disallow path separators and reserved characters: \ / : * ? " < > |
    - Disallow names that end with space or dot
    - Ensure .db extension
    Returns normalized filename (not a path) or None if invalid.
    """
    if not name:
        return None
    name = name.strip()
    if not name:
        return None
    invalid = set('\\/:*?"<>|')
    if any(ch in invalid for ch in name):
        return None
    # No path components allowed
    if os.path.basename(name) != name:
        return None
    # Disallow trailing space or dot (Windows restriction)
    if name.endswith(" ") or name.endswith("."):
        return None
    # Ensure .db extension
    if not name.lower().endswith(".db"):
        name = name + ".db"
    return name


def _rename_database_and_media(db_path: str, new_filename: str) -> Optional[str]:
    """Rename the database file and its media folder to the new filename (same directory).

    Returns the new db path on success, or None on failure.
    """
    try:
        if not db_path or not os.path.isfile(db_path):
            return None
        directory = os.path.dirname(db_path)
        base_old = os.path.basename(db_path)
        # Normalize and validate filename
        new_filename = _sanitize_db_filename(new_filename or "")
        if not new_filename:
            return None
        if new_filename == base_old:
            # Nothing to do
            return db_path
        new_db_path = os.path.join(directory, new_filename)
        if os.path.exists(new_db_path):
            # Don't overwrite existing file
            return None

        # Determine media paths
        try:
            from media_store import media_root_for_db

            old_media = media_root_for_db(db_path)
            new_media = media_root_for_db(new_db_path)
        except Exception:
            old_media = None
            new_media = None

        # Rename database file first
        os.replace(db_path, new_db_path)

        # Rename media folder if present and paths differ
        try:
            if old_media and new_media and os.path.isdir(old_media):
                # If target exists, abort media rename but keep DB renamed
                if not os.path.exists(new_media):
                    os.replace(old_media, new_media)
        except Exception:
            # Non-fatal: DB already renamed; media rename failed
            pass
        return new_db_path
    except Exception:
        return None


def show_rename_database_dialog(window):
    """Show the Rename Database dialog and perform a safe rename of the DB and media folder.

    Fallbacks to a simple input dialog if the .ui form isn't available.
    """
    try:
        from PyQt5 import QtWidgets
    except Exception:
        return

    # Determine current db path and name
    try:
        from settings_manager import get_last_db, set_last_db

        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    except Exception:
        db_path = getattr(window, "_db_path", None) or "notes.db"
        set_last_db = None
    current_name = os.path.basename(db_path)

    # Try to load UI form; on failure, use QInputDialog
    dlg = None
    le = None
    btn_apply = None
    btn_cancel = None
    try:
        from ui_loader import load_dialog

        dlg = load_dialog("rename_database.ui", parent=window)
        if dlg is not None:
            # Support either actual name or spec'd name
            le = dlg.findChild(QtWidgets.QLineEdit, "lineEditDatabaseRename") or dlg.findChild(
                QtWidgets.QLineEdit, "lineEditDatabaseName"
            )
            btn_apply = dlg.findChild(QtWidgets.QPushButton, "pushButtonNewName")
            btn_cancel = dlg.findChild(QtWidgets.QPushButton, "pushButtonCancel")
    except Exception:
        dlg = None

    def _apply(new_text: str):
        norm = _sanitize_db_filename(new_text or "")
        if not norm:
            QtWidgets.QMessageBox.information(
                window, "Rename Database", "Please enter a valid filename (e.g., MyNotes.db)."
            )
            return
        # Compose new path in same directory
        new_path = os.path.join(os.path.dirname(db_path), norm)
        if os.path.exists(new_path) and os.path.normcase(new_path) != os.path.normcase(db_path):
            QtWidgets.QMessageBox.warning(
                window,
                "Rename Database",
                f"A file with that name already exists:\n{new_path}",
            )
            return
        # Try rename
        result = _rename_database_and_media(db_path, norm)
        if not result:
            QtWidgets.QMessageBox.warning(
                window, "Rename Database", "Could not rename the database."
            )
            return
        # Update app state: settings, window title, media root, editor baseUrl
        try:
            if set_last_db is not None:
                set_last_db(result)
            try:
                window._db_path = result
            except Exception:
                pass
            try:
                window.setWindowTitle(f"NoteBook — {result}")
            except Exception:
                pass
            # Update media root and editor base URL
            try:
                from media_store import media_root_for_db
                from PyQt5.QtCore import QUrl

                media_root = media_root_for_db(result)
                try:
                    window._media_root = media_root
                except Exception:
                    pass
                te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
                if te is not None and media_root:
                    if not media_root.endswith(os.sep) and not media_root.endswith("/"):
                        media_root = media_root + os.sep
                    te.document().setBaseUrl(QUrl.fromLocalFile(media_root))
            except Exception:
                pass
        except Exception:
            pass
        try:
            if dlg is not None:
                dlg.accept()
        except Exception:
            pass
        QtWidgets.QMessageBox.information(
            window, "Rename Database", f"Renamed to:\n{result}"
        )

    if dlg is not None:
        try:
            if le is not None:
                le.setText(current_name)
                le.selectAll()
            if btn_apply is not None:
                btn_apply.clicked.connect(lambda: _apply(le.text() if le is not None else current_name))
            # Allow Enter to confirm when focus is in the line edit
            try:
                if le is not None:
                    le.returnPressed.connect(lambda: _apply(le.text()))
            except Exception:
                pass
            if btn_cancel is not None:
                btn_cancel.clicked.connect(lambda: dlg.reject())
            dlg.setWindowTitle("Rename Database")
            dlg.exec_()
            return
        except Exception:
            pass

    # Fallback simple input dialog
    new_name, ok = QtWidgets.QInputDialog.getText(
        window, "Rename Database", "New filename:", text=current_name
    )
    if not ok:
        return
    _apply(new_name)


# -------------------- Export Binder --------------------
def _current_notebook_id_from_ui(window) -> Optional[int]:
    try:
        # Prefer explicit context tracked on the window
        nb_id = getattr(window, "_current_notebook_id", None)
        if nb_id is not None:
            try:
                return int(nb_id)
            except Exception:
                return nb_id
        # Fallback: look at the left tree current selection
        from PyQt5 import QtWidgets as _QtW

        tree = window.findChild(_QtW.QTreeWidget, "notebookName")
        if tree is not None:
            item = tree.currentItem()
            if item is not None:
                if item.parent() is None:
                    val = item.data(0, 1000)
                    try:
                        return int(val)
                    except Exception:
                        return val
                # If a section/page is selected, use its parent binder
                par = item.parent()
                if par is not None:
                    val = par.data(0, 1000)
                    try:
                        return int(val)
                    except Exception:
                        return val
    except Exception:
        pass
    return None


def _fetch_binder_payload(db_path: str, notebook_id: int) -> Optional[dict]:
    """Collect binder, sections, pages and referenced media + refs for export payload."""
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    try:
        cur = con.cursor()
        cur.execute(
            "SELECT id, title, created_at, modified_at, order_index FROM notebooks WHERE id=?",
            (int(notebook_id),),
        )
        nb = cur.fetchone()
        if not nb:
            return None
        cur.execute(
            "SELECT id, title, color_hex, created_at, modified_at, order_index FROM sections WHERE notebook_id=? ORDER BY order_index, id",
            (int(notebook_id),),
        )
        sections = cur.fetchall()
        sec_ids = [int(r["id"]) for r in sections]
        pages_by_sec = {sid: [] for sid in sec_ids}
        pages = []
        if sec_ids:
            q = (
                "SELECT id, section_id, title, content_html, created_at, modified_at, order_index "
                "FROM pages WHERE section_id IN (" + ",".join(["?"] * len(sec_ids)) + ") "
                "ORDER BY order_index, id"
            )
            cur.execute(q, sec_ids)
            pages = cur.fetchall()
            for p in pages:
                sid = int(p["section_id"])
                if sid in pages_by_sec:
                    pages_by_sec[sid].append(p)

        # Collect referenced media ids
        media_ids = set()
        if sec_ids or pages:
            params = [int(notebook_id)] + sec_ids + [int(r["id"]) for r in pages]
            q = (
                "SELECT DISTINCT media_id FROM media_refs WHERE "
                "notebook_id=? OR section_id IN (" + (",".join(["?"] * len(sec_ids)) if sec_ids else "NULL") + ") "
                "OR page_id IN (" + (",".join(["?"] * len(pages)) if pages else "NULL") + ")"
            )
            try:
                cur.execute(q, params)
                for (mid,) in cur.fetchall():
                    try:
                        media_ids.add(int(mid))
                    except Exception:
                        pass
            except Exception:
                pass
        media_rows = []
        if media_ids:
            q = (
                "SELECT id, sha256, mime_type, ext, original_filename, size_bytes, created_at "
                "FROM media WHERE id IN (" + ",".join(["?"] * len(media_ids)) + ")"
            )
            cur.execute(q, list(media_ids))
            media_rows = cur.fetchall()

        # Collect explicit media refs for pages/sections/notebook
        refs = []
        try:
            # Notebook-level refs
            cur.execute(
                "SELECT media_id, role FROM media_refs WHERE notebook_id=?",
                (int(notebook_id),),
            )
            for mid, role in cur.fetchall():
                refs.append({"media_orig_id": int(mid), "notebook_orig_id": int(notebook_id), "role": role})
            # Section-level refs
            if sec_ids:
                q = (
                    "SELECT media_id, section_id, role FROM media_refs WHERE section_id IN ("
                    + ",".join(["?"] * len(sec_ids))
                    + ")"
                )
                cur.execute(q, sec_ids)
                for mid, sid, role in cur.fetchall():
                    refs.append({"media_orig_id": int(mid), "section_orig_id": int(sid), "role": role})
            # Page-level refs
            if pages:
                page_ids = [int(r["id"]) for r in pages]
                q = (
                    "SELECT media_id, page_id, role FROM media_refs WHERE page_id IN ("
                    + ",".join(["?"] * len(page_ids))
                    + ")"
                )
                cur.execute(q, page_ids)
                for mid, pid, role in cur.fetchall():
                    refs.append({"media_orig_id": int(mid), "page_orig_id": int(pid), "role": role})
        except Exception:
            refs = []

        # Optional: source DB uuid
        db_uuid = None
        try:
            cur.execute("SELECT uuid FROM db_metadata WHERE id=1")
            row = cur.fetchone()
            db_uuid = row[0] if row and row[0] else None
        except Exception:
            pass

        # Shape payload with nested sections/pages, include original IDs for import mapping
        out = {
            "format": "notebook-binder-v1",
            "exported_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "source_db_uuid": db_uuid,
            "notebook": {
                "orig_id": int(nb["id"]),
                "title": nb["title"],
                "order_index": int(nb["order_index"] or 0),
                "created_at": nb["created_at"],
                "modified_at": nb["modified_at"],
            },
            "sections": [],
            "media": [],
            "refs": refs,
        }
        for s in sections:
            sec = {
                "orig_id": int(s["id"]),
                "title": s["title"],
                "color_hex": s["color_hex"],
                "order_index": int(s["order_index"] or 0),
                "created_at": s["created_at"],
                "modified_at": s["modified_at"],
                "pages": [],
            }
            for p in pages_by_sec.get(int(s["id"]), []):
                sec["pages"].append(
                    {
                        "orig_id": int(p["id"]),
                        "title": p["title"],
                        "content_html": p["content_html"] or "",
                        "order_index": int(p["order_index"] or 0),
                        "created_at": p["created_at"],
                        "modified_at": p["modified_at"],
                    }
                )
            out["sections"].append(sec)
        for m in media_rows:
            out["media"].append(
                {
                    "orig_id": int(m["id"]),
                    "sha256": m["sha256"],
                    "mime_type": m["mime_type"],
                    "ext": m["ext"],
                    "original_filename": m["original_filename"],
                    "size_bytes": int(m["size_bytes"] or 0),
                    "created_at": m["created_at"],
                }
            )
        return out
    finally:
        try:
            con.close()
        except Exception:
            pass


def export_binder(window):
    """Export the currently selected binder (notebook) and its media into a .binder file (ZIP)."""
    try:
        from PyQt5 import QtWidgets
        from PyQt5.QtCore import QUrl, Qt
    except Exception:
        return

    # Save any unsaved edits first
    try:
        from page_editor import save_current_page

        save_current_page(window)
    except Exception:
        pass

    # Identify db path and selected binder id
    try:
        from settings_manager import get_last_db

        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    except Exception:
        db_path = getattr(window, "_db_path", None) or "notes.db"
    nb_id = _current_notebook_id_from_ui(window)
    if nb_id is None:
        QtWidgets.QMessageBox.information(
            window, "Export Binder", "Please select a binder to export."
        )
        return

    # Fetch payload and determine default filename
    payload = _fetch_binder_payload(db_path, int(nb_id))
    if not payload:
        QtWidgets.QMessageBox.warning(window, "Export Binder", "Could not read binder data.")
        return
    binder_title = payload["notebook"]["title"] or "Binder"
    # Sanitize title for filename
    safe = "".join(ch for ch in binder_title if ch not in '\\/:*?"<>|' ).strip() or "Binder"
    ts = _timestamp()
    default_name = f"{safe}-{ts}.binder"

    # Choose destination path
    initial_dir = os.path.dirname(db_path)
    out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
        window,
        "Export Binder",
        os.path.join(initial_dir, default_name),
        "Binder Export (*.binder);;All Files (*)",
    )
    if not out_path:
        return
    if not out_path.lower().endswith(".binder"):
        out_path = out_path + ".binder"

    # Build ZIP to a temp file then atomically replace
    tmp_path = out_path + ".tmp"
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass

    # Write manifest and media files
    QtWidgets.QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        # Prefer deflated compression
        try:
            compression = zipfile.ZIP_DEFLATED
        except Exception:
            compression = zipfile.ZIP_STORED
        with zipfile.ZipFile(tmp_path, mode="w", compression=compression) as zf:
            # Manifest
            manifest = json.dumps(payload, ensure_ascii=False, indent=2)
            zf.writestr("binder.json", manifest)
            # Media files under media/
            try:
                from media_store import build_rel_path, resolve_media_path

                for m in payload.get("media", []):
                    try:
                        rel = build_rel_path(m["sha256"], m["ext"]).replace("\\", "/")
                        abs_p = resolve_media_path(db_path, rel)
                        if os.path.isfile(abs_p):
                            zf.write(abs_p, rel)
                    except Exception:
                        pass
            except Exception:
                pass
        # Move into place
        try:
            os.replace(tmp_path, out_path)
        except Exception:
            os.rename(tmp_path, out_path)
    except Exception as e:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        try:
            QtWidgets.QApplication.restoreOverrideCursor()
        except Exception:
            pass
        QtWidgets.QMessageBox.warning(
            window, "Export Binder", f"Failed to export binder:\n{e}"
        )
        return
    finally:
        try:
            QtWidgets.QApplication.restoreOverrideCursor()
        except Exception:
            pass

    QtWidgets.QMessageBox.information(
        window, "Export Binder", f"Binder exported to:\n{out_path}"
    )


# -------------------- Import Binder --------------------
def _unique_notebook_title(db_path: str, title: str) -> str:
    try:
        con = sqlite3.connect(db_path)
        cur = con.cursor()
        cur.execute("SELECT title FROM notebooks")
        existing = {r[0] for r in cur.fetchall() if r and r[0]}
        con.close()
    except Exception:
        existing = set()
    base = (title or "Binder").strip() or "Binder"
    if base not in existing:
        return base
    i = 2
    while True:
        cand = f"{base} ({i})"
        if cand not in existing:
            return cand
        i += 1


def import_binder(window):
    """Import a .binder file: create a new binder with sections, pages, media, and refs.

    Title collisions are resolved by auto-suffixing (Title (2), Title (3), ...).
    """
    try:
        from PyQt5 import QtWidgets
        from PyQt5.QtCore import Qt
    except Exception:
        return

    # Choose binder file
    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
        window,
        "Import Binder",
        "",
        "Binder Export (*.binder);;All Files (*)",
    )
    if not file_path:
        return

    # Resolve target db
    try:
        from settings_manager import get_last_db, set_last_state

        db_path = getattr(window, "_db_path", None) or get_last_db() or "notes.db"
    except Exception:
        db_path = getattr(window, "_db_path", None) or "notes.db"
        set_last_state = None

    # Read manifest
    try:
        # Show busy cursor during import
        try:
            QtWidgets.QApplication.setOverrideCursor(Qt.WaitCursor)
        except Exception:
            pass
        with zipfile.ZipFile(file_path, "r") as zf:
            try:
                data = zf.read("binder.json")
            except KeyError:
                QtWidgets.QMessageBox.warning(window, "Import Binder", "Invalid binder: missing binder.json")
                return
            try:
                manifest = json.loads(data.decode("utf-8"))
            except Exception as e:
                QtWidgets.QMessageBox.warning(window, "Import Binder", f"Invalid binder.json: {e}")
                return
            if not isinstance(manifest, dict) or manifest.get("format") != "notebook-binder-v1":
                QtWidgets.QMessageBox.warning(window, "Import Binder", "Unsupported binder format.")
                return

            # Prepare DB inserts (single connection, with a small busy timeout)
            con = sqlite3.connect(db_path, timeout=10.0)
            try:
                con.execute("PRAGMA foreign_keys=ON")
            except Exception:
                pass
            cur = con.cursor()

            # Ensure media tables exist using the SAME connection to avoid cross-connection locks
            try:
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
            except Exception:
                pass

            # Create notebook with unique title
            src_nb = manifest.get("notebook", {})
            title = _unique_notebook_title(db_path, src_nb.get("title") or "Binder")
            nb_created = src_nb.get("created_at")
            nb_modified = src_nb.get("modified_at")
            nb_order = int(src_nb.get("order_index") or 0)
            try:
                cur.execute(
                    "INSERT INTO notebooks(title, created_at, modified_at, order_index) VALUES (?,?,?,?)",
                    (title, nb_created or "datetime('now')", nb_modified or "datetime('now')", nb_order),
                )
            except Exception:
                # Fallback minimal insert
                cur.execute("INSERT INTO notebooks(title, order_index) VALUES (?,?)", (title, nb_order))
            nb_new_id = cur.lastrowid

            # Map sections
            sec_map = {}
            for s in manifest.get("sections", []):
                s_title = s.get("title") or "Untitled Section"
                s_color = s.get("color_hex")
                s_order = int(s.get("order_index") or 0)
                s_created = s.get("created_at")
                s_modified = s.get("modified_at")
                try:
                    cur.execute(
                        "INSERT INTO sections(notebook_id, title, color_hex, created_at, modified_at, order_index) VALUES (?,?,?,?,?,?)",
                        (nb_new_id, s_title, s_color, s_created or "datetime('now')", s_modified or "datetime('now')", s_order),
                    )
                except Exception:
                    cur.execute(
                        "INSERT INTO sections(notebook_id, title, color_hex, order_index) VALUES (?,?,?,?)",
                        (nb_new_id, s_title, s_color, s_order),
                    )
                sec_map[int(s.get("orig_id"))] = cur.lastrowid

            # Map pages
            page_map = {}
            for s in manifest.get("sections", []):
                new_sid = sec_map.get(int(s.get("orig_id")))
                if not new_sid:
                    continue
                for p in s.get("pages", []):
                    p_title = p.get("title") or "Untitled Page"
                    p_html = p.get("content_html") or ""
                    p_order = int(p.get("order_index") or 0)
                    p_created = p.get("created_at")
                    p_modified = p.get("modified_at")
                    try:
                        cur.execute(
                            "INSERT INTO pages(section_id, title, content_html, created_at, modified_at, order_index) VALUES (?,?,?,?,?,?)",
                            (new_sid, p_title, p_html, p_created or "datetime('now')", p_modified or "datetime('now')", p_order),
                        )
                    except Exception:
                        cur.execute(
                            "INSERT INTO pages(section_id, title, content_html, order_index) VALUES (?,?,?,?)",
                            (new_sid, p_title, p_html, p_order),
                        )
                    page_map[int(p.get("orig_id"))] = cur.lastrowid

            # Import media files and upsert media records
            from media_store import media_root_for_db, ensure_dir, build_rel_path

            base_media = media_root_for_db(db_path)
            ensure_dir(base_media)

            # Map sha -> media_id in current DB
            sha_to_id = {}

            for m in manifest.get("media", []):
                sha = m.get("sha256")
                ext = m.get("ext") or "bin"
                if not sha:
                    continue
                rel = build_rel_path(sha, ext)
                abs_p = os.path.join(base_media, rel)
                # Ensure directory exists
                ensure_dir(os.path.dirname(abs_p))
                # Extract file if missing
                try:
                    if not os.path.exists(abs_p):
                        with zf.open(rel) as src, open(abs_p, "wb") as dst:
                            # Stream copy to avoid large allocations
                            for chunk in iter(lambda: src.read(1024 * 1024), b""):
                                dst.write(chunk)
                except KeyError:
                    # Media file missing from bundle; continue but still insert record
                    pass
                size = os.path.getsize(abs_p) if os.path.exists(abs_p) else int(m.get("size_bytes") or 0)
                # Upsert into media table using the SAME connection
                cur.execute("SELECT id FROM media WHERE sha256=?", (sha,))
                row = cur.fetchone()
                if row:
                    media_id = int(row[0])
                else:
                    cur.execute(
                        "INSERT INTO media(sha256, mime_type, ext, original_filename, size_bytes) VALUES (?,?,?,?,?)",
                        (
                            sha,
                            m.get("mime_type") or "application/octet-stream",
                            ext,
                            m.get("original_filename"),
                            int(size or 0),
                        ),
                    )
                    media_id = int(cur.lastrowid)
                sha_to_id[sha] = media_id

            # Recreate refs (fallback to content scan if absent)
            refs = manifest.get("refs") or []
            if refs:
                for r in refs:
                    mid = r.get("media_orig_id")
                    sha = None
                    # Find sha for media_orig_id
                    for m in manifest.get("media", []):
                        if int(m.get("orig_id")) == int(mid):
                            sha = m.get("sha256")
                            break
                    if not sha or sha not in sha_to_id:
                        continue
                    new_mid = sha_to_id[sha]
                    role = r.get("role") or "attachment"
                    if r.get("page_orig_id") is not None:
                        pid = page_map.get(int(r.get("page_orig_id")))
                        if pid:
                            cur.execute(
                                "INSERT INTO media_refs(media_id, page_id, section_id, notebook_id, role) VALUES (?,?,?,?,?)",
                                (int(new_mid), int(pid), None, None, role),
                            )
                    elif r.get("section_orig_id") is not None:
                        sid = sec_map.get(int(r.get("section_orig_id")))
                        if sid:
                            cur.execute(
                                "INSERT INTO media_refs(media_id, page_id, section_id, notebook_id, role) VALUES (?,?,?,?,?)",
                                (int(new_mid), None, int(sid), None, role),
                            )
                    elif r.get("notebook_orig_id") is not None:
                        cur.execute(
                            "INSERT INTO media_refs(media_id, page_id, section_id, notebook_id, role) VALUES (?,?,?,?,?)",
                            (int(new_mid), None, None, int(nb_new_id), role),
                        )
            else:
                # Fallback: scan page HTML for media paths and add page-level refs
                import re
                pat = re.compile(r"media/[0-9a-f]{2}/[0-9a-f]{2}/([0-9a-f]{64})\\.[A-Za-z0-9]+")
                for s in manifest.get("sections", []):
                    new_sid = sec_map.get(int(s.get("orig_id")))
                    for p in s.get("pages", []):
                        new_pid = page_map.get(int(p.get("orig_id")))
                        if not new_pid:
                            continue
                        html = p.get("content_html") or ""
                        for match in pat.finditer(html):
                            sha = match.group(1)
                            new_mid = sha_to_id.get(sha)
                            if new_mid:
                                try:
                                    cur.execute(
                                        "INSERT INTO media_refs(media_id, page_id, section_id, notebook_id, role) VALUES (?,?,?,?,?)",
                                        (int(new_mid), int(new_pid), None, None, "inline"),
                                    )
                                except Exception:
                                    pass

            con.commit()
            con.close()

            # UI: repopulate binders, select the new one, and refresh its sections
            try:
                from left_tree import refresh_for_notebook, ensure_left_tree_sections
                from ui_logic import populate_notebook_names
                from settings_manager import set_last_state

                # Persist state to the new notebook id
                set_last_state(notebook_id=int(nb_new_id), section_id=None, page_id=None)
                try:
                    window._current_notebook_id = int(nb_new_id)
                except Exception:
                    pass
                # Rebuild binder list and select new binder
                try:
                    populate_notebook_names(window, db_path)
                except Exception:
                    pass
                # Now refresh for the new binder id
                refresh_for_notebook(window, int(nb_new_id))
                ensure_left_tree_sections(window, int(nb_new_id))
            except Exception:
                pass

            # Done: restore cursor and notify
            try:
                QtWidgets.QApplication.restoreOverrideCursor()
            except Exception:
                pass
            QtWidgets.QMessageBox.information(
                window, "Import Binder", f"Binder imported as: {title}"
            )
    except zipfile.BadZipFile as e:
        QtWidgets.QMessageBox.warning(window, "Import Binder", f"Invalid binder file: {e}")
        return
    finally:
        # Ensure cursor is restored if any path above early-returns without clearing
        try:
            QtWidgets.QApplication.restoreOverrideCursor()
        except Exception:
            pass
