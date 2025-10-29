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
