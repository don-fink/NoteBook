"""
db_presets.py
Stores table and Planning Register presets in the database, so they travel with the notes
and are included in backups. A preset is the saved HTML of a table or register, exactly as
it was when saved, including any text and amounts in it.
"""

import pathlib
import sqlite3

KIND_TABLE = "table"
KIND_REGISTER = "register"

PRESETS_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS table_presets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL,
    name TEXT NOT NULL,
    html TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    modified_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (kind, name)
);
"""


def _looks_like_register(html: str) -> bool:
    """Guess whether preset HTML saved before kinds existed is a whole Planning Register.

    A register is a container table holding the register table and the cost list, so it
    has nested tables and the register's "Estimated Cost" header."""
    low = (html or "").lower()
    return low.count("<table") >= 2 and "estimated cost" in low


def _import_settings_presets(cur):
    """Copy presets kept in settings.json by earlier versions into a new presets table.

    settings.json is left as it is, so every database gets its own copy once."""
    try:
        from settings_manager import get_table_presets

        old = get_table_presets()
    except Exception:
        return
    for name, data in old.items():
        html = data.get("html") if isinstance(data, dict) else None
        if not isinstance(name, str) or not name.strip() or not isinstance(html, str):
            continue
        if not html.strip():
            continue
        kind = KIND_REGISTER if _looks_like_register(html) else KIND_TABLE
        cur.execute(
            "INSERT OR IGNORE INTO table_presets (kind, name, html) VALUES (?, ?, ?)",
            (kind, name.strip(), html),
        )


def ensure_presets_table(db_path: str):
    """Create the presets table if it's missing, importing any presets from settings.json."""
    con = sqlite3.connect(db_path)
    try:
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='table_presets'")
        if cur.fetchone() is None:
            cur.executescript(PRESETS_TABLE_SQL)
            _import_settings_presets(cur)
            con.commit()
    finally:
        con.close()


def list_preset_names(db_path: str, kind: str) -> list:
    ensure_presets_table(db_path)
    con = sqlite3.connect(db_path)
    try:
        cur = con.cursor()
        cur.execute(
            "SELECT name FROM table_presets WHERE kind = ? ORDER BY name COLLATE NOCASE", (kind,)
        )
        return [row[0] for row in cur.fetchall()]
    finally:
        con.close()


def get_preset_html(db_path: str, kind: str, name: str):
    ensure_presets_table(db_path)
    con = sqlite3.connect(db_path)
    try:
        cur = con.cursor()
        cur.execute("SELECT html FROM table_presets WHERE kind = ? AND name = ?", (kind, name))
        row = cur.fetchone()
        return row[0] if row else None
    finally:
        con.close()


def preset_exists(db_path: str, kind: str, name: str) -> bool:
    return get_preset_html(db_path, kind, name) is not None


def save_preset(db_path: str, kind: str, name: str, html: str):
    """Save a preset, replacing any preset of the same kind and name."""
    ensure_presets_table(db_path)
    con = sqlite3.connect(db_path)
    try:
        con.execute(
            "INSERT INTO table_presets (kind, name, html) VALUES (?, ?, ?) "
            "ON CONFLICT (kind, name) DO UPDATE SET html = excluded.html, "
            "modified_at = datetime('now')",
            (kind, name, html),
        )
        con.commit()
    finally:
        con.close()


def rename_preset(db_path: str, kind: str, old_name: str, new_name: str):
    """Rename a preset, replacing any other preset of the same kind already using new_name."""
    ensure_presets_table(db_path)
    con = sqlite3.connect(db_path)
    try:
        con.execute(
            "DELETE FROM table_presets WHERE kind = ? AND name = ? AND name <> ?",
            (kind, new_name, old_name),
        )
        con.execute(
            "UPDATE table_presets SET name = ?, modified_at = datetime('now') "
            "WHERE kind = ? AND name = ?",
            (new_name, kind, old_name),
        )
        con.commit()
    finally:
        con.close()


def delete_preset(db_path: str, kind: str, name: str):
    ensure_presets_table(db_path)
    con = sqlite3.connect(db_path)
    try:
        con.execute("DELETE FROM table_presets WHERE kind = ? AND name = ?", (kind, name))
        con.commit()
    finally:
        con.close()


def read_presets(db_path: str) -> list:
    """Return another database's presets as (kind, name, html) tuples, without changing it.

    The database is opened read-only. A database that has never stored presets has none."""
    uri = pathlib.Path(db_path).resolve().as_uri() + "?mode=ro"
    con = sqlite3.connect(uri, uri=True)
    try:
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='table_presets'")
        if cur.fetchone() is None:
            return []
        cur.execute("SELECT kind, name, html FROM table_presets ORDER BY kind, name COLLATE NOCASE")
        return [row for row in cur.fetchall() if row[0] in (KIND_TABLE, KIND_REGISTER)]
    finally:
        con.close()


def unique_preset_name(name: str, taken) -> str:
    """Return name, or "name (2)", "name (3)" and so on, whichever isn't in taken."""
    if name not in taken:
        return name
    n = 2
    while f"{name} ({n})" in taken:
        n += 1
    return f"{name} ({n})"
