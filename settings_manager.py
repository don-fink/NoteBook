"""
settings_manager.py
Manages loading and saving application settings, such as the last opened database, in a JSON file.
"""

import json
import os
import shutil
import sys

# --- Settings file location strategy ---
# We now store settings in a per-user configuration directory instead of the
# application working directory to avoid permission issues when the app is
# installed under Program Files (read-only for standard users) and to keep
# user state separate from the program files.
#
# Windows: %LOCALAPPDATA%/NoteBook/settings.json
# Other platforms:
#   - macOS: ~/Library/Application Support/NoteBook/settings.json
#   - Linux/other: ~/.config/NoteBook/settings.json
#
# Backwards compatibility: if an old ./settings.json exists alongside the
# application and the new per-user settings file does not yet exist, we will
# transparently migrate (move) the legacy file to the new location.

_LEGACY_SETTINGS_FILE = "settings.json"  # in CWD / app directory
_SETTINGS_BASENAME = "settings.json"
_POINTER_BASENAME = "settings.loc"  # stores absolute path to settings.json (override)
_CACHED_SETTINGS_PATH = None  # memoize resolved path


def _default_settings_dir() -> str:
    """Return the platform-specific default settings directory (no overrides)."""
    try:
        if os.name == "nt":  # Windows
            base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
            return os.path.join(base, "NoteBook")
        elif sys.platform == "darwin":  # macOS
            return os.path.join(os.path.expanduser("~"), "Library", "Application Support", "NoteBook")
        else:  # Linux / other Unix
            return os.path.join(os.path.expanduser("~"), ".config", "NoteBook")
    except Exception:
        try:
            return os.path.abspath(os.getcwd())
        except Exception:
            return "."


def _pointer_file_path() -> str:
    return os.path.join(_default_settings_dir(), _POINTER_BASENAME)


def _read_settings_pointer() -> str:
    """Return absolute path to settings.json from pointer file if present, else None."""
    try:
        p = _pointer_file_path()
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                line = f.readline().strip()
                if line:
                    return line
    except Exception:
        pass
    return None


def _write_settings_pointer(settings_full_path: str):
    """Persist absolute path to settings.json in pointer file under default dir."""
    try:
        d = _default_settings_dir()
        os.makedirs(d, exist_ok=True)
        with open(_pointer_file_path(), "w", encoding="utf-8") as f:
            f.write(os.path.abspath(settings_full_path))
    except Exception:
        pass


def get_settings_dir() -> str:
    """Return the active settings directory (override if configured), creating it if needed."""
    # If we have an override pointer, return its directory
    override_path = _read_settings_pointer()
    if override_path:
        try:
            d = os.path.dirname(override_path)
            os.makedirs(d, exist_ok=True)
            return d
        except Exception:
            pass
    # Otherwise use default
    d = _default_settings_dir()
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def _resolve_settings_path() -> str:
    """Compute the new settings file path, migrating legacy file if present."""
    global _CACHED_SETTINGS_PATH
    if _CACHED_SETTINGS_PATH:
        return _CACHED_SETTINGS_PATH
    # If an override pointer exists, honor it first
    override_path = _read_settings_pointer()
    if override_path:
        # Ensure parent directory exists
        try:
            os.makedirs(os.path.dirname(override_path), exist_ok=True)
        except Exception:
            pass
        new_path = os.path.abspath(override_path)
    else:
        target_dir = get_settings_dir()
        new_path = os.path.join(target_dir, _SETTINGS_BASENAME)
    # Migration: if legacy exists in CWD and new_path missing, move it
    try:
        if not os.path.exists(new_path) and os.path.exists(_LEGACY_SETTINGS_FILE):
            # Attempt move; if move fails (cross-device), copy instead
            try:
                shutil.move(_LEGACY_SETTINGS_FILE, new_path)
            except Exception:
                try:
                    shutil.copy2(_LEGACY_SETTINGS_FILE, new_path)
                except Exception:
                    pass
    except Exception:
        pass
    _CACHED_SETTINGS_PATH = new_path
    return new_path


def get_settings_file_path() -> str:
    """Return absolute path to the current settings.json file."""
    return os.path.abspath(_resolve_settings_path())


def load_settings():
    path = _resolve_settings_path()
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_settings(settings):
    path = _resolve_settings_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


def set_settings_file_path(full_path: str):
    """Persistently switch settings.json location to the given absolute path.

    This writes a pointer file under the default settings directory so the new
    location is honored across restarts. Also updates the in-memory cache.
    """
    if not isinstance(full_path, str) or not full_path:
        return
    try:
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        _write_settings_pointer(os.path.abspath(full_path))
        global _CACHED_SETTINGS_PATH
        _CACHED_SETTINGS_PATH = os.path.abspath(full_path)
    except Exception:
        pass



def get_last_db():
    settings = load_settings()
    return settings.get("last_db")


def set_last_db(db_path):
    settings = load_settings()
    settings["last_db"] = db_path
    save_settings(settings)


# --- Databases root folder ---
def get_databases_root() -> str:
    """Folder under which .db files and their media folders are stored. Defaults to working directory."""
    s = load_settings()
    val = s.get("databases_root")
    if val and isinstance(val, str) and os.path.isdir(val):
        return val
    # Fallback: current working directory
    try:
        return os.path.abspath(os.getcwd())
    except Exception:
        return os.path.abspath(".")


def set_databases_root(path: str):
    if not isinstance(path, str) or not path:
        return
    s = load_settings()
    s["databases_root"] = path
    save_settings(s)


# --- Last position/state helpers ---
def get_last_state():
    """Return a dict with keys: last_notebook_id, last_section_id, last_page_id."""
    s = load_settings()
    return {
        "last_notebook_id": s.get("last_notebook_id"),
        "last_section_id": s.get("last_section_id"),
        "last_page_id": s.get("last_page_id"),
    }


def set_last_state(notebook_id=None, section_id=None, page_id=None):
    s = load_settings()
    if notebook_id is not None:
        s["last_notebook_id"] = notebook_id
    if section_id is not None:
        s["last_section_id"] = section_id
    if page_id is not None:
        s["last_page_id"] = page_id
    save_settings(s)


def clear_last_state():
    """Clear last selected notebook/section/page from settings."""
    s = load_settings()
    for k in ("last_notebook_id", "last_section_id", "last_page_id"):
        if k in s:
            del s[k]
    save_settings(s)


def get_window_geometry():
    s = load_settings()
    return s.get("window_geometry")  # dict with x, y, w, h


def set_window_geometry(x, y, w, h):
    s = load_settings()
    s["window_geometry"] = {"x": int(x), "y": int(y), "w": int(w), "h": int(h)}
    save_settings(s)


def get_window_maximized():
    s = load_settings()
    return bool(s.get("window_maximized", False))


def set_window_maximized(is_maximized: bool):
    s = load_settings()
    s["window_maximized"] = bool(is_maximized)
    save_settings(s)


# --- Splitter sizes (main horizontal splitter) ---
def get_splitter_sizes():
    """Return a list of ints representing QSplitter sizes, or None if not set."""
    s = load_settings()
    sizes = s.get("splitter_sizes")
    if isinstance(sizes, list) and all(isinstance(x, (int, float)) for x in sizes):
        # Coerce to ints
        return [int(x) for x in sizes]
    return None


def set_splitter_sizes(sizes):
    """Persist splitter sizes list (ints)."""
    try:
        cleaned = [int(x) for x in (sizes or [])]
    except Exception:
        cleaned = None
    s = load_settings()
    if cleaned:
        s["splitter_sizes"] = cleaned
    else:
        # Remove if invalid/empty
        if "splitter_sizes" in s:
            del s["splitter_sizes"]
    save_settings(s)


# --- Optional: per-section color mapping (for colored tabs and right-pane icons) ---
def get_section_colors():
    """Return a dict mapping section_id (as string) -> color hex string (e.g., '#FF8800')."""
    s = load_settings()
    return s.get("section_colors", {})


def set_section_color(section_id: int, color_hex: str):
    """Persist a color hex string for a section id."""
    s = load_settings()
    colors = s.get("section_colors", {})
    colors[str(int(section_id))] = str(color_hex)
    s["section_colors"] = colors
    save_settings(s)


# --- Left tree expanded state ---
def get_expanded_notebooks():
    """Return a set of notebook IDs that should be expanded in the left tree."""
    s = load_settings()
    vals = s.get("expanded_notebooks", [])
    try:
        return set(int(v) for v in vals)
    except Exception:
        return set()


def set_expanded_notebooks(ids):
    """Persist the full set of expanded notebook IDs."""
    s = load_settings()
    s["expanded_notebooks"] = [int(v) for v in ids]
    save_settings(s)


def add_expanded_notebook(notebook_id: int):
    ids = get_expanded_notebooks()
    ids.add(int(notebook_id))
    set_expanded_notebooks(ids)


def remove_expanded_notebook(notebook_id: int):
    ids = get_expanded_notebooks()
    if int(notebook_id) in ids:
        ids.remove(int(notebook_id))
    set_expanded_notebooks(ids)


# --- Right tree expanded sections (per notebook) ---
def get_expanded_sections_by_notebook():
    """Return a dict mapping notebook_id (str) -> set of expanded section IDs (ints)."""
    s = load_settings()
    raw = s.get("expanded_sections_by_notebook", {})
    out = {}
    try:
        for k, vals in raw.items():
            try:
                out[str(int(k))] = set(int(v) for v in vals)
            except Exception:
                out[str(k)] = set()
    except Exception:
        pass
    return out


def set_expanded_sections_for_notebook(notebook_id: int, section_ids):
    s = load_settings()
    raw = s.get("expanded_sections_by_notebook", {})
    raw[str(int(notebook_id))] = [int(v) for v in section_ids]
    s["expanded_sections_by_notebook"] = raw
    save_settings(s)


def add_expanded_section(notebook_id: int, section_id: int):
    m = get_expanded_sections_by_notebook()
    key = str(int(notebook_id))
    cur = m.get(key, set())
    cur.add(int(section_id))
    set_expanded_sections_for_notebook(int(notebook_id), cur)


def remove_expanded_section(notebook_id: int, section_id: int):
    m = get_expanded_sections_by_notebook()
    key = str(int(notebook_id))
    cur = m.get(key, set())
    if int(section_id) in cur:
        cur.remove(int(section_id))
    set_expanded_sections_for_notebook(int(notebook_id), cur)


# --- List schemes (ordered/unordered) persistence ---
def get_list_schemes_settings():
    """Return (ordered_scheme, unordered_scheme) strings.
    ordered_scheme in {'classic','decimal'}; unordered_scheme in {'disc-circle-square','disc-only'}.
    """
    s = load_settings()
    ordered = s.get("list_scheme_ordered", "classic")
    unordered = s.get("list_scheme_unordered", "disc-circle-square")
    return ordered, unordered


def set_list_schemes_settings(ordered: str = None, unordered: str = None):
    s = load_settings()
    if ordered in ("classic", "decimal"):
        s["list_scheme_ordered"] = ordered
    if unordered in ("disc-circle-square", "disc-only"):
        s["list_scheme_unordered"] = unordered
    save_settings(s)


# --- Table theme (colors, grid width) ---
def get_table_theme():
    """Return a dict of table theme settings.

    Keys:
      - grid_color: hex string for grid/border color (default '#000000')
      - grid_width: float border width in px (default 1.5)
      - header_bg: hex string for Planning Register header background (default '#F5F5F5')
      - totals_bg: hex string for Planning Register totals row background (default '#F5F5F5')
      - cost_header_bg: hex string for right-side cost table header background (default '#F5F5F5')
    """
    s = load_settings()
    theme = s.get("table_theme", {}) or {}
    # Fill defaults without mutating saved settings
    out = {
        "grid_color": theme.get("grid_color", "#000000"),
        "grid_width": float(theme.get("grid_width", 1.5)),
        "header_bg": theme.get("header_bg", "#F5F5F5"),
        "totals_bg": theme.get("totals_bg", "#F5F5F5"),
        "cost_header_bg": theme.get("cost_header_bg", "#F5F5F5"),
    }
    return out


def set_table_theme(**kwargs):
    """Update one or more table theme fields and persist.

    Example: set_table_theme(grid_color="#444444", grid_width=1.0)
    """
    s = load_settings()
    theme = s.get("table_theme", {}) or {}
    for k, v in kwargs.items():
        if v is None:
            continue
        if k == "grid_width":
            try:
                theme[k] = float(v)
            except Exception:
                continue
        elif k in {"grid_color", "header_bg", "totals_bg", "cost_header_bg"}:
            theme[k] = str(v)
    s["table_theme"] = theme
    save_settings(s)


# --- Default paste mode ---
def get_default_paste_mode():
    """Return default paste mode string in {'rich','text-only','match-style','clean'}; default 'rich'."""
    s = load_settings()
    mode = s.get("default_paste_mode", "rich")
    if mode in ("rich", "text-only", "match-style", "clean"):
        return mode
    return "rich"


def set_default_paste_mode(mode: str):
    """Persist default paste mode if valid."""
    if mode not in ("rich", "text-only", "match-style", "clean"):
        return
    s = load_settings()
    s["default_paste_mode"] = mode
    save_settings(s)


# --- Editor: plain paragraph indent step (pixels) ---
def get_plain_indent_px() -> int:
    """Return the number of pixels to indent/outdent plain paragraphs when pressing Tab/Shift+Tab outside lists/tables."""
    s = load_settings()
    try:
        val = int(s.get("plain_indent_px", 24))
        return max(4, min(160, val))
    except Exception:
        return 24


def set_plain_indent_px(pixels: int):
    try:
        px = int(pixels)
    except Exception:
        return
    px = max(4, min(160, px))
    s = load_settings()
    s["plain_indent_px"] = px
    save_settings(s)


# --- Theme selection ---
def get_theme_name() -> str:
    """Return the current theme name, e.g., 'Default' or 'High Contrast'. Defaults to 'Default'."""
    s = load_settings()
    name = s.get("theme_name", "Default")
    if isinstance(name, str) and name:
        return name
    return "Default"


def set_theme_name(name: str):
    if not isinstance(name, str) or not name:
        return
    s = load_settings()
    s["theme_name"] = name
    save_settings(s)


# --- Table presets (store table formats for reuse) ---
def get_table_presets() -> dict:
    """Return a dict mapping preset name -> preset data dict.

    Supported schema: {"version": 2, "html": "<table>...</table>"}
    """
    s = load_settings()
    presets = s.get("table_presets")
    if isinstance(presets, dict):
        return presets
    return {}


def save_table_preset(name: str, data: dict):
    if not isinstance(name, str) or not name.strip() or not isinstance(data, dict):
        return
    s = load_settings()
    presets = s.get("table_presets")
    if not isinstance(presets, dict):
        presets = {}
    presets[name.strip()] = data
    s["table_presets"] = presets
    save_settings(s)


def delete_table_preset(name: str):
    if not isinstance(name, str) or not name:
        return
    s = load_settings()
    presets = s.get("table_presets")
    if isinstance(presets, dict) and name in presets:
        del presets[name]
        s["table_presets"] = presets
        save_settings(s)


def rename_table_preset(old_name: str, new_name: str):
    if not old_name or not new_name or old_name == new_name:
        return
    s = load_settings()
    presets = s.get("table_presets")
    if not isinstance(presets, dict):
        return
    if old_name in presets:
        presets[new_name] = presets.pop(old_name)
        s["table_presets"] = presets
        save_settings(s)


def list_table_preset_names() -> list:
    return list(get_table_presets().keys())


# --- Default inserted image long side (px) ---
def get_image_insert_long_side() -> int:
    """Return the default long side in pixels for newly inserted images. Default 400, clamped [100, 8000]."""
    s = load_settings()
    try:
        val = int(s.get("image_insert_long_side", 400))
    except Exception:
        val = 400
    # Clamp to sensible bounds
    if val < 100:
        val = 100
    elif val > 8000:
        val = 8000
    return val


def set_image_insert_long_side(pixels: int):
    try:
        px = int(pixels)
    except Exception:
        return
    # Clamp
    if px < 100:
        px = 100
    elif px > 8000:
        px = 8000
    s = load_settings()
    s["image_insert_long_side"] = px
    save_settings(s)


# --- Default inserted video thumbnail long side (px) ---
def get_video_insert_long_side() -> int:
    """Return default long side (px) for newly inserted video thumbnails. Defaults to image setting if absent; clamp [100,8000]."""
    s = load_settings()
    try:
        val = int(s.get("video_insert_long_side", s.get("image_insert_long_side", 400)))
    except Exception:
        val = 400
    if val < 100:
        val = 100
    elif val > 8000:
        val = 8000
    return val


def set_video_insert_long_side(pixels: int):
    try:
        px = int(pixels)
    except Exception:
        return
    if px < 100:
        px = 100
    elif px > 8000:
        px = 8000
    s = load_settings()
    s["video_insert_long_side"] = px
    save_settings(s)


# --- Backups on exit settings ---
def get_exit_backup_dir() -> str:
    """Return the configured folder path for database backups on exit, or empty string if not set."""
    s = load_settings()
    val = s.get("exit_backup_dir", "")
    return str(val) if isinstance(val, str) else ""


def set_exit_backup_dir(path: str):
    """Set the folder path where backups should be stored (no validation performed here)."""
    if not isinstance(path, str):
        return
    s = load_settings()
    s["exit_backup_dir"] = path
    save_settings(s)


def get_backup_on_exit_enabled() -> bool:
    """Return True if making a backup on application exit is enabled."""
    s = load_settings()
    return bool(s.get("backup_on_exit", False))


def set_backup_on_exit_enabled(enabled: bool):
    s = load_settings()
    s["backup_on_exit"] = bool(enabled)
    save_settings(s)


def get_backups_to_keep() -> int:
    """Return how many exit backups to keep (FIFO rotation). Default 5; clamped [1, 999]."""
    s = load_settings()
    try:
        val = int(s.get("backups_to_keep", 5))
    except Exception:
        val = 5
    if val < 1:
        val = 1
    elif val > 999:
        val = 999
    return val


def set_backups_to_keep(n: int):
    try:
        val = int(n)
    except Exception:
        return
    if val < 1:
        val = 1
    elif val > 999:
        val = 999
    s = load_settings()
    s["backups_to_keep"] = val
    save_settings(s)
