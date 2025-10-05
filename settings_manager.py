"""
settings_manager.py
Manages loading and saving application settings, such as the last opened database, in a JSON file.
"""
import json
import os

SETTINGS_FILE = 'settings.json'

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_settings(settings):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f, indent=2)

def get_last_db():
    settings = load_settings()
    return settings.get('last_db')

def set_last_db(db_path):
    settings = load_settings()
    settings['last_db'] = db_path
    save_settings(settings)


# --- Last position/state helpers ---
def get_last_state():
    """Return a dict with keys: last_notebook_id, last_section_id, last_page_id."""
    s = load_settings()
    return {
        'last_notebook_id': s.get('last_notebook_id'),
        'last_section_id': s.get('last_section_id'),
        'last_page_id': s.get('last_page_id'),
    }

def set_last_state(notebook_id=None, section_id=None, page_id=None):
    s = load_settings()
    if notebook_id is not None:
        s['last_notebook_id'] = notebook_id
    if section_id is not None:
        s['last_section_id'] = section_id
    if page_id is not None:
        s['last_page_id'] = page_id
    save_settings(s)


def clear_last_state():
    """Clear last selected notebook/section/page from settings."""
    s = load_settings()
    for k in ('last_notebook_id', 'last_section_id', 'last_page_id'):
        if k in s:
            del s[k]
    save_settings(s)


def get_window_geometry():
    s = load_settings()
    return s.get('window_geometry')  # dict with x, y, w, h

def set_window_geometry(x, y, w, h):
    s = load_settings()
    s['window_geometry'] = {'x': int(x), 'y': int(y), 'w': int(w), 'h': int(h)}
    save_settings(s)

def get_window_maximized():
    s = load_settings()
    return bool(s.get('window_maximized', False))

def set_window_maximized(is_maximized: bool):
    s = load_settings()
    s['window_maximized'] = bool(is_maximized)
    save_settings(s)


# --- Optional: per-section color mapping (for colored tabs and right-pane icons) ---
def get_section_colors():
    """Return a dict mapping section_id (as string) -> color hex string (e.g., '#FF8800')."""
    s = load_settings()
    return s.get('section_colors', {})

def set_section_color(section_id: int, color_hex: str):
    """Persist a color hex string for a section id."""
    s = load_settings()
    colors = s.get('section_colors', {})
    colors[str(int(section_id))] = str(color_hex)
    s['section_colors'] = colors
    save_settings(s)
