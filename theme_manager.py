"""Themes: Qt stylesheet (.qss) files that set the program's look.

Themes are found in two folders:
  - the themes folder that ships with NoteBook
  - a themes folder next to settings.json, for looks added without rebuilding

A theme's name comes from its file name: "high-contrast.qss" is "High Contrast". A theme in the
settings folder replaces a shipped theme with the same name.

A border-radius on QMenu gives menus rounded corners; Qt draws them on a transparent window.
"""

import os
import re

from qt_compat import QtWidgets

DEFAULT_THEME = "Default"


def bundled_themes_dir() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "themes")


def user_themes_dir() -> str:
    """The themes folder next to settings.json (may not exist yet)."""
    try:
        from settings_manager import get_settings_dir

        return os.path.join(get_settings_dir(), "themes")
    except Exception:
        return ""


def _name_from_file(filename: str) -> str:
    stem = os.path.splitext(filename)[0]
    return " ".join(w[:1].upper() + w[1:] for w in re.split(r"[-_\s]+", stem) if w)


def list_themes() -> dict:
    """Return {name: path} for every theme, Default first, then by name."""
    found = {}
    for folder in (bundled_themes_dir(), user_themes_dir()):
        try:
            files = sorted(os.listdir(folder)) if folder and os.path.isdir(folder) else []
        except Exception:
            files = []
        for fn in files:
            if fn.lower().endswith(".qss"):
                found[_name_from_file(fn)] = os.path.join(folder, fn)
    names = sorted(found, key=lambda n: (n != DEFAULT_THEME, n.lower()))
    return {n: found[n] for n in names}


def apply_theme(name: str = None) -> str:
    """Apply a theme to the whole program. Returns the name of the theme applied.

    Falls back to Default if the theme can't be found or read."""
    app = QtWidgets.QApplication.instance()
    if app is None:
        return ""
    themes = list_themes()
    if name not in themes:
        name = DEFAULT_THEME if DEFAULT_THEME in themes else next(iter(themes), None)
    qss = ""
    if name:
        try:
            with open(themes[name], "r", encoding="utf-8-sig") as f:
                qss = f.read()
        except Exception:
            qss = ""
    app.setStyleSheet(qss)
    return name or ""
