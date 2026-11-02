"""
ui_presets.py
Insert, save and manage table and Planning Register presets, from the Insert menu and the
page's right-click menu, and import them from another database (Tools menu). Presets are
stored in the database (see db_presets.py).
"""

import os

from db_presets import (
    KIND_REGISTER,
    KIND_TABLE,
    delete_preset,
    get_preset_html,
    list_preset_names,
    preset_exists,
    read_presets,
    rename_preset,
    save_preset,
    unique_preset_name,
)
from qt_compat import QtCore, QtGui, QtWidgets

QTextCursor = QtGui.QTextCursor
Qt = QtCore.Qt

_LABELS = {
    KIND_TABLE: ("Table", "table"),
    KIND_REGISTER: ("Planning Register", "Planning Register"),
}
_DEFAULT_NAMES = {KIND_TABLE: "My Table", KIND_REGISTER: "My Planning Register"}


def _db_path(widget):
    try:
        return getattr(widget.window(), "_db_path", None)
    except Exception:
        return None


def _require_db(widget, title: str):
    db_path = _db_path(widget)
    if not db_path:
        QtWidgets.QMessageBox.information(widget, title, "Please open a database first.")
    return db_path


# --- Finding the table or register to save ---


def _parent_table(doc, tbl):
    """Return the table whose cell contains tbl, or None."""
    probe = QTextCursor(doc)
    probe.setPosition(max(0, tbl.firstPosition() - 1))
    candidate = probe.currentTable()
    if (
        candidate is not None
        and candidate.firstPosition() < tbl.firstPosition()
        and candidate.lastPosition() >= tbl.lastPosition()
    ):
        return candidate
    return None


def _outermost_table(doc, tbl):
    seen = set()
    while tbl is not None:
        key = (tbl.firstPosition(), tbl.lastPosition())
        if key in seen:
            break
        seen.add(key)
        parent = _parent_table(doc, tbl)
        if parent is None:
            break
        tbl = parent
    return tbl


def tables_in(doc, outer):
    """Yield every table from outer's start to its end, outer included."""
    seen = set()
    block = doc.findBlock(outer.firstPosition())
    while block.isValid() and block.position() <= outer.lastPosition():
        tbl = QTextCursor(block).currentTable()
        if tbl is not None:
            key = (tbl.firstPosition(), tbl.lastPosition())
            if key not in seen:
                seen.add(key)
                yield tbl
        block = block.next()


def register_root(text_edit, tbl):
    """Return the whole Planning Register containing tbl, or None if tbl isn't in one."""
    if tbl is None:
        return None
    from ui_planning_register import _is_planning_register_table

    outer = _outermost_table(text_edit.document(), tbl)
    for t in tables_in(text_edit.document(), outer):
        if _is_planning_register_table(text_edit, t):
            return outer
    return None


def is_in_planning_register(text_edit, tbl) -> bool:
    try:
        return register_root(text_edit, tbl) is not None
    except Exception:
        return False


def _table_html(text_edit, tbl) -> str:
    cur = QTextCursor(text_edit.document())
    cur.setPosition(tbl.firstPosition())
    cur.setPosition(tbl.lastPosition(), QTextCursor.KeepAnchor)
    html = cur.selection().toHtml()
    low = html.lower()
    start = low.find("<table")
    if start < 0:
        return ""
    end = low.rfind("</table>")
    end = end + len("</table>") if end >= 0 else len(html)
    return html[start:end].strip()


# --- Dialogs ---


class _PresetListDialog(QtWidgets.QDialog):
    """A list of preset names, with buttons supplied by the caller."""

    def __init__(self, parent, title: str, prompt: str, names: list):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(340)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel(prompt, self))
        self.list = QtWidgets.QListWidget(self)
        self.list.addItems(names)
        if names:
            self.list.setCurrentRow(0)
        layout.addWidget(self.list)
        self.buttons = QtWidgets.QDialogButtonBox(self)
        layout.addWidget(self.buttons)

    def selected_name(self):
        item = self.list.currentItem()
        return item.text() if item is not None else None

    def set_names(self, names: list, select: str = None):
        self.list.clear()
        self.list.addItems(names)
        rows = [i for i, n in enumerate(names) if n == select]
        if rows:
            self.list.setCurrentRow(rows[0])
        elif names:
            self.list.setCurrentRow(0)


def _none_saved_message(parent, title: str, kind: str):
    label, noun = _LABELS[kind]
    QtWidgets.QMessageBox.information(
        parent,
        title,
        f"No {noun} presets saved yet.\n\n"
        f"To save one, place the caret in a {noun} and choose "
        f"Insert → Save {label} as Preset…, or right-click it.",
    )


def _ask_replace(parent, title: str, kind: str, name: str) -> bool:
    _label, noun = _LABELS[kind]
    reply = QtWidgets.QMessageBox.question(
        parent,
        title,
        f"A {noun} preset named '{name}' already exists.\n\nReplace it?",
        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        QtWidgets.QMessageBox.No,
    )
    return reply == QtWidgets.QMessageBox.Yes


# --- Commands ---


def insert_from_preset(text_edit, kind: str):
    """Let the user pick a preset and insert an exact copy of it at the caret."""
    if text_edit is None:
        return
    label, _noun = _LABELS[kind]
    title = f"Insert {label} from Preset"
    from ui_richtext import _require_editable_page, insert_preset_html

    if not _require_editable_page(text_edit, title):
        return
    db_path = _require_db(text_edit, title)
    if not db_path:
        return
    names = list_preset_names(db_path, kind)
    if not names:
        _none_saved_message(text_edit, title, kind)
        return
    dlg = _PresetListDialog(text_edit, title, "Choose a preset to insert:", names)
    insert_btn = dlg.buttons.addButton("Insert", QtWidgets.QDialogButtonBox.ButtonRole.AcceptRole)
    dlg.buttons.addButton(QtWidgets.QDialogButtonBox.StandardButton.Cancel)
    insert_btn.setDefault(True)
    dlg.buttons.accepted.connect(dlg.accept)
    dlg.buttons.rejected.connect(dlg.reject)
    dlg.list.itemDoubleClicked.connect(lambda _item: dlg.accept())
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return
    name = dlg.selected_name()
    html = get_preset_html(db_path, kind, name) if name else None
    if not html:
        return
    insert_preset_html(text_edit, html)


def _save(text_edit, kind: str, tbl):
    label, noun = _LABELS[kind]
    title = f"Save {label} as Preset"
    db_path = _require_db(text_edit, title)
    if not db_path:
        return
    html = _table_html(text_edit, tbl)
    if not html:
        QtWidgets.QMessageBox.information(text_edit, title, f"Couldn't read the {noun}.")
        return
    name = _DEFAULT_NAMES[kind]
    while True:
        name, ok = QtWidgets.QInputDialog.getText(text_edit, title, "Preset name:", text=name)
        name = (name or "").strip()
        if not ok or not name:
            return
        if not preset_exists(db_path, kind, name) or _ask_replace(text_edit, title, kind, name):
            break
    save_preset(db_path, kind, name, html)
    try:
        from ui_toast import show_toast

        show_toast(text_edit, f"Saved {noun} preset '{name}'")
    except Exception:
        pass


def save_table_preset(text_edit):
    """Save the table under the caret as a preset.

    Tables nested in a layout table are saved with their container, so the layout comes
    back whole. Inside a Planning Register, only the table under the caret is saved."""
    if text_edit is None:
        return
    tbl = text_edit.textCursor().currentTable()
    if tbl is None:
        QtWidgets.QMessageBox.information(
            text_edit, "Save Table as Preset", "Place the caret inside a table to save it."
        )
        return
    if not is_in_planning_register(text_edit, tbl):
        tbl = _outermost_table(text_edit.document(), tbl)
    _save(text_edit, KIND_TABLE, tbl)


def save_register_preset(text_edit):
    """Save the whole Planning Register under the caret as a preset."""
    if text_edit is None:
        return
    tbl = text_edit.textCursor().currentTable()
    root = register_root(text_edit, tbl) if tbl is not None else None
    if root is None:
        QtWidgets.QMessageBox.information(
            text_edit,
            "Save Planning Register as Preset",
            "Place the caret inside a Planning Register to save it.",
        )
        return
    _save(text_edit, KIND_REGISTER, root)


def manage_presets(parent, kind: str):
    """Rename or delete presets of one kind."""
    label, noun = _LABELS[kind]
    title = f"Manage {label} Presets"
    db_path = _require_db(parent, title)
    if not db_path:
        return
    names = list_preset_names(db_path, kind)
    if not names:
        _none_saved_message(parent, title, kind)
        return
    dlg = _PresetListDialog(parent, title, f"Saved {noun} presets:", names)
    rename_btn = dlg.buttons.addButton("Rename…", QtWidgets.QDialogButtonBox.ButtonRole.ActionRole)
    delete_btn = dlg.buttons.addButton("Delete…", QtWidgets.QDialogButtonBox.ButtonRole.ActionRole)
    dlg.buttons.addButton(QtWidgets.QDialogButtonBox.StandardButton.Close)
    dlg.buttons.rejected.connect(dlg.reject)

    def _refresh(select=None):
        current = list_preset_names(db_path, kind)
        dlg.set_names(current, select)
        has = bool(current)
        rename_btn.setEnabled(has)
        delete_btn.setEnabled(has)

    def _rename():
        old = dlg.selected_name()
        if not old:
            return
        new = old
        while True:
            new, ok = QtWidgets.QInputDialog.getText(dlg, "Rename Preset", "New name:", text=new)
            new = (new or "").strip()
            if not ok or not new or new == old:
                return
            if not preset_exists(db_path, kind, new) or _ask_replace(
                dlg, "Rename Preset", kind, new
            ):
                break
        rename_preset(db_path, kind, old, new)
        _refresh(new)

    def _delete():
        name = dlg.selected_name()
        if not name:
            return
        reply = QtWidgets.QMessageBox.question(
            dlg,
            "Delete Preset",
            f"Delete the {noun} preset '{name}'?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if reply != QtWidgets.QMessageBox.Yes:
            return
        delete_preset(db_path, kind, name)
        _refresh()

    rename_btn.clicked.connect(_rename)
    delete_btn.clicked.connect(_delete)
    dlg.list.itemDoubleClicked.connect(lambda _item: _rename())
    dlg.exec_()


# --- Importing from another database ---


class _ImportPresetsDialog(QtWidgets.QDialog):
    """Lists another database's presets for the user to choose from.

    Click, Shift+click, Ctrl+click and Ctrl+A select, as in Windows."""

    def __init__(self, parent, source_path: str, presets: list):
        super().__init__(parent)
        self.setWindowTitle("Import Preset Tables and Registers")
        self.setMinimumSize(460, 360)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(
            QtWidgets.QLabel(
                f"Presets in {os.path.basename(source_path)}.\n"
                "Select the ones to import (Shift+click or Ctrl+click to select several).",
                self,
            )
        )
        self.tree = QtWidgets.QTreeWidget(self)
        self.tree.setColumnCount(2)
        self.tree.setHeaderLabels(["Name", "Type"])
        self.tree.setRootIsDecorated(False)
        self.tree.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        for kind, name, html in presets:
            item = QtWidgets.QTreeWidgetItem([name, _LABELS[kind][0]])
            item.setData(0, Qt.ItemDataRole.UserRole, (kind, name, html))
            self.tree.addTopLevelItem(item)
        self.tree.resizeColumnToContents(0)
        layout.addWidget(self.tree)
        self.buttons = QtWidgets.QDialogButtonBox(self)
        select_all = self.buttons.addButton(
            "Select All", QtWidgets.QDialogButtonBox.ButtonRole.ActionRole
        )
        self.import_btn = self.buttons.addButton(
            "Import", QtWidgets.QDialogButtonBox.ButtonRole.AcceptRole
        )
        self.buttons.addButton(QtWidgets.QDialogButtonBox.StandardButton.Cancel)
        self.import_btn.setDefault(True)
        self.import_btn.setEnabled(False)
        select_all.clicked.connect(self.tree.selectAll)
        self.tree.itemSelectionChanged.connect(
            lambda: self.import_btn.setEnabled(bool(self.tree.selectedItems()))
        )
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def selected_presets(self) -> list:
        items = sorted(self.tree.selectedItems(), key=self.tree.indexOfTopLevelItem)
        return [item.data(0, Qt.ItemDataRole.UserRole) for item in items]


def _ask_clash(parent, kind: str, name: str, remaining: int):
    """Ask what to do about an incoming preset whose name is already used.

    Returns ("replace" | "rename" | "skip" | "cancel", apply_to_remaining)."""
    _label, noun = _LABELS[kind]
    box = QtWidgets.QMessageBox(parent)
    box.setWindowTitle("Preset Already Exists")
    box.setIcon(QtWidgets.QMessageBox.Question)
    box.setText(f"This database already has a {noun} preset named '{name}'.")
    box.setInformativeText("Replace it with the imported one, or import this one under a new name?")
    replace_btn = box.addButton("Replace", QtWidgets.QMessageBox.ButtonRole.AcceptRole)
    rename_btn = box.addButton("Rename…", QtWidgets.QMessageBox.ButtonRole.ActionRole)
    skip_btn = box.addButton("Skip", QtWidgets.QMessageBox.ButtonRole.ActionRole)
    box.addButton(QtWidgets.QMessageBox.StandardButton.Cancel)
    box.setDefaultButton(rename_btn)
    check = None
    if remaining > 0:
        check = QtWidgets.QCheckBox(f"Do this for the next {remaining} name conflicts", box)
        box.setCheckBox(check)
    box.exec_()
    clicked = box.clickedButton()
    choice = {replace_btn: "replace", rename_btn: "rename", skip_btn: "skip"}.get(clicked, "cancel")
    return choice, bool(check is not None and check.isChecked())


def _ask_new_name(parent, kind: str, suggested: str, taken: set):
    """Ask for a name not already taken. Returns None if the user cancels."""
    _label, noun = _LABELS[kind]
    name = suggested
    while True:
        name, ok = QtWidgets.QInputDialog.getText(
            parent, "Rename Imported Preset", f"Import the {noun} preset as:", text=name
        )
        name = (name or "").strip()
        if not ok or not name:
            return None
        if name not in taken:
            return name
        QtWidgets.QMessageBox.information(
            parent, "Rename Imported Preset", f"The name '{name}' is already used."
        )


def import_presets(window):
    """Copy chosen presets from another database into the current one."""
    title = "Import Preset Tables and Registers"
    db_path = _require_db(window, title)
    if not db_path:
        return
    try:
        from settings_manager import get_databases_root

        start_dir = get_databases_root()
    except Exception:
        start_dir = ""
    if not start_dir or not os.path.isdir(start_dir):
        start_dir = os.path.dirname(db_path)
    source, _ = QtWidgets.QFileDialog.getOpenFileName(
        window, "Import Presets From Database", start_dir, "SQLite DB (*.db);;All Files (*)"
    )
    if not source:
        return
    if os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(db_path)):
        QtWidgets.QMessageBox.information(
            window, title, "That's the database you have open. Choose a different one."
        )
        return
    try:
        presets = read_presets(source)
    except Exception as e:
        QtWidgets.QMessageBox.warning(window, title, f"Couldn't read presets from that file:\n{e}")
        return
    if not presets:
        QtWidgets.QMessageBox.information(
            window, title, f"{os.path.basename(source)} has no saved presets."
        )
        return
    dlg = _ImportPresetsDialog(window, source, presets)
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return
    chosen = dlg.selected_presets()

    # Decide every name first, so Cancel leaves the database unchanged
    taken = {kind: set(list_preset_names(db_path, kind)) for kind in (KIND_TABLE, KIND_REGISTER)}
    remaining = sum(1 for kind, name, _html in chosen if name in taken[kind])
    plan = []  # (kind, name to save under, html, outcome)
    standing = None  # A choice the user applied to all remaining conflicts
    for kind, name, html in chosen:
        if name not in taken[kind]:
            plan.append((kind, name, html, "added"))
            taken[kind].add(name)
            continue
        remaining -= 1
        choice = standing
        while True:
            if choice is None:
                choice, for_all = _ask_clash(window, kind, name, remaining)
                if for_all:
                    standing = choice
            if choice == "cancel":
                return
            if choice == "skip":
                break
            if choice == "replace":
                plan.append((kind, name, html, "replaced"))
                break
            suggested = unique_preset_name(name, taken[kind])
            if standing == "rename":
                new_name = suggested
            else:
                new_name = _ask_new_name(window, kind, suggested, taken[kind])
            if new_name is None:
                choice = None  # Ask again about this preset
                continue
            plan.append((kind, new_name, html, "renamed"))
            taken[kind].add(new_name)
            break

    for kind, name, html, _outcome in plan:
        save_preset(db_path, kind, name, html)
    counts = {}
    for *_rest, outcome in plan:
        counts[outcome] = counts.get(outcome, 0) + 1
    parts = [f"{counts[k]} {k}" for k in ("added", "replaced", "renamed") if counts.get(k)]
    skipped = len(chosen) - len(plan)
    if skipped:
        parts.append(f"{skipped} skipped")
    total = len(plan)
    summary = f"Imported {total} preset{'s' if total != 1 else ''}"
    QtWidgets.QMessageBox.information(
        window, title, summary + (f" ({', '.join(parts)})." if parts else ".")
    )
