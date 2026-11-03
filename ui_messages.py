"""Message windows: one place that builds every pop-up message and prompt.

Each window is tagged so theme files can style it:
  - Message boxes are QMessageBox#nbMessage with a "kind" property:
    "info", "warning", "error" or "question".
  - Their buttons have a "role" property: "primary" (the main action), "destructive"
    (an action that deletes or overwrites) or "secondary".
  - Prompts for a name, number or list choice are QInputDialog#nbPrompt.

Wording: the text is a short headline saying what happened or what to do. The optional info
line explains or suggests the next step. Technical error text goes in details, shown only
when the user clicks Show Details….
"""

from qt_compat import QtCore, QtWidgets

QMB = QtWidgets.QMessageBox

_ICONS = {
    "info": QMB.Icon.Information,
    "warning": QMB.Icon.Warning,
    "error": QMB.Icon.Critical,
    "question": QMB.Icon.Question,
    "about": QMB.Icon.Information,  # Long rich text, so the headline style doesn't apply
}
_QT_ROLES = {
    "primary": QMB.ButtonRole.AcceptRole,
    "destructive": QMB.ButtonRole.DestructiveRole,
    "secondary": QMB.ButtonRole.ActionRole,
}

NO_PAGE_TEXT = "Open or create a page first."


def _repolish(widget):
    """Re-apply the stylesheet after changing a property that selectors look at."""
    try:
        widget.style().unpolish(widget)
        widget.style().polish(widget)
    except Exception:
        pass


# --- Building blocks, for windows with their own buttons ---


def make_box(parent, kind: str, title: str, text: str, info: str = None, details: str = None):
    """A message box of the given kind ("info", "warning", "error" or "question")."""
    box = QMB(parent)
    box.setObjectName("nbMessage")
    box.setProperty("kind", kind)
    box.setIcon(_ICONS[kind])
    box.setWindowTitle(title)
    box.setText(text)
    if info:
        box.setInformativeText(info)
    if details:
        box.setDetailedText(str(details))
    return box


def add_button(box, label: str, role: str = "secondary", cancel: bool = False):
    """Add a button. role is "primary", "destructive" or "secondary".

    cancel=True makes it the button that Escape and closing the window choose."""
    qt_role = QMB.ButtonRole.RejectRole if cancel else _QT_ROLES[role]
    btn = box.addButton(label, qt_role)
    btn.setProperty("role", role)
    _repolish(btn)
    # The box sizes itself from its buttons' minimum widths, which would cut off long labels
    btn.setMinimumWidth(btn.sizeHint().width())
    if cancel:
        box.setEscapeButton(btn)
    return btn


def run(box, default=None):
    """Show a box and return the button that was clicked (None if closed another way)."""
    if default is not None:
        box.setDefaultButton(default)
    try:
        box.exec()
    except AttributeError:
        box.exec_()
    return box.clickedButton()


# --- Messages ---


def _notify(parent, kind, title, text, info=None, details=None):
    try:
        box = make_box(parent, kind, title, text, info, details)
        ok = add_button(box, "OK", "primary")
        box.setEscapeButton(ok)
        run(box, ok)
    except Exception:
        pass


def info(parent, title: str, text: str, info: str = None):
    """Tell the user something: a result, or why a command can't run yet."""
    _notify(parent, "info", title, text, info)


def warning(parent, title: str, text: str, info: str = None, details: str = None):
    """Something didn't go as expected, but nothing was lost."""
    _notify(parent, "warning", title, text, info, details)


def error(parent, title: str, text: str, info: str = None, details: str = None):
    """Something failed. Put the technical error text in details."""
    _notify(parent, "error", title, text, info, details)


def no_page(parent, title: str):
    """The command needs an open, editable page."""
    _notify(parent, "info", title, NO_PAGE_TEXT)


def confirm(
    parent,
    title: str,
    text: str,
    info: str = None,
    action: str = "OK",
    destructive: bool = False,
    kind: str = "question",
) -> bool:
    """Ask before doing something. Returns True if the user chose the action.

    The action button is labelled with what it does ("Delete", "Replace"). For destructive
    actions Cancel is the default, so pressing Enter doesn't delete anything."""
    try:
        box = make_box(parent, kind, title, text, info)
        act = add_button(box, action, "destructive" if destructive else "primary")
        cancel = add_button(box, "Cancel", "secondary", cancel=True)
        return run(box, cancel if destructive else act) is act
    except Exception:
        return False


def choose(
    parent,
    title: str,
    text: str,
    choices: list,
    info: str = None,
    default: str = None,
    kind: str = "question",
    cancel_label: str = "Cancel",
):
    """Ask the user to pick one of several actions. Returns the chosen label, or None for Cancel.

    choices is a list of labels; a label can be given as (label, "destructive") to mark it.
    The default choice is shown as the main button; with no default, Enter means Cancel."""
    try:
        box = make_box(parent, kind, title, text, info)
        buttons = {}
        for choice in choices:
            label, role = choice if isinstance(choice, tuple) else (choice, None)
            role = role or ("primary" if label == default else "secondary")
            buttons[add_button(box, label, role)] = label
        cancel = add_button(box, cancel_label, "secondary", cancel=True)
        default_btn = next((b for b, lbl in buttons.items() if lbl == default), None)
        return buttons.get(run(box, default_btn or cancel))
    except Exception:
        return None


# --- Prompts: same arguments and (value, ok) results as QInputDialog's ---


class _Prompt(QtWidgets.QInputDialog):
    """A QInputDialog whose spacing a theme can set, since stylesheet padding can't move its
    controls: QInputDialog#nbPrompt { qproperty-boxPadding: 18; qproperty-boxSpacing: 10; }"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._padding = 18
        self._spacing = 10

    def _get_padding(self):
        return self._padding

    def _set_padding(self, value):
        self._padding = int(value)

    def _get_spacing(self):
        return self._spacing

    def _set_spacing(self, value):
        self._spacing = int(value)

    boxPadding = QtCore.pyqtProperty(int, _get_padding, _set_padding)
    boxSpacing = QtCore.pyqtProperty(int, _get_spacing, _set_spacing)

    def showEvent(self, event):
        # Qt builds the layout just before the first show, so adjust it here
        try:
            self.ensurePolished()
            layout = self.layout()
            if layout is not None:
                p = self._padding
                layout.setContentsMargins(p, p, p, p)
                layout.setSpacing(self._spacing)
                self.adjustSize()
        except Exception:
            pass
        super().showEvent(event)


def _prompt(parent, title: str, label: str):
    dlg = _Prompt(parent)
    dlg.setObjectName("nbPrompt")
    dlg.setWindowTitle(title)
    dlg.setLabelText(label)
    return dlg


def _accepted(dlg) -> bool:
    try:
        result = dlg.exec()
    except AttributeError:
        result = dlg.exec_()
    return result == QtWidgets.QDialog.DialogCode.Accepted


def get_text(parent, title: str, label: str, echo=None, text: str = ""):
    """Ask for a line of text. Returns (text, ok)."""
    try:
        dlg = _prompt(parent, title, label)
        if echo is not None:
            dlg.setTextEchoMode(echo)
        dlg.setTextValue(text or "")
        ok = _accepted(dlg)
        return (dlg.textValue() if ok else ""), ok
    except Exception:
        return "", False


def get_int(
    parent, title: str, label: str, value: int = 0, minimum=-2147483647, maximum=2147483647, step=1
):
    """Ask for a whole number. Returns (value, ok)."""
    try:
        dlg = _prompt(parent, title, label)
        dlg.setInputMode(QtWidgets.QInputDialog.InputMode.IntInput)
        dlg.setIntRange(int(minimum), int(maximum))
        dlg.setIntStep(int(step))
        dlg.setIntValue(int(value))
        ok = _accepted(dlg)
        return (dlg.intValue() if ok else int(value)), ok
    except Exception:
        return int(value), False


def get_double(
    parent,
    title: str,
    label: str,
    value=0.0,
    minimum=-2147483647.0,
    maximum=2147483647.0,
    decimals=1,
):
    """Ask for a number with decimals. Returns (value, ok)."""
    try:
        dlg = _prompt(parent, title, label)
        dlg.setInputMode(QtWidgets.QInputDialog.InputMode.DoubleInput)
        dlg.setDoubleRange(float(minimum), float(maximum))
        dlg.setDoubleDecimals(int(decimals))
        dlg.setDoubleValue(float(value))
        ok = _accepted(dlg)
        return (dlg.doubleValue() if ok else float(value)), ok
    except Exception:
        return float(value), False


def get_item(parent, title: str, label: str, items, current: int = 0, editable: bool = False):
    """Ask the user to pick from a list. Returns (item, ok)."""
    try:
        items = [str(i) for i in items]
        dlg = _prompt(parent, title, label)
        dlg.setComboBoxItems(items)
        dlg.setComboBoxEditable(bool(editable))
        if items:
            dlg.setTextValue(items[max(0, min(int(current), len(items) - 1))])
        ok = _accepted(dlg)
        return (dlg.textValue() if ok else ""), ok
    except Exception:
        return "", False
