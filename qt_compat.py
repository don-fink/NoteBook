"""Compatibility helpers to run legacy PyQt5-style code on PyQt6.

This module is intentionally small and targeted for migration bootstrapping.
Import it early in startup so modules that still do ``from PyQt5 ...`` keep
working while the codebase is migrated incrementally.
"""

from __future__ import annotations

import sys
import types


USING_PYQT6 = False


try:
    from PyQt6 import QtCore, QtGui, QtPrintSupport, QtWidgets, uic

    USING_PYQT6 = True
except Exception:  # pragma: no cover - fallback for old environments
    from PyQt5 import QtCore, QtGui, QtPrintSupport, QtWidgets, uic


def _get_attr_path(root, path: str):
    cur = root
    for part in path.split("."):
        cur = getattr(cur, part)
    return cur


def _set_if_missing(obj, name: str, value):
    if not hasattr(obj, name):
        setattr(obj, name, value)


def _install_qt_enum_aliases() -> None:
    if not USING_PYQT6:
        return

    qt = QtCore.Qt
    aliases = {
        # Cursor / context / drag-drop
        "WaitCursor": "CursorShape.WaitCursor",
        "SizeHorCursor": "CursorShape.SizeHorCursor",
        "PointingHandCursor": "CursorShape.PointingHandCursor",
        "IBeamCursor": "CursorShape.IBeamCursor",
        "CustomContextMenu": "ContextMenuPolicy.CustomContextMenu",
        "MoveAction": "DropAction.MoveAction",
        # Focus / shortcut / modality / window flags / state
        "OtherFocusReason": "FocusReason.OtherFocusReason",
        "WidgetWithChildrenShortcut": "ShortcutContext.WidgetWithChildrenShortcut",
        "ApplicationModal": "WindowModality.ApplicationModal",
        "WindowNoState": "WindowState.WindowNoState",
        "Window": "WindowType.Window",
        "ToolTip": "WindowType.ToolTip",
        "FramelessWindowHint": "WindowType.FramelessWindowHint",
        "WindowStaysOnTopHint": "WindowType.WindowStaysOnTopHint",
        # Attributes / style / text interaction
        "WA_TransparentForMouseEvents": "WidgetAttribute.WA_TransparentForMouseEvents",
        "WA_NoSystemBackground": "WidgetAttribute.WA_NoSystemBackground",
        "DashLine": "PenStyle.DashLine",
        "ToolButtonIconOnly": "ToolButtonStyle.ToolButtonIconOnly",
        "LinksAccessibleByMouse": "TextInteractionFlag.LinksAccessibleByMouse",
        "RichText": "TextFormat.RichText",
        # Alignment
        "AlignLeft": "AlignmentFlag.AlignLeft",
        "AlignRight": "AlignmentFlag.AlignRight",
        "AlignHCenter": "AlignmentFlag.AlignHCenter",
        "AlignVCenter": "AlignmentFlag.AlignVCenter",
        "AlignCenter": "AlignmentFlag.AlignCenter",
        "AlignJustify": "AlignmentFlag.AlignJustify",
        "TextSingleLine": "TextFlag.TextSingleLine",
        # Global colors/brush styles
        "transparent": "GlobalColor.transparent",
        "NoBrush": "BrushStyle.NoBrush",
        "NoPen": "PenStyle.NoPen",
        # Modifiers
        "ControlModifier": "KeyboardModifier.ControlModifier",
        "ShiftModifier": "KeyboardModifier.ShiftModifier",
        "AltModifier": "KeyboardModifier.AltModifier",
        # Mouse button
        "LeftButton": "MouseButton.LeftButton",
        # Item flags
        "ItemIsEnabled": "ItemFlag.ItemIsEnabled",
        "ItemIsSelectable": "ItemFlag.ItemIsSelectable",
        "ItemIsDragEnabled": "ItemFlag.ItemIsDragEnabled",
        "ItemIsDropEnabled": "ItemFlag.ItemIsDropEnabled",
        # Common keys used throughout app
        "Key_Backspace": "Key.Key_Backspace",
        "Key_Delete": "Key.Key_Delete",
        "Key_Return": "Key.Key_Return",
        "Key_Enter": "Key.Key_Enter",
        "Key_V": "Key.Key_V",
        "Key_X": "Key.Key_X",
        "Key_Insert": "Key.Key_Insert",
        "Key_Tab": "Key.Key_Tab",
        "Key_Backtab": "Key.Key_Backtab",
    }
    for name, path in aliases.items():
        try:
            _set_if_missing(qt, name, _get_attr_path(qt, path))
        except Exception:
            pass


def _install_api_aliases() -> None:
    if not USING_PYQT6:
        return

    # Qt6 moved these classes; preserve old lookup paths used in legacy code.
    if not hasattr(QtWidgets, "QAction") and hasattr(QtGui, "QAction"):
        QtWidgets.QAction = QtGui.QAction
    if not hasattr(QtWidgets, "QActionGroup") and hasattr(QtGui, "QActionGroup"):
        QtWidgets.QActionGroup = QtGui.QActionGroup
    if not hasattr(QtWidgets, "QShortcut") and hasattr(QtGui, "QShortcut"):
        QtWidgets.QShortcut = QtGui.QShortcut
    if not hasattr(QtWidgets, "QKeySequence") and hasattr(QtGui, "QKeySequence"):
        QtWidgets.QKeySequence = QtGui.QKeySequence

    # QSizePolicy enum values moved under QSizePolicy.Policy in PyQt6.
    _sp_policy_names = [
        "Fixed", "Minimum", "Maximum", "Preferred", "Expanding",
        "MinimumExpanding", "Ignored",
    ]
    for _spn in _sp_policy_names:
        try:
            _set_if_missing(
                QtWidgets.QSizePolicy, _spn,
                getattr(QtWidgets.QSizePolicy.Policy, _spn)
            )
        except Exception:
            pass

    # QToolButton.ToolButtonPopupMode values moved under nested enum in PyQt6.
    for _tbpn in ("DelayedPopup", "MenuButtonPopup", "InstantPopup"):
        try:
            _set_if_missing(
                QtWidgets.QToolButton, _tbpn,
                getattr(QtWidgets.QToolButton.ToolButtonPopupMode, _tbpn)
            )
        except Exception:
            pass

    # QFrame shape/shadow enums moved under nested enums in PyQt6.
    for _fn in ("NoFrame", "Box", "Panel", "WinPanel", "HLine", "VLine", "StyledPanel"):
        try:
            _set_if_missing(QtWidgets.QFrame, _fn,
                            getattr(QtWidgets.QFrame.Shape, _fn))
        except Exception:
            pass
    for _fsn in ("Plain", "Raised", "Sunken"):
        try:
            _set_if_missing(QtWidgets.QFrame, _fsn,
                            getattr(QtWidgets.QFrame.Shadow, _fsn))
        except Exception:
            pass

    # Provide exec_ aliases expected by PyQt5-era code.
    for cls in (
        QtWidgets.QApplication,
        QtWidgets.QDialog,
        QtWidgets.QMenu,
        QtWidgets.QMessageBox,
        QtWidgets.QFileDialog,
        QtWidgets.QColorDialog,
        QtWidgets.QFontDialog,
        QtPrintSupport.QPrintDialog,
        QtPrintSupport.QPrintPreviewDialog,
    ):
        try:
            if hasattr(cls, "exec") and not hasattr(cls, "exec_"):
                def _exec_compat(self, *args, **kwargs):
                    return self.exec(*args, **kwargs)

                setattr(cls, "exec_", _exec_compat)
        except Exception:
            pass

    for _name in ("Accepted", "Rejected"):
        try:
            _set_if_missing(
                QtWidgets.QDialog,
                _name,
                getattr(QtWidgets.QDialog.DialogCode, _name),
            )
        except Exception:
            pass

    for _name in (
        "NoIcon",
        "Question",
        "Information",
        "Warning",
        "Critical",
    ):
        try:
            _set_if_missing(
                QtWidgets.QMessageBox,
                _name,
                getattr(QtWidgets.QMessageBox.Icon, _name),
            )
        except Exception:
            pass

    for _name in (
        "NoButton",
        "Ok",
        "Open",
        "Save",
        "Cancel",
        "Close",
        "Discard",
        "Apply",
        "Reset",
        "RestoreDefaults",
        "Help",
        "SaveAll",
        "Yes",
        "YesToAll",
        "No",
        "NoToAll",
        "Abort",
        "Retry",
        "Ignore",
    ):
        try:
            _set_if_missing(
                QtWidgets.QMessageBox,
                _name,
                getattr(QtWidgets.QMessageBox.StandardButton, _name),
            )
        except Exception:
            pass

    for _name in (
        "InvalidRole",
        "AcceptRole",
        "RejectRole",
        "DestructiveRole",
        "ActionRole",
        "HelpRole",
        "YesRole",
        "NoRole",
        "ResetRole",
        "ApplyRole",
    ):
        try:
            _set_if_missing(
                QtWidgets.QMessageBox,
                _name,
                getattr(QtWidgets.QMessageBox.ButtonRole, _name),
            )
        except Exception:
            pass

    try:
        if not hasattr(QtWidgets.QFileDialog, "Options") and hasattr(QtWidgets.QFileDialog, "Option"):
            def _qfiledialog_options(value=0):
                return QtWidgets.QFileDialog.Option(value)

            QtWidgets.QFileDialog.Options = staticmethod(_qfiledialog_options)
    except Exception:
        pass

    for _name in (
        "ShowDirsOnly",
        "DontResolveSymlinks",
        "DontConfirmOverwrite",
        "DontUseNativeDialog",
        "ReadOnly",
        "HideNameFilterDetails",
        "DontUseCustomDirectoryIcons",
    ):
        try:
            _set_if_missing(
                QtWidgets.QFileDialog,
                _name,
                getattr(QtWidgets.QFileDialog.Option, _name),
            )
        except Exception:
            pass

    # QTextFormat enum names moved under nested enums in Qt6.
    try:
        _set_if_missing(
            QtGui.QTextFormat,
            "UserProperty",
            QtGui.QTextFormat.Property.UserProperty,
        )
    except Exception:
        pass
    try:
        _set_if_missing(
            QtGui.QTextFormat,
            "ImageObject",
            QtGui.QTextFormat.ObjectType.ImageObject,
        )
    except Exception:
        pass
    try:
        _set_if_missing(
            QtGui.QTextFormat,
            "ImageName",
            QtGui.QTextFormat.Property.ImageName,
        )
    except Exception:
        pass
    try:
        _set_if_missing(QtGui.QFont, "Bold", int(QtGui.QFont.Weight.Bold))
    except Exception:
        pass
    try:
        _set_if_missing(QtGui.QFont, "Normal", int(QtGui.QFont.Weight.Normal))
    except Exception:
        pass
    try:
        _set_if_missing(
            QtGui.QPainter,
            "Antialiasing",
            QtGui.QPainter.RenderHint.Antialiasing,
        )
    except Exception:
        pass

    # PyQt5-style QKeySequence constants moved under StandardKey in Qt6.
    key_aliases = {
        "Open": "Open",
        "Save": "Save",
        "SaveAs": "SaveAs",
        "Print": "Print",
        "Undo": "Undo",
        "Redo": "Redo",
        "Paste": "Paste",
        "Bold": "Bold",
        "Italic": "Italic",
        "Underline": "Underline",
    }
    for old_name, std_name in key_aliases.items():
        try:
            std_val = getattr(QtGui.QKeySequence.StandardKey, std_name)
            _set_if_missing(
                QtGui.QKeySequence,
                old_name,
                std_val,
            )
        except Exception:
            pass

    # PyQt5-style QTextCursor enums moved under nested enums in Qt6.
    cursor_aliases = {
        # MoveMode
        "MoveAnchor": "MoveMode.MoveAnchor",
        "KeepAnchor": "MoveMode.KeepAnchor",
        # MoveOperation
        "NoMove": "MoveOperation.NoMove",
        "Start": "MoveOperation.Start",
        "End": "MoveOperation.End",
        "Left": "MoveOperation.Left",
        "Right": "MoveOperation.Right",
        "NextCharacter": "MoveOperation.NextCharacter",
        "PreviousCharacter": "MoveOperation.PreviousCharacter",
        "NextBlock": "MoveOperation.NextBlock",
        "PreviousBlock": "MoveOperation.PreviousBlock",
        # SelectionType
        "WordUnderCursor": "SelectionType.WordUnderCursor",
        "BlockUnderCursor": "SelectionType.BlockUnderCursor",
    }
    for old_name, path in cursor_aliases.items():
        try:
            _set_if_missing(QtGui.QTextCursor, old_name, _get_attr_path(QtGui.QTextCursor, path))
        except Exception:
            pass

    # PyQt5-style QEvent enums moved under QEvent.Type in Qt6.
    qevent_type_names = [
        "None_",
        "Timer", "MouseButtonPress", "MouseButtonRelease", "MouseButtonDblClick",
        "MouseMove", "KeyPress", "KeyRelease", "FocusIn", "FocusOut",
        "FocusAboutToChange", "Enter", "Leave", "Paint", "Move", "Resize",
        "Show", "Hide", "Close", "Quit", "ParentChange", "ParentAboutToChange",
        "ThreadChange", "WindowActivate", "WindowDeactivate", "ShowToParent",
        "HideToParent", "Wheel", "WindowTitleChange", "WindowIconChange",
        "ApplicationWindowIconChange", "ApplicationFontChange",
        "ApplicationLayoutDirectionChange", "ApplicationPaletteChange",
        "PaletteChange", "Clipboard", "Speech", "MetaCall", "SockAct",
        "WinEventAct", "DeferredDelete", "DragEnter", "DragMove", "DragLeave",
        "Drop", "DragResponse", "ChildAdded", "ChildPolished", "ChildRemoved",
        "ShowWindowRequest", "PolishRequest", "Polish", "LayoutRequest",
        "UpdateRequest", "UpdateLater", "EmbeddingControl",
        "ActivateControl", "DeactivateControl", "ContextMenu", "InputMethod",
        "TabletMove", "LocaleChange", "LanguageChange",
        "LayoutDirectionChange", "Style", "TabletPress", "TabletRelease",
        "OkRequest", "IconDrag", "FontChange", "EnabledChange",
        "ActivationChange", "StyleChange", "IconTextChange", "ModifiedChange",
        "WindowBlocked", "WindowUnblocked", "WindowStateChange",
        "ReadOnlyChange", "MouseTrackingChange", "ToolTip", "WhatsThis",
        "StatusTip", "ActionChanged", "ActionAdded", "ActionRemoved",
        "FileOpen", "Shortcut", "ShortcutOverride",
        "WhatsThisClicked", "ToolBarChange",
        "ApplicationActivated", "ApplicationDeactivated",
        "QueryWhatsThis", "EnterWhatsThisMode", "LeaveWhatsThisMode",
        "ZOrderChange", "HoverEnter", "HoverLeave", "HoverMove",
        "AcceptDropsChange", "GraphicsSceneMouseMove",
        "GraphicsSceneMousePress", "GraphicsSceneMouseRelease",
        "GraphicsSceneMouseDoubleClick", "GraphicsSceneContextMenu",
        "GraphicsSceneHoverEnter", "GraphicsSceneHoverMove",
        "GraphicsSceneHoverLeave", "GraphicsSceneHelp",
        "GraphicsSceneDragEnter", "GraphicsSceneDragMove",
        "GraphicsSceneDragLeave", "GraphicsSceneDrop",
        "GraphicsSceneWheel", "KeyboardLayoutChange",
        "DynamicPropertyChange", "TabletEnterProximity",
        "TabletLeaveProximity", "NonClientAreaMouseMove",
        "NonClientAreaMouseButtonPress", "NonClientAreaMouseButtonRelease",
        "NonClientAreaMouseButtonDblClick", "MacSizeChange",
        "ContentsRectChange", "MacGLWindowChange",
        "FutureCallOut", "GraphicsSceneResize", "GraphicsSceneMove",
        "CursorChange", "ToolTipChange", "NetworkReplyUpdated",
        "GrabMouse", "UngrabMouse", "GrabKeyboard", "UngrabKeyboard",
        "StateMachineSignal", "StateMachineWrapped",
        "TouchBegin", "TouchUpdate", "TouchEnd",
        "NativeGesture", "RequestSoftwareInputPanel",
        "CloseSoftwareInputPanel", "WinIdChange",
        "Gesture", "GestureOverride",
        "ScrollPrepare", "Scroll",
        "Expose", "InputMethodQuery", "OrientationChange",
        "TouchCancel", "ThemeChange", "SockClose",
        "PlatformPanel", "StyleAnimationUpdate",
        "ApplicationStateChange", "WindowChangeInternal",
        "ScreenChangeInternal", "PlatformSurface",
        "Pointer", "TabletTrackingChange",
        "ChildWindowAdded", "ChildWindowRemoved",
        "EnterEditFocus", "LeaveEditFocus",
    ]
    for _name in qevent_type_names:
        try:
            _val = getattr(QtCore.QEvent.Type, _name, None)
            if _val is not None:
                _set_if_missing(QtCore.QEvent, _name, _val)
        except Exception:
            pass


def _install_pyqt5_module_shim() -> None:
    """Expose PyQt6 modules under PyQt5 import paths during migration."""
    if not USING_PYQT6:
        return

    if "PyQt5" in sys.modules:
        return

    shim = types.ModuleType("PyQt5")
    shim.QtCore = QtCore
    shim.QtGui = QtGui
    shim.QtWidgets = QtWidgets
    shim.QtPrintSupport = QtPrintSupport
    shim.uic = uic

    sys.modules["PyQt5"] = shim
    sys.modules["PyQt5.QtCore"] = QtCore
    sys.modules["PyQt5.QtGui"] = QtGui
    sys.modules["PyQt5.QtWidgets"] = QtWidgets
    sys.modules["PyQt5.QtPrintSupport"] = QtPrintSupport
    sys.modules["PyQt5.uic"] = uic


_install_qt_enum_aliases()
_install_api_aliases()
_install_pyqt5_module_shim()
