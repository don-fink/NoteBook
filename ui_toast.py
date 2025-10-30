from PyQt5 import QtCore, QtGui, QtWidgets


def show_toast(parent: QtWidgets.QWidget, text: str, duration_ms: int = 2500):
    """
    Show a small, temporary toast notification over the given parent window.

    - parent: top-level window (e.g., QMainWindow)
    - text: message to display
    - duration_ms: how long the toast stays visible before auto-dismiss
    """
    if parent is None:
        return

    # Ensure we have a top-level window for correct positioning
    window = parent.window() if isinstance(parent, QtWidgets.QWidget) else None
    if window is None or not isinstance(window, QtWidgets.QWidget):
        return

    # Create a transient, frameless label
    label = QtWidgets.QLabel(window)
    label.setText(text)
    label.setAlignment(QtCore.Qt.AlignCenter)
    label.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)
    label.setWindowFlags(
        QtCore.Qt.ToolTip | QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint
    )

    # Styling: dark rounded background with white text
    # Keep font relatively small to avoid intrusive look
    label.setStyleSheet(
        """
        QLabel {
            background-color: rgba(33, 33, 33, 220);
            color: #ffffff;
            border-radius: 8px;
            padding: 8px 12px;
            font-size: 11pt;
        }
        """
    )

    # Size to content
    label.adjustSize()

    # Position: bottom-center above status bar area with a small margin
    margin = 24
    try:
        # Compute global position for the bottom center of the window
        win_geo = window.frameGeometry()
        win_pos = win_geo.topLeft()
        win_width = win_geo.width()
        win_height = win_geo.height()

        label_width = label.width()
        label_height = label.height()

        x = win_pos.x() + (win_width - label_width) // 2
        y = win_pos.y() + win_height - label_height - margin

        label.move(x, y)
    except Exception:
        # Fallback: center on parent
        parent_rect = window.geometry()
        x = parent_rect.x() + (parent_rect.width() - label.width()) // 2
        y = parent_rect.y() + (parent_rect.height() - label.height()) // 2
        label.move(x, y)

    label.show()

    # Auto-close after duration
    QtCore.QTimer.singleShot(duration_ms, label.close)

    return label
