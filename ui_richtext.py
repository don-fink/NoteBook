"""
ui_richtext.py
Adds a lightweight rich text toolbar to a QTextEdit inside a tab.
Supports: Undo/Redo, Bold/Italic/Underline/Strike, Font family/size,
Text color/Highlight, Align L/C/R/Justify, Bullets/Numbers, Clear formatting,
Insert image, Horizontal rule.
"""

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtGui import QTextBlockFormat
import re
import os
from PyQt5.QtCore import QEvent, QObject, QPoint, QRect, QSize, Qt, QUrl, QTimer
from PyQt5.QtGui import (
    QColor,
    QDesktopServices,
    QFont,
    QIcon,
    QKeySequence,
    QPainter,
    QBrush,
    QPalette,
    QPen,
    QImage,
    QPixmap,
    QCursor,
    QTextCharFormat,
    QTextCursor,
    QTextFrameFormat,
    QTextImageFormat,
    QTextList,
    QTextListFormat,
    QTextFormat,
    QTextLength,
    QTextTableFormat,
    QTextTableCellFormat,
    QSyntaxHighlighter,
    QTextCharFormat as _QTextCharFormat,
)

# ----------------------------- Image helpers -----------------------------
_RAW_EXTS = {
    "dng", "nef", "cr2", "cr3", "arw", "orf", "rw2", "raf", "srw", "pef",
    "rw1", "3fr", "erf", "kdc", "mrw", "nrw", "ptx", "r3d", "sr2", "x3f"
}

def _is_raw_ext(name: str) -> bool:
    try:
        if not name:
            return False
        import os as _os
        ext = _os.path.splitext(str(name))[1].lstrip(".").lower()
        return ext in _RAW_EXTS
    except Exception:
        return False
def _qimage_dims(text_edit: QtWidgets.QTextEdit, name: str):
    try:
        if not name:
            return None, None
        base = getattr(text_edit.window(), "_media_root", None)
        path = name
        if base and name and not os.path.isabs(name):
            path = os.path.join(base, name)
        # Avoid trying to load RAW formats; Qt/libtiff may spam stderr and fail
        if _is_raw_ext(path):
            return None, None
        img = QImage(path)
        if not img.isNull():
            return img.width(), img.height()
    except Exception:
        pass
    return None, None


def _image_info_at_cursor(text_edit: QtWidgets.QTextEdit):
    cur = text_edit.textCursor()
    # Check char under cursor
    fmt = cur.charFormat()
    if fmt is not None and (
        (hasattr(fmt, "isImageFormat") and fmt.isImageFormat())
        or fmt.objectType() == QTextFormat.ImageObject
    ):
        imgf = QTextImageFormat(fmt)
        w = float(imgf.width() or 0.0)
        h = float(imgf.height() or 0.0)
        iw, ih = _qimage_dims(text_edit, imgf.name())
        return {"cursor_pos": cur.position(), "name": imgf.name(), "w": w, "h": h, "iw": iw, "ih": ih}
    # Try selecting next char (caret just before image)
    c2 = QTextCursor(cur)
    c2.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
    fmt2 = c2.charFormat()
    if fmt2 is not None and (
        (hasattr(fmt2, "isImageFormat") and fmt2.isImageFormat())
        or fmt2.objectType() == QTextFormat.ImageObject
    ):
        imgf = QTextImageFormat(fmt2)
        w = float(imgf.width() or 0.0)
        h = float(imgf.height() or 0.0)
        iw, ih = _qimage_dims(text_edit, imgf.name())
        return {"cursor_pos": cur.position(), "name": imgf.name(), "w": w, "h": h, "iw": iw, "ih": ih}
    # Try selecting previous char (caret just after image)
    c3 = QTextCursor(cur)
    c3.movePosition(QTextCursor.Left, QTextCursor.KeepAnchor, 1)
    fmt3 = c3.charFormat()
    if fmt3 is not None and (
        (hasattr(fmt3, "isImageFormat") and fmt3.isImageFormat())
        or fmt3.objectType() == QTextFormat.ImageObject
    ):
        imgf = QTextImageFormat(fmt3)
        w = float(imgf.width() or 0.0)
        h = float(imgf.height() or 0.0)
        iw, ih = _qimage_dims(text_edit, imgf.name())
        # Use the start of the selection as the image position
        pos = min(c3.position(), c3.anchor())
        return {"cursor_pos": pos, "name": imgf.name(), "w": w, "h": h, "iw": iw, "ih": ih}
    return None


def _image_info_at_view_pos(text_edit: QtWidgets.QTextEdit, view_pos: QPoint):
    try:
        cur = text_edit.cursorForPosition(view_pos)
        cands = [QTextCursor(cur)]
        c1 = QTextCursor(cur)
        c1.movePosition(QTextCursor.Left)
        cands.append(c1)
        c2 = QTextCursor(cur)
        c2.movePosition(QTextCursor.Right)
        cands.append(c2)
        for c in cands:
            fmt = c.charFormat()
            if fmt is not None and (
                (hasattr(fmt, "isImageFormat") and fmt.isImageFormat())
                or fmt.objectType() == QTextFormat.ImageObject
            ):
                imgf = QTextImageFormat(fmt)
                w = float(imgf.width() or 0.0)
                h = float(imgf.height() or 0.0)
                iw, ih = _qimage_dims(text_edit, imgf.name())
                return {"cursor_pos": c.position(), "name": imgf.name(), "w": w, "h": h, "iw": iw, "ih": ih}
            csel = QTextCursor(c)
            csel.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            fmt2 = csel.charFormat()
            if fmt2 is not None and (
                (hasattr(fmt2, "isImageFormat") and fmt2.isImageFormat())
                or fmt2.objectType() == QTextFormat.ImageObject
            ):
                imgf = QTextImageFormat(fmt2)
                w = float(imgf.width() or 0.0)
                h = float(imgf.height() or 0.0)
                iw, ih = _qimage_dims(text_edit, imgf.name())
                return {"cursor_pos": c.position(), "name": imgf.name(), "w": w, "h": h, "iw": iw, "ih": ih}
    except Exception:
        pass
    return None


def _open_image_properties(text_edit: QtWidgets.QTextEdit, info: dict = None):
    if info is None:
        info = _image_info_at_cursor(text_edit)
    if not info:
        return
    _image_properties_dialog_apply(text_edit, info)


def _install_image_context_menu(text_edit: QtWidgets.QTextEdit):
    # Defined later in the file using a robust handler class; this stub remains for
    # backward compatibility if referenced before the full definition is parsed.
    try:
        pass
    except Exception:
        pass


# Public installer: enable image support (menu, shortcuts, disable drops)
def install_image_support(text_edit: QtWidgets.QTextEdit):
    if text_edit is None:
        return
    try:
        text_edit.setAcceptDrops(False)
        if hasattr(text_edit, "viewport") and text_edit.viewport() is not None:
            text_edit.viewport().setAcceptDrops(False)
    except Exception:
        pass
        _install_image_context_menu(text_edit)
    try:
        _install_image_shortcuts(text_edit)
    except Exception:
        pass


# ----------------------------- Image insertion -----------------------------


def _install_image_shortcuts(text_edit: QtWidgets.QTextEdit):
    # Ctrl+Shift+I opens Image Properties; F2 as backup
    sc1 = QtWidgets.QShortcut(QKeySequence("Ctrl+Shift+I"), text_edit)
    sc1.setContext(Qt.WidgetWithChildrenShortcut)
    sc1.activated.connect(lambda: _open_image_properties(text_edit))
    sc2 = QtWidgets.QShortcut(QKeySequence("F2"), text_edit)
    sc2.setContext(Qt.WidgetWithChildrenShortcut)
    sc2.activated.connect(lambda: _open_image_properties(text_edit))


# ----------------------------- HTML Source dialog -----------------------------
class _HtmlHighlighter(QSyntaxHighlighter):
    def __init__(self, doc):
        super().__init__(doc)
        self._fmt_tag = _QTextCharFormat()
        self._fmt_tag.setForeground(QColor("#0033aa"))
        self._fmt_attr = _QTextCharFormat()
        self._fmt_attr.setForeground(QColor("#aa5500"))
        self._fmt_str = _QTextCharFormat()
        self._fmt_str.setForeground(QColor("#228822"))

    def highlightBlock(self, text: str):
        # Tags
        for m in re.finditer(r"<[^>]+>", text):
            self.setFormat(m.start(), m.end() - m.start(), self._fmt_tag)
            inner = text[m.start():m.end()]
            # Attributes inside tag
            for a in re.finditer(r"\b([A-Za-z_:][-A-Za-z0-9_:.]*)\s*=\s*(\"[^\"]*\"|'[^']*')", inner):
                a_start = m.start() + a.start(1)
                a_len = len(a.group(1))
                self.setFormat(a_start, a_len, self._fmt_attr)
                s_start = m.start() + a.start(2)
                s_len = len(a.group(2))
                self.setFormat(s_start, s_len, self._fmt_str)


def _reapply_base_url(text_edit: QtWidgets.QTextEdit):
    try:
        media_root = getattr(text_edit.window(), "_media_root", None)
        if isinstance(media_root, str) and media_root:
            base = media_root if media_root.endswith(os.sep) else media_root + os.sep
            text_edit.document().setBaseUrl(QUrl.fromLocalFile(base))
    except Exception:
        pass


def _open_html_source_dialog(text_edit: QtWidgets.QTextEdit):
    dlg = QtWidgets.QDialog(text_edit)
    dlg.setWindowTitle("HTML Source")
    dlg.resize(800, 600)
    v = QtWidgets.QVBoxLayout(dlg)
    edit = QtWidgets.QPlainTextEdit(dlg)
    try:
        f = edit.font(); f.setFamily("Consolas"); f.setPointSize(10); edit.setFont(f)
    except Exception:
        pass
    # Load current HTML
    try:
        html = text_edit.document().toHtml()
    except Exception:
        html = ""
    edit.setPlainText(html)
    # Syntax highlighting
    try:
        _HtmlHighlighter(edit.document())
    except Exception:
        pass
    v.addWidget(edit)
    btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, parent=dlg)
    v.addWidget(btns)

    def _apply():
        new_html = edit.toPlainText()
        try:
            text_edit.setHtml(new_html)
        except Exception:
            pass
        _reapply_base_url(text_edit)
        dlg.accept()

    btns.accepted.connect(_apply)
    btns.rejected.connect(dlg.reject)
    dlg.exec_()
class _ImagePropertiesDialog(QtWidgets.QDialog):
    def __init__(self, parent, src_name: str, current_w: float, current_h: float, intrinsic_w: int, intrinsic_h: int, current_align: str):
        super().__init__(parent)
        self.setWindowTitle("Image Properties")
        layout = QtWidgets.QFormLayout(self)
        self.sp_width = QtWidgets.QDoubleSpinBox(self)
        self.sp_width.setRange(16.0, 10000.0)
        self.sp_width.setDecimals(1)
        self.sp_width.setValue(float(current_w or intrinsic_w or 400))
        self.cb_keep = QtWidgets.QCheckBox("Keep aspect ratio", self)
        self.cb_keep.setChecked(True)
        self.lbl_height = QtWidgets.QLabel(self)
        # Alt / Title
        self.le_alt = QtWidgets.QLineEdit(self)
        try:
            base = os.path.basename(src_name) if src_name else ""
        except Exception:
            base = ""
        self.le_alt.setText(base)
        self.le_title = QtWidgets.QLineEdit(self)
        # Alignment
        self.combo_align = QtWidgets.QComboBox(self)
        self.combo_align.addItems(["None", "Left", "Center", "Right"])
        try:
            idx = {"none":0, "left":1, "center":2, "right":3}.get((current_align or "none").lower(), 0)
            self.combo_align.setCurrentIndex(idx)
        except Exception:
            pass
        # Compute height preview
        self._iw = float(intrinsic_w or 0)
        self._ih = float(intrinsic_h or 0)
        def _refresh_h():
            w = self.sp_width.value()
            if self.cb_keep.isChecked():
                if self._iw > 0 and self._ih > 0:
                    h = w * (self._ih / self._iw)
                    self.lbl_height.setText(f"Height: {h:.0f} px (auto)")
                else:
                    self.lbl_height.setText("Height: (auto)")
            else:
                self.lbl_height.setText("")
        self.sp_width.valueChanged.connect(lambda _v: _refresh_h())
        self.cb_keep.toggled.connect(lambda _v: _refresh_h())
        _refresh_h()
        layout.addRow("Width (px):", self.sp_width)
        layout.addRow("", self.cb_keep)
        layout.addRow("", self.lbl_height)
        layout.addRow("Alt:", self.le_alt)
        layout.addRow("Title:", self.le_title)
        layout.addRow("Alignment:", self.combo_align)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, parent=self)
        layout.addRow(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
    def values(self):
        w = float(self.sp_width.value())
        align_idx = self.combo_align.currentIndex()
        align = {0:"none",1:"left",2:"center",3:"right"}.get(align_idx, "none")
        return w, align, self.cb_keep.isChecked(), self.le_alt.text().strip(), self.le_title.text().strip()


def _apply_block_alignment_for_image(text_edit: QtWidgets.QTextEdit, cursor_pos: int, align: str):
    try:
        c = QTextCursor(text_edit.document())
        c.setPosition(int(cursor_pos))
        blk = c.block()
        bf = blk.blockFormat()
        a = align.lower() if isinstance(align, str) else "none"
        if a == "none":
            return  # leave as-is
        if a == "left":
            bf.setAlignment(Qt.AlignLeft)
        elif a == "center":
            bf.setAlignment(Qt.AlignHCenter)
        elif a == "right":
            bf.setAlignment(Qt.AlignRight)
        else:
            bf.setAlignment(Qt.AlignLeft)
        c.setBlockFormat(bf)
    except Exception:
        pass


def _image_properties_dialog_apply(text_edit: QtWidgets.QTextEdit, info: dict):
    # Determine current alignment from block
    try:
        cur = QTextCursor(text_edit.document())
        cur.setPosition(int(info.get("cursor_pos", 0)))
        blk = cur.block()
        current_align = "center" if blk.blockFormat().alignment() & Qt.AlignHCenter else ("right" if blk.blockFormat().alignment() & Qt.AlignRight else "left")
    except Exception:
        current_align = "left"
    dlg = _ImagePropertiesDialog(text_edit, info.get("name"), info.get("w"), info.get("h"), info.get("iw") or 0, info.get("ih") or 0, current_align)  # (Removed legacy _table_recalculate_formulas after formula feature rollback.)
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return
    new_w, new_align, keep, alt_txt, title_txt = dlg.values()
    # Compute new height if keeping ratio
    iw = info.get("iw")
    ih = info.get("ih")
    cur_w = float(info.get("w") or (iw or 0.0))
    cur_h = float(info.get("h") or (ih or 0.0))
    if keep and iw and ih and iw > 0:
        new_h = max(1.0, float(new_w) * (float(ih) / float(iw)))
    elif keep and cur_w and cur_h and cur_w > 0:
        new_h = max(1.0, float(new_w) * (float(cur_h) / float(cur_w)))
    else:
        # If not keeping ratio, fall back to proportional based on current dims
        ratio = (float(cur_h) / float(cur_w)) if cur_w else 1.0
        new_h = max(1.0, float(new_w) * ratio)
    _apply_image_properties(text_edit, info["cursor_pos"], info["name"], float(new_w), float(new_h), alt_txt, title_txt)
    if (new_align or "none").lower() != "none":
        _apply_block_alignment_for_image(text_edit, info["cursor_pos"], new_align)


def _html_escape(s: str) -> str:
    try:
        s = s.replace("&", "&amp;")
        s = s.replace("\"", "&quot;")
        s = s.replace("<", "&lt;")
        s = s.replace(">", "&gt;")
        return s
    except Exception:
        return s


def _apply_image_properties(text_edit: QtWidgets.QTextEdit, cursor_pos: int, name: str, w: float, h: float, alt_txt: str, title_txt: str):
    try:
        doc = text_edit.document()
        c = QTextCursor(doc)
        c.setPosition(int(cursor_pos))
        # Select object replacement char if present
        c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
        try:
            c.removeSelectedText()
        except Exception:
            pass
        # Build HTML img tag with attributes
        src_attr = f'src="{_html_escape(name)}"' if name else ""
        w_attr = f' width="{int(max(1.0, w))}"' if w else ""
        h_attr = f' height="{int(max(1.0, h))}"' if h else ""
        alt_attr = f' alt="{_html_escape(alt_txt)}"' if alt_txt else ""
        title_attr = f' title="{_html_escape(title_txt)}"' if title_txt else ""
        html = f'<img {src_attr}{w_attr}{h_attr}{(" " + alt_attr) if alt_attr else ""}{(" " + title_attr) if title_attr else ""} />'
        c.insertHtml(html)
    except Exception:
        # Fallback to size-only application
        _apply_image_size_at(text_edit, cursor_pos, name, w, h)
import os
import tempfile
import imghdr

# Safe checks for deleted Qt objects (helps prevent native crashes)
def _is_alive(obj) -> bool:
    try:
        if obj is None:
            return False
        # Try to use sip.isdeleted if available, but don't hard-require sip
        try:
            import sip  # type: ignore
            try:
                return not sip.isdeleted(obj)
            except Exception:
                pass
        except Exception:
            # sip not available; assume object is alive
            pass
        return True
    except Exception:
        return False

# Feature flag: allow disabling image resize overlay via environment for diagnostics
def _is_image_resize_enabled() -> bool:
    """Feature flag for experimental image-resize overlay.

    Default: OFF (opt-in). Enable by setting NOTEBOOK_ENABLE_IMAGE_RESIZE=1.
    You can still force-disable with NOTEBOOK_DISABLE_IMAGE_RESIZE=1.
    """
    try:
        # Hard disable has priority
        v_disable = os.environ.get("NOTEBOOK_DISABLE_IMAGE_RESIZE", "0").strip().lower()
        if v_disable in ("1", "true", "yes"):
            return False
        # Opt-in enable
        v_enable = os.environ.get("NOTEBOOK_ENABLE_IMAGE_RESIZE", "0").strip().lower()
        return v_enable in ("1", "true", "yes")
    except Exception:
        return False


def _ensure_layout(widget: QtWidgets.QWidget) -> QtWidgets.QVBoxLayout:
    layout = widget.layout()
    if isinstance(layout, QtWidgets.QVBoxLayout):
        return layout
    layout = QtWidgets.QVBoxLayout(widget)
    widget.setLayout(layout)
    return layout


# Defaults you can change
DEFAULT_FONT_FAMILY = "Arial"  # e.g., "Arial", "Calibri", "Times New Roman"
DEFAULT_FONT_SIZE_PT = 12  # in points
DEFAULT_IMAGE_LONG_SIDE = 400  # px; long side target when inserting images
DEFAULT_VIDEO_LONG_SIDE = 400  # px; default long side for video thumbnails (can differ via settings)

# List scheme configuration (can be changed at runtime from main menu)
_ORDERED_SCHEME = "classic"  # 'classic' or 'decimal'
_UNORDERED_SCHEME = "disc-circle-square"  # 'disc-circle-square' or 'disc-only'


def _make_icon(kind: str, size: QSize = QSize(24, 24), fg: QColor = QColor("#303030")) -> QIcon:
    pm = QPixmap(size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(fg)
    pen.setWidth(2)
    p.setPen(pen)
    w, h = size.width(), size.height()

    def _draw_text(txt: str, bold=False, italic=False, underline=False, strike=False):
        f = QFont()
        f.setBold(bold)
        f.setItalic(italic)
        f.setPointSizeF(max(8.0, h * 0.48))
        p.setFont(f)
        metrics = QtWidgets.QStyleOption().fontMetrics
        # Fallback: draw centered
        rect = QRect(2, 2, w - 4, h - 4)
        p.drawText(rect, Qt.AlignCenter, txt)
        if underline:
            y = int(h * 0.72)
            p.drawLine(int(w * 0.25), y, int(w * 0.75), y)
        if strike:
            y = int(h * 0.5)
            p.drawLine(int(w * 0.25), y, int(w * 0.75), y)

    def _draw_align(mode: str):
        y = 5
        lines = [
            (0.9, 0.2),
            (0.7, 0.2),
            (0.9, 0.2),
            (0.6, 0.2),
        ]
        for i, (lw, pad) in enumerate(lines):
            yy = y + i * int(h * 0.18)
            if mode == "left":
                x1, x2 = 3, 3 + int((w - 6) * lw)
            elif mode == "center":
                span = int((w - 6) * lw)
                x1 = (w - span) // 2
                x2 = x1 + span
            elif mode == "right":
                span = int((w - 6) * lw)
                x2 = w - 3
                x1 = x2 - span
            else:  # justify
                x1, x2 = 3, w - 3
            p.drawLine(x1, yy, x2, yy)

    def _draw_list(bullets: bool):
        y = 5
        for i in range(3):
            yy = y + i * int(h * 0.24)
            if bullets:
                p.drawEllipse(QPoint(6, yy), 2, 2)
            else:
                # tiny '1.' like marker
                p.drawText(
                    QRect(2, yy - 6, 10, 12), Qt.AlignLeft | Qt.AlignVCenter, str(i + 1) + "."
                )
            p.drawLine(12, yy, w - 4, yy)

    if kind == "undo":
        # left arrow
        p.drawLine(int(w * 0.75), int(h * 0.3), int(w * 0.35), int(h * 0.3))
        p.drawLine(int(w * 0.35), int(h * 0.3), int(w * 0.45), int(h * 0.2))
        p.drawLine(int(w * 0.35), int(h * 0.3), int(w * 0.45), int(h * 0.4))
    elif kind == "redo":
        p.drawLine(int(w * 0.25), int(h * 0.3), int(w * 0.65), int(h * 0.3))
        p.drawLine(int(w * 0.65), int(h * 0.3), int(w * 0.55), int(h * 0.2))
        p.drawLine(int(w * 0.65), int(h * 0.3), int(w * 0.55), int(h * 0.4))
    elif kind == "bold":
        _draw_text("B", bold=True)
    elif kind == "italic":
        _draw_text("I", italic=True)
    elif kind == "underline":
        _draw_text("U", underline=True)
    elif kind == "strike":
        _draw_text("S", strike=True)
    elif kind == "align_left":
        _draw_align("left")
    elif kind == "align_center":
        _draw_align("center")
    elif kind == "align_right":
        _draw_align("right")
    elif kind == "align_justify":
        _draw_align("justify")
    elif kind == "list_bullets":
        _draw_list(True)
    elif kind == "list_numbers":
        _draw_list(False)
    elif kind == "indent":
        # right pointing arrow/step
        p.drawLine(4, h // 2, w - 8, h // 2)
        p.drawLine(w - 10, h // 2 - 6, w - 4, h // 2)
        p.drawLine(w - 10, h // 2 + 6, w - 4, h // 2)
        p.drawLine(4, 6, 4, h - 6)
    elif kind == "outdent":
        # left pointing arrow/step
        p.drawLine(w - 4, h // 2, 8, h // 2)
        p.drawLine(10, h // 2 - 6, 4, h // 2)
        p.drawLine(10, h // 2 + 6, 4, h // 2)
        p.drawLine(w - 4, 6, w - 4, h - 6)
    elif kind == "hr":
        y = h // 2
        p.drawLine(4, y, w - 4, y)
    elif kind == "image":
        # simple picture: frame + mountain + sun
        p.drawRect(3, 5, w - 6, h - 10)
        p.drawLine(6, h - 8, w - 10, h - 14)
        p.drawEllipse(QPoint(w - 10, 9), 2, 2)
    elif kind == "table":
        # 3x3 grid
        p.drawRect(3, 5, w - 6, h - 10)
        for i in range(1, 3):
            x = 3 + i * (w - 6) // 3
            p.drawLine(x, 5, x, h - 5)
            y = 5 + i * (h - 10) // 3
            p.drawLine(3, y, w - 3, y)
    elif kind == "video":
        # simple 'play' triangle inside a rounded rectangle
        rect = QRect(3, 5, w - 6, h - 10)
        p.drawRoundedRect(rect, 3, 3)
        px = rect.center().x()
        py = rect.center().y()
        size = int(min(rect.width(), rect.height()) * 0.5)
        pts = [QPoint(px - size // 3, py - size // 2), QPoint(px - size // 3, py + size // 2), QPoint(px + size // 2, py)]
        p.setBrush(QBrush(QColor(60, 60, 60)))
        p.setPen(Qt.NoPen)
        p.drawPolygon(*pts)
    elif kind == "code":
        # simple HTML/code glyph: </>
        pen2 = QPen(fg)
        pen2.setWidth(2)
        p.setPen(pen2)
        # Draw '<' and '>' as angled lines and a slash in between
        y_mid = h // 2
        left_x = int(w * 0.22)
        right_x = int(w * 0.78)
        span_y = int(h * 0.22)
        # '<'
        p.drawLine(left_x + 8, y_mid - span_y, left_x, y_mid)
        p.drawLine(left_x, y_mid, left_x + 8, y_mid + span_y)
        # '/'
        p.drawLine(w // 2 - 3, y_mid + span_y, w // 2 + 3, y_mid - span_y)
        # '>'
        p.drawLine(right_x - 8, y_mid - span_y, right_x, y_mid)
        p.drawLine(right_x, y_mid, right_x - 8, y_mid + span_y)
    elif kind == "color":
        _draw_text("A")
        p.drawRect(5, h - 8, w - 10, 4)
    elif kind == "highlight":
        p.fillRect(QRect(4, h - 10, w - 8, 6), QColor("#ffe680"))
        _draw_text("A")

    p.end()
    return QIcon(pm)


def add_rich_text_toolbar(
    parent_tab: QtWidgets.QWidget,
    text_edit: QtWidgets.QTextEdit,
    before_widget: QtWidgets.QWidget = None,
):
    if text_edit is None or parent_tab is None:
        return None
    layout = _ensure_layout(parent_tab)
    toolbar = QtWidgets.QToolBar(parent_tab)
    # Visual polish: icon-only toolbar, compact icons
    toolbar.setIconSize(QSize(20, 20))
    toolbar.setStyleSheet(
        """
        QToolBar {
            background: #f6f6f6;
            border-bottom: 1px solid #d0d0d0;
            spacing: 2px;
        }
        QToolButton {
            padding: 1px 3px;
        }
        QToolButton:checked {
            background: #e9f3ff;
            border: 1px solid #8ab9ff;
            border-radius: 3px;
        }
        QToolButton::menu-indicator { image: none; }
        """
    )
    toolbar.setToolButtonStyle(Qt.ToolButtonIconOnly)

    # Undo/Redo
    act_undo = toolbar.addAction(_make_icon("undo"), "", text_edit.undo)
    act_undo.setShortcut(QKeySequence.Undo)
    act_undo.setToolTip("Undo (Ctrl+Z)")
    act_redo = toolbar.addAction(_make_icon("redo"), "", text_edit.redo)
    act_redo.setShortcut(QKeySequence.Redo)
    act_redo.setToolTip("Redo (Ctrl+Y)")
    toolbar.addSeparator()

    # Bold/Italic/Underline/Strike
    def toggle_format(flag_attr: str, on: bool):
        fmt = QTextCharFormat()
        if flag_attr == "bold":
            fmt.setFontWeight(QFont.Bold if on else QFont.Normal)
        elif flag_attr == "italic":
            fmt.setFontItalic(on)
        elif flag_attr == "underline":
            fmt.setFontUnderline(on)
        elif flag_attr == "strike":
            fmt.setFontStrikeOut(on)
        cursor = text_edit.textCursor()
        if not cursor.hasSelection():
            # Apply to current word/cursor moving forward
            cursor.select(cursor.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        text_edit.mergeCurrentCharFormat(fmt)

    act_bold = QtWidgets.QAction(_make_icon("bold"), "", toolbar)
    act_bold.setCheckable(True)
    act_bold.setShortcut(QKeySequence.Bold)
    act_bold.setToolTip("Bold (Ctrl+B)")
    act_bold.triggered.connect(lambda on: toggle_format("bold", on))
    toolbar.addAction(act_bold)

    act_italic = QtWidgets.QAction(_make_icon("italic"), "", toolbar)
    act_italic.setCheckable(True)
    act_italic.setShortcut(QKeySequence.Italic)
    act_italic.setToolTip("Italic (Ctrl+I)")
    act_italic.triggered.connect(lambda on: toggle_format("italic", on))
    toolbar.addAction(act_italic)

    act_underline = QtWidgets.QAction(_make_icon("underline"), "", toolbar)
    act_underline.setCheckable(True)
    act_underline.setShortcut(QKeySequence.Underline)
    act_underline.setToolTip("Underline (Ctrl+U)")
    act_underline.triggered.connect(lambda on: toggle_format("underline", on))
    toolbar.addAction(act_underline)

    act_strike = QtWidgets.QAction(_make_icon("strike"), "", toolbar)
    act_strike.setCheckable(True)
    act_strike.setToolTip("Strikethrough")
    act_strike.triggered.connect(lambda on: toggle_format("strike", on))
    toolbar.addAction(act_strike)

    toolbar.addSeparator()

    # Font family and size
    font_box = QtWidgets.QFontComboBox(toolbar)
    # Make the font family control more compact horizontally
    try:
        font_box.setMaximumWidth(140)
        font_box.setMinimumContentsLength(8)
    except Exception:
        pass
    font_box.currentFontChanged.connect(lambda f: _apply_font_family(text_edit, f.family()))
    font_box.setToolTip("Font family")
    toolbar.addWidget(font_box)

    size_box = QtWidgets.QComboBox(toolbar)
    for sz in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32]:
        size_box.addItem(str(sz), sz)
    size_box.setEditable(False)
    try:
        size_box.setMaximumWidth(72)
    except Exception:
        pass
    size_box.currentIndexChanged.connect(
        lambda _i: _apply_font_size(text_edit, size_box.currentData())
    )
    size_box.setToolTip("Font size")
    toolbar.addWidget(size_box)

    # Apply default font family and size for new content (does not overwrite existing styled HTML)
    try:
        text_edit.document().setDefaultFont(QFont(DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT))
        # Reflect defaults in the pickers
        font_box.setCurrentFont(QFont(DEFAULT_FONT_FAMILY))
        # Ensure the size combo selects the default size if present
        for i in range(size_box.count()):
            if int(size_box.itemData(i)) == int(DEFAULT_FONT_SIZE_PT):
                size_box.setCurrentIndex(i)
                break
    except Exception:
        pass

    toolbar.addSeparator()

    # Text color & highlight
    btn_color = QtWidgets.QToolButton(toolbar)
    btn_color.setText("A")
    btn_color.setToolTip("Text color")
    btn_color.clicked.connect(lambda: _apply_text_color(text_edit, foreground=True))
    try:
        btn_color.setFixedSize(24, 24)
    except Exception:
        pass
    toolbar.addWidget(btn_color)

    btn_bg = QtWidgets.QToolButton(toolbar)
    btn_bg.setText("Bg")
    btn_bg.setToolTip("Highlight")
    btn_bg.clicked.connect(lambda: _apply_text_color(text_edit, foreground=False))
    try:
        btn_bg.setFixedSize(24, 24)
    except Exception:
        pass
    toolbar.addWidget(btn_bg)

    # Clear only background highlight (keep bold/italic/etc.)
    btn_bg_clear = QtWidgets.QToolButton(toolbar)
    btn_bg_clear.setText("NoBg")
    btn_bg_clear.setToolTip("Remove highlight (background)")
    btn_bg_clear.clicked.connect(lambda: _clear_background(text_edit))
    try:
        btn_bg_clear.setFixedSize(28, 24)
    except Exception:
        pass
    toolbar.addWidget(btn_bg_clear)

    toolbar.addSeparator()

    # Alignment
    group_align = QtWidgets.QActionGroup(toolbar)
    group_align.setExclusive(True)
    act_align_left = QtWidgets.QAction(_make_icon("align_left"), "", toolbar, checkable=True)
    act_align_left.setToolTip("Align Left")
    act_align_center = QtWidgets.QAction(_make_icon("align_center"), "", toolbar, checkable=True)
    act_align_center.setToolTip("Align Center")
    act_align_right = QtWidgets.QAction(_make_icon("align_right"), "", toolbar, checkable=True)
    act_align_right.setToolTip("Align Right")
    act_align_justify = QtWidgets.QAction(_make_icon("align_justify"), "", toolbar, checkable=True)
    act_align_justify.setToolTip("Justify")
    for a in (act_align_left, act_align_center, act_align_right, act_align_justify):
        group_align.addAction(a)
        toolbar.addAction(a)
    act_align_left.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignLeft))
    act_align_center.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignHCenter))
    act_align_right.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignRight))
    act_align_justify.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignJustify))

    toolbar.addSeparator()

    # Lists
    act_bullets = toolbar.addAction(
        _make_icon("list_bullets"), "", lambda: _toggle_list(text_edit, ordered=False)
    )
    act_bullets.setToolTip("Bulleted list")
    act_numbers = toolbar.addAction(
        _make_icon("list_numbers"), "", lambda: _toggle_list(text_edit, ordered=True)
    )
    act_numbers.setToolTip("Numbered list")

    # Indent/Outdent for list nesting
    act_indent = toolbar.addAction(
        _make_icon("indent"), "", lambda: _change_list_indent(text_edit, +1)
    )
    act_indent.setShortcut(QKeySequence("Ctrl+]"))
    try:
        act_indent.setShortcutContext(Qt.WidgetWithChildrenShortcut)
    except Exception:
        pass
    act_indent.setToolTip("Indent (Tab, Ctrl+])")
    act_outdent = toolbar.addAction(
        _make_icon("outdent"), "", lambda: _change_list_indent(text_edit, -1)
    )
    act_outdent.setShortcut(QKeySequence("Ctrl["))
    try:
        act_outdent.setShortcutContext(Qt.WidgetWithChildrenShortcut)
    except Exception:
        pass
    act_outdent.setToolTip("Outdent (Shift+Tab, Ctrl+[)")

    # Enable Tab/Shift+Tab to control list levels
    _install_list_tab_handler(text_edit)
    # Enable table cell Tab navigation
    _install_table_tab_handler(text_edit)
    # Enable plain paragraph indent/outdent with Tab/Shift+Tab when not in lists/tables
    _install_plain_indent_tab_handler(text_edit)

    # Disable drag-and-drop into the editor per current requirements
    try:
        text_edit.setAcceptDrops(False)
        if hasattr(text_edit, "viewport") and text_edit.viewport() is not None:
            text_edit.viewport().setAcceptDrops(False)
    except Exception:
        pass

    toolbar.addSeparator()

    # Clear formatting, HR, Insert image/video
    act_clear = toolbar.addAction(
        _make_icon("color"), "", lambda: text_edit.setCurrentCharFormat(QTextCharFormat())
    )
    act_clear.setToolTip("Clear formatting")
    def _insert_horizontal_rule():
        cur = text_edit.textCursor()
        cur.beginEditBlock()
        try:
            # Collapse selection and move to block boundary so the rule is on its own line
            if cur.hasSelection():
                pos = max(cur.position(), cur.anchor())
                cur.setPosition(pos)
            if cur.positionInBlock() != 0:
                cur.insertBlock()
            # Build a thin 1x1 table spanning full width with a 1px black top border on the cell
            tf = QTextTableFormat()
            try:
                tf.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
            except Exception:
                pass
            tf.setCellPadding(0)
            tf.setCellSpacing(0)
            tf.setBorder(0)
            # Mark as HR so global border enforcement can skip it (runtime marker; HTML reload uses pattern detection)
            try:
                tf.setProperty(int(QTextFormat.UserProperty) + 101, True)
            except Exception:
                pass
            tbl = cur.insertTable(1, 1, tf)
            try:
                cf = QTextTableCellFormat()
                # Only top border to create a single visible line
                cf.setTopBorder(1.0)
                cf.setBottomBorder(0.0)
                cf.setLeftBorder(0.0)
                cf.setRightBorder(0.0)
                try:
                    cf.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)
                except Exception:
                    pass
                try:
                    cf.setBorderBrush(QBrush(QColor("#000000")))
                except Exception:
                    pass
                cell = tbl.cellAt(0, 0)
                cell.setFormat(cf)
            except Exception:
                pass
            # Move cursor after the table and insert a new block for continued typing
            after = QTextCursor(text_edit.document())
            try:
                after.setPosition(tbl.lastPosition())
            except Exception:
                pass
            text_edit.setTextCursor(after)
            text_edit.textCursor().insertBlock()
        finally:
            cur.endEditBlock()

    act_hr = toolbar.addAction(_make_icon("hr"), "", _insert_horizontal_rule)
    act_hr.setToolTip("Insert horizontal rule")
    # Image insert: split-button with dropdown for sizing modes
    btn_img = QtWidgets.QToolButton(toolbar)
    btn_img.setIcon(_make_icon("image"))
    btn_img.setToolTip("Insert image from file")
    # Open the menu on any click (no direct action)
    btn_img.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    img_menu = QtWidgets.QMenu(btn_img)
    try:
        act_def = img_menu.addAction(f"Use default ({int(DEFAULT_IMAGE_LONG_SIDE)} px)…")
    except Exception:
        act_def = img_menu.addAction("Use default size…")
    act_fit = img_menu.addAction("Fit to editor width…")
    act_orig = img_menu.addAction("Original size…")
    act_cust = img_menu.addAction("Custom width…")

    def _choose_custom_and_insert():
        try:
            w, ok = QtWidgets.QInputDialog.getInt(
                toolbar, "Insert Image", "Width (px):", int(DEFAULT_IMAGE_LONG_SIDE), 50, 8000, 10
            )
            if not ok:
                return
            _insert_image_via_dialog(text_edit, mode="custom", custom_width=float(w))
        except Exception:
            pass

    act_def.triggered.connect(lambda: _insert_image_via_dialog(text_edit, mode="default"))
    act_fit.triggered.connect(lambda: _insert_image_via_dialog(text_edit, mode="fit-width"))
    act_orig.triggered.connect(lambda: _insert_image_via_dialog(text_edit, mode="original"))
    act_cust.triggered.connect(_choose_custom_and_insert)
    btn_img.setMenu(img_menu)
    toolbar.addWidget(btn_img)
    # Video insert: split-button with options
    btn_vid = QtWidgets.QToolButton(toolbar)
    btn_vid.setIcon(_make_icon("video"))
    btn_vid.setToolTip("Insert video from file")
    btn_vid.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    vid_menu = QtWidgets.QMenu(btn_vid)
    # Sizing section (mirrors image sizing semantics)
    try:
        act_vsz_def = vid_menu.addAction(f"Size: Default ({int(DEFAULT_IMAGE_LONG_SIDE)} px long side)")
    except Exception:
        act_vsz_def = vid_menu.addAction("Size: Default")
    act_vsz_fit = vid_menu.addAction("Size: Fit editor width")
    act_vsz_orig = vid_menu.addAction("Size: Original")
    act_vsz_custom = vid_menu.addAction("Size: Custom width…")
    vid_menu.addSeparator()
    # Capture time section
    act_v1 = vid_menu.addAction("Frame at 1.0s…")
    act_v3 = vid_menu.addAction("Frame at 3.0s…")
    act_v5 = vid_menu.addAction("Frame at 5.0s…")
    act_vc = vid_menu.addAction("Custom time…")
    vid_menu.addSeparator()
    act_vs = vid_menu.addAction("Use synthetic placeholder…")

    def _choose_custom_video_time():
        try:
            secs, ok = QtWidgets.QInputDialog.getDouble(toolbar, "Insert Video", "Capture time (seconds):", 1.0, 0.0, 36000.0, 1)
            if not ok:
                return
            _insert_video_via_dialog(text_edit, capture_seconds=float(secs), force_synthetic=False)
        except Exception:
            pass

    # Track chosen sizing mode for subsequent inserts
    _video_size_state = {"mode": "default", "custom_width": None}

    def _set_video_size(mode: str, custom_w: float = None):
        _video_size_state["mode"] = mode
        _video_size_state["custom_width"] = custom_w

    act_vsz_def.triggered.connect(lambda: _set_video_size("default"))
    act_vsz_fit.triggered.connect(lambda: _set_video_size("fit-width"))
    act_vsz_orig.triggered.connect(lambda: _set_video_size("original"))

    def _choose_video_custom_width():
        try:
            w, ok = QtWidgets.QInputDialog.getInt(toolbar, "Video width", "Width (px):", int(DEFAULT_IMAGE_LONG_SIDE), 50, 8000, 10)
            if ok:
                _set_video_size("custom", float(w))
        except Exception:
            pass
    act_vsz_custom.triggered.connect(_choose_video_custom_width)

    act_v1.triggered.connect(lambda: _insert_video_via_dialog(text_edit, capture_seconds=1.0, force_synthetic=False, size_mode=_video_size_state["mode"], custom_width=_video_size_state["custom_width"]))
    act_v3.triggered.connect(lambda: _insert_video_via_dialog(text_edit, capture_seconds=3.0, force_synthetic=False, size_mode=_video_size_state["mode"], custom_width=_video_size_state["custom_width"]))
    act_v5.triggered.connect(lambda: _insert_video_via_dialog(text_edit, capture_seconds=5.0, force_synthetic=False, size_mode=_video_size_state["mode"], custom_width=_video_size_state["custom_width"]))
    act_vc.triggered.connect(_choose_custom_video_time)
    act_vs.triggered.connect(lambda: _insert_video_via_dialog(text_edit, capture_seconds=None, force_synthetic=True, size_mode=_video_size_state["mode"], custom_width=_video_size_state["custom_width"]))
    btn_vid.setMenu(vid_menu)
    toolbar.addWidget(btn_vid)

    # Paste Text Only quick action
    act_paste_plain = toolbar.addAction(_make_icon("color"), "", lambda: paste_text_only(text_edit))
    act_paste_plain.setToolTip("Paste Text Only (Ctrl+Shift+V)")

    # HTML Source editor
    act_html = toolbar.addAction(_make_icon("code"), "", lambda: _open_html_source_dialog(text_edit))
    act_html.setToolTip("HTML Source…")

    # Table: insert or edit if caret is inside a table (as an action so it participates in overflow menu)
    toolbar.addSeparator()
    act_table = QtWidgets.QAction(_make_icon("table"), "", toolbar)
    act_table.setToolTip("Insert/edit table")
    try:
        act_table.setPriority(QtWidgets.QAction.HighPriority)
    except Exception:
        pass
    act_table.triggered.connect(lambda: _table_insert_or_edit(text_edit))
    toolbar.addAction(act_table)

    # (Image Actions toolbar button removed by request)

    # Place toolbar in layout
    if before_widget is not None:
        # Insert before the specified widget if possible
        idx = layout.indexOf(before_widget)
        layout.insertWidget(max(0, idx), toolbar)
    else:
        layout.insertWidget(0, toolbar)

    # Reflect current formatting in toolbar toggles when cursor moves
    def _sync_toolbar():
        fmt = text_edit.currentCharFormat()
        act_bold.setChecked(fmt.fontWeight() == QFont.Bold)
        act_italic.setChecked(fmt.fontItalic())
        act_underline.setChecked(fmt.fontUnderline())
        act_strike.setChecked(fmt.fontStrikeOut())
        try:
            if fmt.fontPointSize() > 0:
                _select_combo_value(size_box, int(fmt.fontPointSize()))
            fam = fmt.fontFamily()
            if fam:
                font_box.setCurrentFont(QFont(fam))
        except Exception:
            pass

    text_edit.cursorPositionChanged.connect(_sync_toolbar)

    # Install image context menu and shortcuts
    try:
        _install_image_context_menu(text_edit)
        _install_image_shortcuts(text_edit)
    except Exception:
        pass
    try:
        text_edit.selectionChanged.connect(_sync_toolbar)
    except Exception:
        pass

    # Install Ctrl+V override to honor Default Paste Mode without relying on a window-level shortcut
    _install_default_paste_override(text_edit)
    # Enable Ctrl+Click to open links in the system browser
    _install_link_click_handler(text_edit)
    # Enable right-click table context menu
    _install_table_context_menu(text_edit)
    # Enable right-click image resize menu (stable alternative to overlay)
    # (Already installed above together with image shortcuts)
    # Disable drag-and-drop to avoid accidental inserts while we focus on stable HTML-style flows
    try:
        text_edit.setAcceptDrops(False)
        if text_edit.viewport() is not None:
            text_edit.viewport().setAcceptDrops(False)
    except Exception:
        pass
    # Enable mouse-based image resizing with aspect ratio preserved (can disable via env)
    if _is_image_resize_enabled():
        _install_image_resize_handler(text_edit)
    # Use a custom context menu handler so we can reliably detect images
    try:
        text_edit.setContextMenuPolicy(Qt.CustomContextMenu)
    except Exception:
        pass

    # Improve selection visibility: use a clearer highlight and text color
    try:
        _apply_selection_colors(text_edit, QColor("#4d84b7"), QColor("#000000"))
    except Exception:
        pass

    # Deterministic keyboard shortcut to open Image Properties for image at/near caret
    def _open_image_properties_from_caret():
        try:
            info = _image_info_from_selection(text_edit)
            if info is None:
                info = _image_info_at_cursor(text_edit)
            if info is None:
                cur = text_edit.textCursor()
                info = _image_info_near_doc_pos(text_edit, cur.position())
            if info is None:
                info = _image_info_in_block(text_edit, text_edit.textCursor().block(), prefer_pos=text_edit.textCursor().position())
            if info is None:
                # No image nearby, do nothing
                return
            _image_properties_dialog_apply(text_edit, info)
        except Exception:
            pass
    try:
        sc1 = QtWidgets.QShortcut(QKeySequence("Ctrl+Shift+I"), text_edit)
        sc1.setContext(Qt.WidgetWithChildrenShortcut)
        sc1.activated.connect(_open_image_properties_from_caret)
    except Exception:
        pass
    try:
        sc2 = QtWidgets.QShortcut(QKeySequence("F2"), text_edit)
        sc2.setContext(Qt.WidgetWithChildrenShortcut)
        sc2.activated.connect(_open_image_properties_from_caret)
    except Exception:
        pass

    # Install a small image HUD button that appears when caret is on/near an image
    try:
        _install_image_hud(text_edit)
    except Exception:
        pass
    return toolbar


def _effective_default_font(text_edit: QtWidgets.QTextEdit) -> QFont:
    try:
        df = text_edit.document().defaultFont()
        # Some platforms report 0 size; fallback to widget font
        if df.pointSizeF() <= 0 and df.pointSize() <= 0:
            df = text_edit.font()
    except Exception:
        df = text_edit.font()
    # Ensure a sensible point size
    try:
        sz = df.pointSizeF() if df.pointSizeF() > 0 else float(df.pointSize())
    except Exception:
        sz = 0.0
    if not sz or sz <= 0:
        df.setPointSizeF(float(DEFAULT_FONT_SIZE_PT))
    return df


def paste_text_only(text_edit: QtWidgets.QTextEdit):
    """Paste clipboard contents as plain text and normalize formatting to defaults (no bold/size/bg)."""
    cb = QtWidgets.QApplication.clipboard()
    md = cb.mimeData()
    if md is None:
        return
    if md.hasText():
        text = md.text()
    else:
        # Fallback: try plain text format explicitly
        text = (
            md.data("text/plain").data().decode("utf-8", errors="replace")
            if md.hasFormat("text/plain")
            else ""
        )
    if not text:
        return
    # Capture current format/cursor and create a neutral style based on current selection style
    cursor = text_edit.textCursor()
    old_fmt = text_edit.currentCharFormat()
    neutral = QTextCharFormat()
    # Base on current family/size so paste follows user-selected toolbar values
    # family/size fallback to document defaults if current format is unspecified
    doc_font = _effective_default_font(text_edit)
    fam = old_fmt.fontFamily() or doc_font.family()
    if fam:
        neutral.setFontFamily(fam)
    sz = old_fmt.fontPointSize()
    if not sz or sz <= 0:
        try:
            sz = doc_font.pointSizeF() if doc_font.pointSizeF() > 0 else float(doc_font.pointSize())
        except Exception:
            sz = 12.0
    if sz and sz > 0:
        neutral.setFontPointSize(float(sz))
    # Do not bring background; keep foreground consistent
    try:
        neutral.setForeground(text_edit.palette().text().color())
    except Exception:
        pass
    neutral.clearBackground()
    # Insert, then normalize the inserted range, then restore original typing format
    start = cursor.position()
    cursor.insertText(text)
    end = cursor.position()
    rng = QTextCursor(text_edit.document())
    rng.setPosition(start)
    rng.setPosition(end, QTextCursor.KeepAnchor)
    # ensure transparent background for inserted text
    neutral_bg = QTextCharFormat(neutral)
    neutral_bg.setBackground(Qt.transparent)
    rng.mergeCharFormat(neutral_bg)
    # Restore current typing format (clear background so it doesn't persist)
    restored = QTextCharFormat(old_fmt)
    restored.setBackground(Qt.transparent)
    text_edit.setCurrentCharFormat(restored)
    cursor.setPosition(end)
    text_edit.setTextCursor(cursor)


def paste_match_style(text_edit: QtWidgets.QTextEdit):
    """Paste as rich text but normalize to current char format (font family/size/color)."""
    cb = QtWidgets.QApplication.clipboard()
    md = cb.mimeData()
    if md is None:
        return
    # Prefer HTML if available; otherwise fallback to plain text
    if md.hasHtml():
        html = md.html()
        try:
            html = _strip_match_style_html(html)
        except Exception:
            pass
    elif md.hasText():
        txt = md.text().strip()
        # If the clipboard contains a raw URL, make it an anchor so it remains clickable
        if _looks_like_url(txt):
            url = _normalize_url_scheme(txt)
            html = f'<a href="{url}">{txt}</a>'
        else:
            html = txt.replace("\n", "<br/>")
    else:
        return
    cursor = text_edit.textCursor()
    # Preserve current format so toolbar stays unchanged
    pre_fmt = text_edit.currentCharFormat()
    # Insert the HTML, then apply a sanitized current format to the inserted fragment
    before = cursor.position()
    cursor.insertHtml(html)
    after = cursor.position()
    rng = QTextCursor(text_edit.document())
    rng.setPosition(before)
    rng.setPosition(after, QTextCursor.KeepAnchor)
    # Build a normalized format with explicit family/size fallback to document defaults
    doc_font = _effective_default_font(text_edit)
    fmt = QTextCharFormat(pre_fmt)
    fam = fmt.fontFamily() or doc_font.family()
    if fam:
        fmt.setFontFamily(fam)
    sz = fmt.fontPointSize()
    if not sz or sz <= 0:
        try:
            sz = doc_font.pointSizeF() if doc_font.pointSizeF() > 0 else float(doc_font.pointSize())
        except Exception:
            sz = 12.0
    fmt.setFontPointSize(float(sz))
    # Ensure background doesn't carry over
    try:
        fmt.setBackground(Qt.transparent)
    except Exception:
        fmt.clearBackground()
    rng.mergeCharFormat(fmt)
    # Restore typing format and place caret at end
    restored = QTextCharFormat(pre_fmt)
    try:
        restored.setBackground(Qt.transparent)
    except Exception:
        restored.clearBackground()
    text_edit.setCurrentCharFormat(restored)
    cursor.setPosition(after)
    text_edit.setTextCursor(cursor)


def _image_info_at_position(text_edit: QtWidgets.QTextEdit, pos_or_posint):
    """Detect image info at a specific document position (int) or viewport QPoint.
    Returns {cursor_pos,name,w,h,iw,ih} or None.
    """
    try:
        # Map from viewport QPoint to document position if needed
        if isinstance(pos_or_posint, QPoint):
            cur = text_edit.cursorForPosition(pos_or_posint)
        else:
            cur = QTextCursor(text_edit.document())
            cur.setPosition(int(pos_or_posint))

        def _mk(imgf: QTextImageFormat, c: QTextCursor):
            try:
                name = imgf.name() if hasattr(imgf, "name") else imgf.property(QTextFormat.ImageName)
            except Exception:
                name = ""
            w = float(imgf.width() or 0.0)
            h = float(imgf.height() or 0.0)
            iw = ih = None
            try:
                base = getattr(text_edit.window(), "_media_root", None)
                path = name
                if base and name and not os.path.isabs(name):
                    path = os.path.join(base, name)
                if not _is_raw_ext(path):
                    img = QImage(path)
                    if not img.isNull():
                        iw, ih = img.width(), img.height()
            except Exception:
                pass
            return {"cursor_pos": c.position(), "name": name, "w": w, "h": h, "iw": iw, "ih": ih}

        # Check current, previous, and next positions
        candidates = []
        for offset in (0, -1, +1):
            c = QTextCursor(cur)
            if offset < 0:
                c.movePosition(QTextCursor.Left)
            elif offset > 0:
                c.movePosition(QTextCursor.Right)
            candidates.append(c)
        for c in candidates:
            fmt = c.charFormat()
            if fmt is not None and (
                (hasattr(fmt, "isImageFormat") and fmt.isImageFormat())
                or fmt.objectType() == QTextFormat.ImageObject
            ):
                return _mk(QTextImageFormat(fmt), QTextCursor(c))
            # Try selecting the char at this position to fetch the actual object format
            try:
                csel = QTextCursor(c)
                csel.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                fmt2 = csel.charFormat()
                if fmt2 is not None and (
                    (hasattr(fmt2, "isImageFormat") and fmt2.isImageFormat())
                    or fmt2.objectType() == QTextFormat.ImageObject
                ):
                    return _mk(QTextImageFormat(fmt2), QTextCursor(c))
            except Exception:
                pass
    except Exception:
        return None
    return None


def _image_info_at_cursor(text_edit: QtWidgets.QTextEdit):
    """Robustly detect an image under/near the caret or selection start.
    Strategy:
    1) Check current pos and neighbors
    2) Check selection start and neighbors
    3) Scan left/right within a small radius of positions
    4) Sample viewport points around the caret rect center
    """
    try:
        cur = text_edit.textCursor()
        # Try current pos and neighbors
        info = _image_info_at_position(text_edit, cur.position())
        if info:
            return info
        # Try selection anchor/start if available
        try:
            start_pos = min(cur.anchor(), cur.position())
        except Exception:
            start_pos = cur.position()
        if start_pos != cur.position():
            info = _image_info_at_position(text_edit, start_pos)
            if info:
                return info
            info = _image_info_at_position(text_edit, max(0, start_pos - 1))
            if info:
                return info
        # Scan positions around caret within a small radius
        try:
            p0 = cur.position()
            for d in (1, 2, 3, 4, 6, 8, 12, 16, 24, 32):
                for pos_try in (max(0, p0 - d), p0 + d):
                    info = _image_info_at_position(text_edit, pos_try)
                    if info:
                        return info
        except Exception:
            pass
        # Sample viewport around caret rect center
        try:
            r = text_edit.cursorRect(cur)
            if r is not None and r.isValid():
                vp = text_edit.viewport()
                cx = r.center().x()
                cy = r.center().y()
                # Sweep horizontally and slightly vertically
                for dy in (0, int(r.height()/3) if r.height() > 0 else 6, -int(r.height()/3) if r.height() > 0 else -6):
                    for dx in (-64, -48, -32, -24, -16, -8, -4, 0, 4, 8, 16, 24, 32, 48, 64):
                        pt = QPoint(int(cx + dx), int(cy + dy))
                        # Clamp within viewport
                        if vp is not None:
                            rr = vp.rect()
                            if not rr.contains(pt):
                                continue
                        info = _image_info_at_position(text_edit, pt)
                        if info:
                            return info
        except Exception:
            pass
    except Exception:
        return None
    return None


def _image_info_near_doc_pos(text_edit: QtWidgets.QTextEdit, doc_pos: int, max_scan: int = 512):
    """Find image nearest to a given document position, scanning within the current block first.
    Returns {cursor_pos,name,w,h,iw,ih} or None.
    """
    try:
        doc = text_edit.document()
        cur = QTextCursor(doc)
        cur.setPosition(int(doc_pos))
        blk = cur.block()
        # Limit search to the paragraph/block for precision
        start = blk.position()
        end = blk.position() + blk.length() - 1
        # Scan outward from doc_pos within [start,end]
        def _try_at(p):
            if p < start or p > end:
                return None
            c = QTextCursor(doc)
            c.setPosition(p)
            # Select the char at this position to retrieve object format
            c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            fmt = c.charFormat()
            if fmt is not None and ((hasattr(fmt, "isImageFormat") and fmt.isImageFormat()) or fmt.objectType() == QTextFormat.ImageObject):
                imgf = QTextImageFormat(fmt)
                name = imgf.name() if hasattr(imgf, "name") else imgf.property(QTextFormat.ImageName)
                w = float(imgf.width() or 0.0)
                h = float(imgf.height() or 0.0)
                # Try to get intrinsic size
                iw = ih = None
                try:
                    base = getattr(text_edit.window(), "_media_root", None)
                    path = name
                    if base and name and not os.path.isabs(name):
                        path = os.path.join(base, name)
                    img = QImage(path)
                    if not img.isNull():
                        iw, ih = img.width(), img.height()
                except Exception:
                    pass
                return {"cursor_pos": p, "name": name, "w": w, "h": h, "iw": iw, "ih": ih}
            return None
        # Check exact pos and neighbors out to max_scan (capped by block length)
        max_delta = int(min(max_scan, max(0, end - start)))
        # Prefer exact, then +/-1, +/-2, ...
        for d in range(0, max_delta + 1):
            # exact
            if d == 0:
                info = _try_at(doc_pos)
                if info:
                    return info
                continue
            # left then right
            info = _try_at(doc_pos - d)
            if info:
                return info
            info = _try_at(doc_pos + d)
            if info:
                return info
    except Exception:
        return None
    return None


def _image_info_from_selection(text_edit: QtWidgets.QTextEdit):
    """If there's a selection, scan the selected range for an image object and return its info."""
    try:
        cur = text_edit.textCursor()
        a = cur.anchor()
        p = cur.position()
        start = min(a, p)
        end = max(a, p)
        if start == end:
            return None
        doc = text_edit.document()
        pos = start
        while pos <= end:
            c = QTextCursor(doc)
            c.setPosition(pos)
            c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            fmt = c.charFormat()
            if fmt is not None and (
                (hasattr(fmt, "isImageFormat") and fmt.isImageFormat())
                or fmt.objectType() == QTextFormat.ImageObject
            ):
                imgf = QTextImageFormat(fmt)
                try:
                    name = imgf.name() if hasattr(imgf, "name") else imgf.property(QTextFormat.ImageName)
                except Exception:
                    name = ""
                w = float(imgf.width() or 0.0)
                h = float(imgf.height() or 0.0)
                iw = ih = None
                try:
                    base = getattr(text_edit.window(), "_media_root", None)
                    path = name
                    if base and name and not os.path.isabs(name):
                        path = os.path.join(base, name)
                    img = QImage(path)
                    if not img.isNull():
                        iw, ih = img.width(), img.height()
                except Exception:
                    pass
                return {"cursor_pos": c.position(), "name": name, "w": w, "h": h, "iw": iw, "ih": ih}
            pos += 1
    except Exception:
        return None
    return None


def _image_info_in_block(text_edit: QtWidgets.QTextEdit, block, prefer_pos: int = None):
    """Scan QTextBlock fragments to find an embedded image. If multiple, pick the one nearest prefer_pos.
    Returns {cursor_pos,name,w,h,iw,ih} or None.
    """
    try:
        # Collect all image fragments with their positions
        imgs = []
        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid():
                fmt = frag.charFormat()
                if fmt is not None and ((hasattr(fmt, "isImageFormat") and fmt.isImageFormat()) or fmt.objectType() == QTextFormat.ImageObject):
                    imgf = QTextImageFormat(fmt)
                    try:
                        name = imgf.name() if hasattr(imgf, "name") else imgf.property(QTextFormat.ImageName)
                    except Exception:
                        name = ""
                    w = float(imgf.width() or 0.0)
                    h = float(imgf.height() or 0.0)
                    iw = ih = None
                    try:
                        base = getattr(text_edit.window(), "_media_root", None)
                        path = name
                        if base and name and not os.path.isabs(name):
                            path = os.path.join(base, name)
                        img = QImage(path)
                        if not img.isNull():
                            iw, ih = img.width(), img.height()
                    except Exception:
                        pass
                    imgs.append({
                        "cursor_pos": frag.position(),
                        "name": name,
                        "w": w,
                        "h": h,
                        "iw": iw,
                        "ih": ih,
                    })
            it += 1
        if not imgs:
            return None
        if prefer_pos is None:
            return imgs[0]
        # Choose the image with position closest to prefer_pos
        imgs.sort(key=lambda d: abs(int(d.get("cursor_pos", 0)) - int(prefer_pos)))
        return imgs[0]
    except Exception:
        return None


def _apply_image_size_at(text_edit: QtWidgets.QTextEdit, cursor_pos: int, name: str, w: float, h: float):
    try:
        doc = text_edit.document()
        # Ensure we target the actual image character; adjust if needed
        info_here = _image_info_at_position(text_edit, cursor_pos)
        if info_here is None:
            info_here = _image_info_at_position(text_edit, max(0, cursor_pos - 1))
        if info_here is not None:
            target_pos = int(info_here.get("cursor_pos", cursor_pos))
        else:
            target_pos = int(cursor_pos)

        # Preferred: replace the object with a new image fragment at the desired size
        try:
            c = QTextCursor(doc)
            c.setPosition(target_pos)
            # Select the object replacement char if present
            c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            # If selection is not an image, fall back later
            fmt = c.charFormat()
            is_img_sel = fmt is not None and (
                (hasattr(fmt, "isImageFormat") and fmt.isImageFormat()) or fmt.objectType() == QTextFormat.ImageObject
            )
            if is_img_sel:
                try:
                    c.removeSelectedText()
                except Exception:
                    pass
                imgf = QTextImageFormat()
                if name:
                    imgf.setName(name)
                imgf.setWidth(float(max(1.0, w)))
                imgf.setHeight(float(max(1.0, h)))
                c.insertImage(imgf)
                return
        except Exception:
            pass

        # Fallback: setCharFormat on a 1-char selection at the requested position
        try:
            c = QTextCursor(doc)
            c.setPosition(target_pos)
            c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            imgf = QTextImageFormat()
            if name:
                imgf.setName(name)
            imgf.setWidth(float(max(1.0, w)))
            imgf.setHeight(float(max(1.0, h)))
            c.setCharFormat(imgf)
        except Exception:
            pass
    except Exception:
        pass


def _fit_image_to_width_at(text_edit: QtWidgets.QTextEdit, cursor_pos: int, name: str, iw: int, ih: int):
    try:
        if not (iw and ih):
            return
        vp = text_edit.viewport()
        avail = max(16.0, float(vp.width() - 24)) if vp is not None else float(iw)
        scale = avail / float(iw)
        _apply_image_size_at(text_edit, cursor_pos, name, avail, max(1.0, ih * scale))
    except Exception:
        pass


def _reset_image_size_at(text_edit: QtWidgets.QTextEdit, cursor_pos: int, name: str, iw: int, ih: int):
    try:
        if iw and ih:
            _apply_image_size_at(text_edit, cursor_pos, name, float(iw), float(ih))
    except Exception:
        pass


def _prompt_resize_image(text_edit: QtWidgets.QTextEdit, cursor_pos: int, name: str, cur_w: float, cur_h: float, iw: int, ih: int):
    try:
        aspect = (float(iw) / float(ih)) if (iw and ih and ih != 0) else (float(cur_w) / float(cur_h) if cur_h else 1.0)
        w_default = int(cur_w or (iw or 400))
        w_txt, ok = QtWidgets.QInputDialog.getText(text_edit, "Resize Image", "Width (px):", text=str(w_default))
        if not (ok and w_txt and w_txt.strip()):
            return
        try:
            new_w = float(int(w_txt.strip()))
        except Exception:
            return
        new_w = max(16.0, new_w)
        new_h = max(1.0, new_w / (aspect if aspect else 1.0))
        _apply_image_size_at(text_edit, cursor_pos, name, new_w, new_h)
    except Exception:
        pass


def _strip_match_style_html(html: str) -> str:
    """Remove background, font-size, and font-family related styles/attributes so current style applies immediately."""
    import re

    s = html
    # Remove any <style>...</style> blocks entirely to prevent global overrides affecting pasted fragment
    s = re.sub(r"<\s*style\b[^>]*>.*?<\s*/\s*style\s*>", "", s, flags=re.IGNORECASE | re.DOTALL)
    # Remove bgcolor attribute
    s = re.sub(r'\sbgcolor\s*=\s*"[^"]*"', "", s, flags=re.IGNORECASE)
    s = re.sub(r"\sbgcolor\s*=\s*'[^']*'", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\sbgcolor\s*=\s*[^\s>]+", "", s, flags=re.IGNORECASE)
    # Replace deprecated <font> tags with span
    s = re.sub(r"<\s*font\b[^>]*>", "<span>", s, flags=re.IGNORECASE)
    s = re.sub(r"<\s*/\s*font\s*>", "</span>", s, flags=re.IGNORECASE)
    # Drop face/size/color attributes
    s = re.sub(r'\s(face|size|color)\s*=\s*"[^"]*"', "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s(face|size|color)\s*=\s*'[^']*'", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s(face|size|color)\s*=\s*[^\s>]+", "", s, flags=re.IGNORECASE)

    # Clean style attributes: remove background*, font-size, font-family, shorthand font, and line-height
    def _clean_style(m):
        inner = m.group(1)
        parts = [p.strip() for p in inner.split(";") if p.strip()]
        kept = []
        for p in parts:
            key = p.split(":", 1)[0].strip().lower()
            if key.startswith("background") or key in (
                "font-size",
                "font-family",
                "font",
                "line-height",
            ):
                continue
            kept.append(p)
        if not kept:
            return ""
        return ' style="' + "; ".join(kept) + '"'

    s = re.sub(r'\sstyle\s*=\s*"([^"]*)"', _clean_style, s, flags=re.IGNORECASE)
    s = re.sub(
        r"\sstyle\s*=\s*'([^']*)'",
        lambda m: _clean_style(m).replace('"', '"'),
        s,
        flags=re.IGNORECASE,
    )
    return s


def sanitize_html_for_storage(raw_html: str) -> str:
    """Strip inline font/background styles and classes before saving to avoid size/background regressions on reload.
    Keeps structure (p, br, lists, basic formatting), links and images.
    """
    if not isinstance(raw_html, str) or not raw_html:
        return raw_html
    from html.parser import HTMLParser

    class _StoreCleaner(HTMLParser):
        def __init__(self):
            super().__init__()
            self.out = []
            self._skip_style = False

        def handle_starttag(self, tag, attrs):
            tag_l = tag.lower()
            if tag_l == "style":
                # Skip entire style blocks
                self._skip_style = True
                return
            # Convert deprecated <font> to span
            if tag_l == "font":
                tag_l = "span"
            allowed = []
            buffered_style = None
            for k, v in attrs:
                lk = k.lower()
                if lk in ("class", "color", "face", "size"):
                    continue
                if lk == "bgcolor":
                    # Preserve bgcolor on table elements to retain shading
                    if tag_l in ("td", "th", "tr"):
                        allowed.append((k, v))
                    continue
                if lk == "style":
                    # Keep safe styles, including inline character formatting so user-applied
                    # bold/italic/underline/strike, colors, and custom font family/size persist.
                    #
                    # We preserve:
                    # - list-related (-qt-list-*, -qt-paragraph-type)
                    # - paragraph/div margin-left and text-align
                    # - table cell/row background-color and text-align
                    # - character styles: font-weight, font-style, text-decoration, color,
                    #   background/background-color, font-family, font-size
                    if tag_l in ("ol", "ul", "li", "p", "div", "td", "th", "tr", "span", "a", "em", "strong", "b", "i", "u", "s", "hr"):
                        try:
                            decls = [d.strip() for d in str(v).split(";") if d.strip()]
                            kept = []
                            for d in decls:
                                parts = d.split(":", 1)
                                key = parts[0].strip().lower() if len(parts) == 2 else ""
                                if key.startswith("-qt-list-") or key == "-qt-paragraph-type":
                                    kept.append(d)
                                elif tag_l in ("p", "div") and key in ("margin-left", "text-align"):
                                    kept.append(d)
                                elif tag_l in ("td", "th", "tr", "hr") and key in (
                                    "background",
                                    "background-color",
                                    "text-align",
                                    "border",
                                    "border-top",
                                    "border-right",
                                    "border-bottom",
                                    "border-left",
                                ):
                                    kept.append(d)
                                # Allow inline char formatting on common tags
                                elif key in (
                                    "font-weight",
                                    "font-style",
                                    "text-decoration",
                                    "color",
                                    "background",
                                    "background-color",
                                    "font-family",
                                    "font-size",
                                ):
                                    kept.append(d)
                            if kept:
                                buffered_style = "; ".join(kept)
                        except Exception:
                            pass
                    continue
                # Preserve explicit HTML alignment on paragraphs/divs
                if tag_l in ("p", "div") and lk == "align":
                    allowed.append((k, v))
                # Preserve list semantics
                if tag_l in ("ol", "ul") and lk in (
                    "type",
                    "start",
                ):  # type may be set by Qt for list appearance
                    allowed.append((k, v))
                elif tag_l == "li" and lk in ("value",):  # value allows continuing numbering
                    allowed.append((k, v))
                elif tag_l == "a" and lk in ("href", "title"):
                    allowed.append((k, v))
                elif tag_l == "img" and lk in ("src", "alt", "title", "width", "height"):
                    allowed.append((k, v))
                elif lk.startswith("data-"):
                    # Previously preserved formula-related attributes; now dropped after rollback.
                    continue
                if tag_l == "div" and lk == "id" and str(v) == "NB_DATA_FORMULAS":
                    # Sidecar div no longer used; skip.
                    continue
                elif lk in (
                    "width",
                    "height",
                    "cellpadding",
                    "cellspacing",
                    "border",
                ) and tag_l in ("table", "td", "th", "tr"):
                    allowed.append((k, v))
                elif tag_l in ("td", "th") and lk in ("colspan", "rowspan", "align", "valign"):
                    allowed.append((k, v))
                # drop everything else
            if buffered_style:
                allowed.append(("style", buffered_style))
            attrs_txt = "".join(f' {k}="{v}"' for k, v in allowed)
            self.out.append(f"<{tag_l}{attrs_txt}>")

        def handle_endtag(self, tag):
            tag_l = tag.lower()
            if tag_l == "style":
                self._skip_style = False
                return
            if tag_l == "font":
                tag_l = "span"
            self.out.append(f"</{tag_l}>")

        def handle_startendtag(self, tag, attrs):
            tag_l = tag.lower()
            if tag_l == "style":
                return
            allowed = []
            buffered_style = None
            for k, v in attrs:
                lk = k.lower()
                if lk in ("class", "color", "face", "size"):
                    continue
                if lk == "bgcolor":
                    if tag_l in ("td", "th", "tr"):
                        allowed.append((k, v))
                    continue
                if lk == "style":
                    if tag_l in ("ol", "ul", "li", "p", "div", "td", "th", "tr", "span", "a", "em", "strong", "b", "i", "u", "s", "hr"):
                        try:
                            decls = [d.strip() for d in str(v).split(";") if d.strip()]
                            kept = []
                            for d in decls:
                                parts = d.split(":", 1)
                                key = parts[0].strip().lower() if len(parts) == 2 else ""
                                if key.startswith("-qt-list-") or key == "-qt-paragraph-type":
                                    kept.append(d)
                                elif tag_l in ("p", "div") and key in ("margin-left", "text-align"):
                                    kept.append(d)
                                elif tag_l in ("td", "th", "tr", "hr") and key in (
                                    "background",
                                    "background-color",
                                    "text-align",
                                    "border",
                                    "border-top",
                                    "border-right",
                                    "border-bottom",
                                    "border-left",
                                ):
                                    kept.append(d)
                                elif key in (
                                    "font-weight",
                                    "font-style",
                                    "text-decoration",
                                    "color",
                                    "background",
                                    "background-color",
                                    "font-family",
                                    "font-size",
                                ):
                                    kept.append(d)
                            if kept:
                                buffered_style = "; ".join(kept)
                        except Exception:
                            pass
                    continue
                if tag_l in ("p", "div") and lk == "align":
                    allowed.append((k, v))
                if tag_l == "div" and lk == "id" and str(v) == "NB_DATA_FORMULAS":
                    # Skip legacy sidecar id.
                    continue
                if lk.startswith("data-"):
                    # Drop legacy formula attributes.
                    continue
                if tag_l == "img" and lk in ("src", "alt", "title", "width", "height"):
                    allowed.append((k, v))
                elif tag_l in ("td", "th") and lk in ("colspan", "rowspan", "align", "valign"):
                    allowed.append((k, v))
            if buffered_style:
                allowed.append(("style", buffered_style))
            attrs_txt = "".join(f' {k}="{v}"' for k, v in allowed)
            self.out.append(f"<{tag_l}{attrs_txt}/>")

        def handle_data(self, data):
            if not self._skip_style:
                self.out.append(data)

    try:
        cl = _StoreCleaner()
        cl.feed(raw_html)
        return "".join(cl.out)
    except Exception:
        return raw_html


def _merge_with_adjacent_lists(cursor: QTextCursor, cur_list: QTextList):
    """If the current list is adjacent to another with the same style/level, merge them so numbering continues.
    This helps when indenting/outdenting so the sequence doesn't restart.
    """
    try:
        block = cursor.block()
        prev = block.previous()
        nextb = block.next()
        cur_fmt = cur_list.format()
        # Search backward for the nearest compatible previous list (same style & indent)
        search = prev
        steps = 0
        current_level = cur_fmt.indent()
        while search.isValid() and steps < 500:
            tl = search.textList()
            if tl is None:
                # blank/unstyled paragraph: treat as a boundary
                break
            pf = tl.format()
            if pf.indent() < current_level:
                # crossed into a parent or higher-level boundary; don't merge across parents
                break
            if pf.indent() == current_level and pf.style() == cur_fmt.style():
                tl.add(block)
                return
            search = search.previous()
            steps += 1
        # Merge with next list if same style/indent
        if nextb.isValid() and nextb.textList() is not None:
            nl = nextb.textList()
            nf = nl.format()
            if nf.style() == cur_fmt.style() and nf.indent() == cur_fmt.indent():
                # Move next blocks into current list
                while nextb.isValid() and nextb.textList() is nl:
                    cur_list.add(nextb)
                    nextb = nextb.next()
    except Exception:
        pass


def _is_ordered_style(style: QTextListFormat.Style) -> bool:
    return style in (
        QTextListFormat.ListDecimal,
        QTextListFormat.ListLowerAlpha,
        QTextListFormat.ListUpperAlpha,
        QTextListFormat.ListLowerRoman,
        QTextListFormat.ListUpperRoman,
    )


def _ordered_style_for_level(level: int) -> QTextListFormat.Style:
    lvl = max(1, level)
    if _ORDERED_SCHEME == "decimal":
        return QTextListFormat.ListDecimal
    # 'classic' default: I, A, 1, a, i (repeat)
    idx = (lvl - 1) % 5
    if idx == 0:
        return QTextListFormat.ListUpperRoman
    if idx == 1:
        return QTextListFormat.ListUpperAlpha
    if idx == 2:
        return QTextListFormat.ListDecimal
    if idx == 3:
        return QTextListFormat.ListLowerAlpha
    return QTextListFormat.ListLowerRoman


def _unordered_style_for_level(level: int) -> QTextListFormat.Style:
    lvl = max(1, level)
    if _UNORDERED_SCHEME == "disc-only":
        return QTextListFormat.ListDisc
    # 'disc-circle-square' default (cycles every 3 levels)
    idx = (lvl - 1) % 3
    if idx == 0:
        return QTextListFormat.ListDisc
    if idx == 1:
        return QTextListFormat.ListCircle
    return QTextListFormat.ListSquare


def set_list_schemes(ordered: str = None, unordered: str = None):
    """Set list numbering/bullet schemes globally for all editors.
    ordered: 'classic' or 'decimal'
    unordered: 'disc-circle-square' or 'disc-only'
    """
    global _ORDERED_SCHEME, _UNORDERED_SCHEME
    if ordered in (None, ""):
        pass
    elif ordered in ("classic", "decimal"):
        _ORDERED_SCHEME = ordered
    if unordered in (None, ""):
        pass
    elif unordered in ("disc-circle-square", "disc-only"):
        _UNORDERED_SCHEME = unordered


def get_list_schemes():
    return _ORDERED_SCHEME, _UNORDERED_SCHEME


def _insert_image_via_dialog(text_edit: QtWidgets.QTextEdit, mode: str = "default", custom_width: float = None):
    """Open a file dialog and insert the chosen image using the specified sizing mode.

    mode in {"default", "fit-width", "original", "custom"}
    custom_width: used when mode == "custom" (pixels)
    """
    path, _ = QtWidgets.QFileDialog.getOpenFileName(
        text_edit, "Insert Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp);;All Files (*)"
    )
    if not path:
        return
    _insert_image_from_path(text_edit, path, mode=mode, custom_width=custom_width)


def _insert_image_from_path(text_edit: QtWidgets.QTextEdit, src_path: str, mode: str = "default", custom_width: float = None):
    """Insert an image scaled to DEFAULT_IMAGE_LONG_SIDE using QTextImageFormat.

    - Saves into the media store when a DB is open and uses a relative src resolved via document baseUrl.
    - Otherwise inserts from the absolute file path.
    - Computes intrinsic dimensions and sets width/height on the image format to enforce display size.
    """
    try:
        if not (isinstance(src_path, str) and src_path):
            return
        # Determine if it's an image; if not, just insert as-is
        if not os.path.exists(src_path) or imghdr.what(src_path) is None:
            try:
                text_edit.textCursor().insertImage(src_path)
            except Exception:
                pass
            return

        win = text_edit.window()
        db_path = getattr(win, "_db_path", None)
        media_root = getattr(win, "_media_root", None)

        rel_or_abs = src_path
        # If we have a DB/media root, store and use a relative path and ensure baseUrl
        if db_path and media_root:
            try:
                from media_store import save_file_into_store

                _, rel_path = save_file_into_store(db_path, src_path)
                rel_or_abs = rel_path
                base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                text_edit.document().setBaseUrl(QUrl.fromLocalFile(base))
            except Exception:
                rel_or_abs = src_path

        # Compute intrinsic image size
        try:
            # Prefer absolute path for probing dimensions
            abs_probe = src_path
            if db_path and media_root and rel_or_abs and not os.path.isabs(rel_or_abs):
                abs_probe = os.path.join(media_root, rel_or_abs)
            img = QImage(abs_probe)
            iw = int(img.width()) if not img.isNull() else 0
            ih = int(img.height()) if not img.isNull() else 0
        except Exception:
            iw = ih = 0

        # Determine display size per mode
        disp_w = disp_h = None
        if mode == "original" and iw > 0 and ih > 0:
            disp_w = float(iw)
            disp_h = float(ih)
        elif mode == "fit-width" and iw > 0 and ih > 0:
            try:
                vp = text_edit.viewport()
                avail = float(vp.width() - 24) if vp is not None else float(iw)
            except Exception:
                avail = float(iw)
            avail = max(16.0, avail)
            scale = avail / float(iw) if iw > 0 else 1.0
            disp_w = avail
            disp_h = max(1.0, float(ih) * scale)
        elif mode == "custom" and custom_width and iw > 0 and ih > 0:
            disp_w = max(1.0, float(custom_width))
            disp_h = max(1.0, disp_w * (float(ih) / float(iw)))
        else:
            # default mode: scale long side to DEFAULT_IMAGE_LONG_SIDE (preserve aspect)
            target_long = float(DEFAULT_IMAGE_LONG_SIDE)
            if iw > 0 and ih > 0:
                scale = target_long / float(max(iw, ih)) if max(iw, ih) > 0 else 1.0
                disp_w = max(1.0, float(iw) * scale)
                disp_h = max(1.0, float(ih) * scale)
            else:
                # Unknown size: default to target width/height
                disp_w = target_long
                disp_h = target_long

        # Insert using QTextImageFormat with width/height so size is enforced
        try:
            imgf = QTextImageFormat()
            imgf.setName(rel_or_abs)
            imgf.setWidth(disp_w)
            imgf.setHeight(disp_h)
            text_edit.textCursor().insertImage(imgf)
        except Exception:
            # Fallback to HTML if insertImage fails
            w_attr = f' width="{int(disp_w)}"' if disp_w else ""
            h_attr = f' height="{int(disp_h)}"' if disp_h else ""
            html = f'<img src="{rel_or_abs}"{w_attr}{h_attr} />'
            try:
                text_edit.textCursor().insertHtml(html)
            except Exception:
                pass
    except Exception:
        return




def _insert_video_via_dialog(text_edit: QtWidgets.QTextEdit, capture_seconds: float = None, force_synthetic: bool = False, size_mode: str = "default", custom_width: float = None):
    path, _ = QtWidgets.QFileDialog.getOpenFileName(
        text_edit,
        "Insert Video",
        "",
        "Videos (*.mp4 *.webm *.mov *.avi *.mkv);;All Files (*)",
    )
    if not path:
        return
    _insert_video_from_path(
        text_edit,
        path,
        capture_seconds=capture_seconds,
        force_synthetic=force_synthetic,
        size_mode=size_mode,
        custom_width=custom_width,
    )


def _insert_video_from_path(
    text_edit: QtWidgets.QTextEdit,
    src_path: str,
    capture_seconds: float = None,
    force_synthetic: bool = False,
    size_mode: str = "default",
    custom_width: float = None,
):
    try:
        if not os.path.exists(src_path):
            return
        win = text_edit.window()
        db_path = getattr(win, "_db_path", None)
        media_root = getattr(win, "_media_root", None)
        if db_path and media_root:
            try:
                from media_store import save_file_into_store

                # Save the video into the media store
                _, rel_video = save_file_into_store(db_path, src_path)
                # Ensure baseUrl so relative src/href resolve
                try:
                    base = media_root if media_root.endswith(os.sep) else media_root + os.sep
                    text_edit.document().setBaseUrl(QUrl.fromLocalFile(base))
                except Exception:
                    pass

                # Attempt a real frame thumbnail via OpenCV at configurable seconds; fallback to synthetic
                def _video_thumb_seconds() -> float:
                    try:
                        v = os.environ.get("NOTEBOOK_VIDEO_THUMB_SECONDS", "1.0").strip()
                        return float(v)
                    except Exception:
                        return 1.0

                def _opencv_available() -> bool:
                    try:
                        import cv2  # type: ignore
                        _ = cv2.__version__
                        return True
                    except Exception:
                        return False

                def _extract_frame_with_opencv(source: str, out_png: str, t_sec: float) -> bool:
                    try:
                        import cv2  # type: ignore
                        cap = cv2.VideoCapture(source)
                        if not cap.isOpened():
                            return False
                        # Seek by time when supported
                        ok_seek = cap.set(cv2.CAP_PROP_POS_MSEC, max(0.0, float(t_sec)) * 1000.0)
                        if not ok_seek:
                            # Fallback: estimate frame index by FPS
                            fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
                            idx = int(max(0.0, float(t_sec)) * float(fps))
                            cap.set(cv2.CAP_PROP_POS_FRAMES, idx)
                        ret, frame = cap.read()
                        cap.release()
                        if not ret or frame is None:
                            return False
                        # Write PNG
                        ok = cv2.imwrite(out_png, frame)
                        return bool(ok and os.path.exists(out_png))
                    except Exception:
                        return False

                tmp_dir = tempfile.gettempdir()
                tmp_thumb = os.path.join(tmp_dir, f"nb_thumb_{os.getpid()}_{abs(hash(src_path))}.png")
                made_real = False
                # Prefer OpenCV if available
                use_sec = float(capture_seconds) if capture_seconds is not None else _video_thumb_seconds()
                if (not force_synthetic) and _opencv_available():
                    if _extract_frame_with_opencv(src_path, tmp_thumb, use_sec):
                        made_real = True
                if not made_real:
                    # Synthetic 16:9 thumbnail with play icon and filename
                    thumb_w, thumb_h = 1280, 720
                    pm = QPixmap(QSize(thumb_w, thumb_h))
                    pm.fill(QColor(20, 20, 20))
                    p = QPainter(pm)
                    try:
                        p.setRenderHint(QPainter.Antialiasing, True)
                    except Exception:
                        pass
                    try:
                        pen = QPen(QColor(0, 0, 0, 140))
                        pen.setWidth(6)
                        p.setPen(pen)
                        p.drawRect(3, 3, thumb_w - 6, thumb_h - 6)
                    except Exception:
                        pass
                    try:
                        play_size = int(min(thumb_w, thumb_h) * 0.22)
                        cx, cy = thumb_w // 2, thumb_h // 2
                        pts = [
                            QPoint(cx - play_size // 3, cy - play_size // 2),
                            QPoint(cx - play_size // 3, cy + play_size // 2),
                            QPoint(cx + play_size // 2, cy),
                        ]
                        p.setBrush(QBrush(QColor(255, 255, 255, 230)))
                        p.setPen(Qt.NoPen)
                        p.drawPolygon(*pts)
                    except Exception:
                        pass
                    try:
                        name = os.path.basename(src_path)
                        overlay_h = int(thumb_h * 0.16)
                        p.setPen(Qt.NoPen)
                        p.setBrush(QBrush(QColor(0, 0, 0, 140)))
                        p.drawRect(0, thumb_h - overlay_h, thumb_w, overlay_h)
                        p.setPen(QColor(240, 240, 240))
                        f = QFont()
                        f.setPointSizeF(max(12.0, overlay_h * 0.35))
                        f.setBold(False)
                        p.setFont(f)
                        p.drawText(QRect(12, thumb_h - overlay_h, thumb_w - 24, overlay_h), Qt.AlignVCenter | Qt.TextSingleLine, name)
                    except Exception:
                        pass
                    try:
                        p.end()
                    except Exception:
                        pass
                    try:
                        pm.save(tmp_thumb, "PNG")
                    except Exception:
                        tmp_thumb = None

                rel_thumb = None
                if tmp_thumb and os.path.exists(tmp_thumb):
                    try:
                        _, rel_thumb = save_file_into_store(db_path, tmp_thumb, original_filename=os.path.basename(src_path) + ".thumb.png")
                    except Exception:
                        rel_thumb = None
                    try:
                        os.remove(tmp_thumb)
                    except Exception:
                        pass

                if rel_thumb is not None:
                    try:
                        # Probe intrinsic size from stored file
                        abs_probe = os.path.join(media_root, rel_thumb) if not os.path.isabs(rel_thumb) else rel_thumb
                        img_probe = QImage(abs_probe)
                        iw = int(img_probe.width()) if not img_probe.isNull() else 1280
                        ih = int(img_probe.height()) if not img_probe.isNull() else 720
                        # Determine display size based on sizing mode (video sizing can use its own default)
                        try:
                            if size_mode == "original":
                                disp_w, disp_h = iw, ih
                            elif size_mode == "fit-width":
                                try:
                                    vp = text_edit.viewport()
                                    avail = max(16, vp.width() - 32)
                                except Exception:
                                    avail = int(DEFAULT_IMAGE_LONG_SIDE)
                                # Scale preserving aspect ratio to fit width
                                scale = float(avail) / float(iw) if iw > 0 else 1.0
                                disp_w = int(avail)
                                disp_h = max(1, int(ih * scale))
                            elif size_mode == "custom" and custom_width and custom_width > 0:
                                scale = float(custom_width) / float(iw) if iw > 0 else 1.0
                                disp_w = int(custom_width)
                                disp_h = max(1, int(ih * scale))
                            else:  # default long-side scaling
                                # Prefer separate video default if defined
                                target_long_default = float(DEFAULT_VIDEO_LONG_SIDE) if 'DEFAULT_VIDEO_LONG_SIDE' in globals() and DEFAULT_VIDEO_LONG_SIDE else float(DEFAULT_IMAGE_LONG_SIDE)
                                target_long = float(custom_width) if (size_mode == "custom" and custom_width and custom_width > 0) else target_long_default
                                scale = target_long / float(max(iw, ih)) if max(iw, ih) > 0 else 1.0
                                disp_w = max(1, int(iw * scale))
                                disp_h = max(1, int(ih * scale))
                        except Exception:
                            target_long = float(DEFAULT_VIDEO_LONG_SIDE) if 'DEFAULT_VIDEO_LONG_SIDE' in globals() else float(DEFAULT_IMAGE_LONG_SIDE)
                            scale = target_long / float(max(iw, ih)) if max(iw, ih) > 0 else 1.0
                            disp_w = max(1, int(iw * scale))
                            disp_h = max(1, int(ih * scale))
                    except Exception:
                        base_long = float(DEFAULT_VIDEO_LONG_SIDE) if 'DEFAULT_VIDEO_LONG_SIDE' in globals() else float(DEFAULT_IMAGE_LONG_SIDE)
                        disp_w = int(base_long)
                        disp_h = int(base_long * 9 / 16)
                    alt = os.path.basename(src_path)
                    html = (
                        f'<a href="{rel_video}">'
                        f'<img src="{rel_thumb}" width="{disp_w}" height="{disp_h}" alt="{_html_escape(alt)}" /></a>'
                    )
                    text_edit.textCursor().insertHtml(html)
                else:
                    # Fallback: simple link
                    name = os.path.basename(src_path)
                    html = f'<a href="{rel_video}">📹 {name}</a>'
                    text_edit.textCursor().insertHtml(html)
                return
            except Exception:
                pass
        # Fallback (no DB/media root): insert absolute-path link
        name = os.path.basename(src_path)
        html = f'<a href="{src_path}">📹 {name}</a>'
        text_edit.textCursor().insertHtml(html)
    except Exception:
        return




class _ImageResizeHandler(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit
        self._resizing = False
        self._start_pos = None
        self._cursor_pos = None
        self._img_src = None
        self._orig_w = None
        self._orig_h = None
        self._overlay = None  # type: QtWidgets.QWidget
        # Clean up safely if the editor is destroyed while handlers/overlay are active
        try:
            self._edit.destroyed.connect(self._on_edit_destroyed)
        except Exception:
            pass

    def set_overlay(self, overlay_widget: QtWidgets.QWidget):
        self._overlay = overlay_widget

    def _on_edit_destroyed(self, *args, **kwargs):
        # Avoid operating on deleted widgets
        try:
            self._overlay = None
            self._edit = None
        except Exception:
            pass

    def _image_at(self, pos):
        try:
            # Helper to detect if a cursor sits on an image char and return its image format + adjusted cursor
            def _img_info_for_cursor(cur: QTextCursor):
                try:
                    fmt = cur.charFormat()
                    if fmt is None:
                        return None
                    if hasattr(fmt, "isImageFormat") and fmt.isImageFormat():
                        return QTextImageFormat(fmt), QTextCursor(cur)
                    if fmt.objectType() == QTextFormat.ImageObject:
                        return QTextImageFormat(fmt), QTextCursor(cur)
                    # Try previous char (common when caret is just after the image)
                    prev = QTextCursor(cur)
                    prev.movePosition(QTextCursor.Left)
                    pf = prev.charFormat()
                    if pf is not None and (
                        (hasattr(pf, "isImageFormat") and pf.isImageFormat())
                        or pf.objectType() == QTextFormat.ImageObject
                    ):
                        return QTextImageFormat(pf), QTextCursor(prev)
                except Exception:
                    return None
                return None

            base_cursor = self._edit.cursorForPosition(pos)
            # Build candidates safely: base, base-1, base+1
            candidates = []
            try:
                c0 = QTextCursor(base_cursor)
                candidates.append(c0)
                c1 = QTextCursor(base_cursor)
                c1.movePosition(QTextCursor.Left)
                candidates.append(c1)
                c2 = QTextCursor(base_cursor)
                c2.movePosition(QTextCursor.Right)
                candidates.append(c2)
            except Exception:
                candidates = [QTextCursor(base_cursor)]

            for candidate in candidates:
                info = _img_info_for_cursor(candidate)
                if info is None:
                    continue
                imgf, c = info
                try:
                    name = imgf.name() if hasattr(imgf, "name") else imgf.property(QTextFormat.ImageName)
                except Exception:
                    name = ""

                # Prefer on-screen rect via adjacent cursor rectangles for accuracy
                # cursorRect can crash on invalid positions; guard it
                try:
                    r1 = self._edit.cursorRect(c)
                except Exception:
                    continue
                c_after = QTextCursor(c)
                c_after.movePosition(QTextCursor.Right)
                try:
                    r2 = self._edit.cursorRect(c_after)
                except Exception:
                    continue
                x_left = min(r1.left(), r2.left())
                x_right = max(r1.left(), r2.left())
                width_from_layout = max(0, x_right - x_left)
                # Estimate height from line rects
                y_top = min(r1.top(), r2.top())
                h_from_layout = max(r1.height(), r2.height())

                # Fallback to image format/intrinsic dims if layout width is zero
                w = imgf.width() if hasattr(imgf, "width") and imgf.width() else 0
                h = imgf.height() if hasattr(imgf, "height") and imgf.height() else 0
                if (not w or not h):
                    # Try intrinsic by loading image from media root
                    abs_path = name
                    try:
                        base = getattr(self._edit.window(), "_media_root", None)
                        if base and not os.path.isabs(name):
                            abs_path = os.path.join(base, name)
                        img = QImage(abs_path)
                        if not img.isNull():
                            w = img.width()
                            h = img.height()
                    except Exception:
                        pass
                try:
                    iw = int(max(1, float(width_from_layout or w)))
                    ih = int(max(1, float(h_from_layout if width_from_layout else h)))
                except Exception:
                    iw = int(width_from_layout or (w or 1))
                    ih = int(h_from_layout if width_from_layout else (h or 1))

                # If layout didn't give a width, compute left edge from the right caret position
                if width_from_layout == 0 and iw > 0:
                    x_left = int(r2.left() - iw)
                    y_top = int(r1.center().y() - (ih / 2))

                screen_rect = QRect(int(x_left), int(y_top), int(iw), int(ih))
                return {"cursor": c, "format": imgf, "name": name, "rect": screen_rect, "w": iw, "h": ih}
        except Exception:
            return None
        return None

    def eventFilter(self, obj, event):
        try:
            if not _is_alive(self._edit):
                return False
            vp = self._edit.viewport() if _is_alive(self._edit) else None
            if not _is_alive(vp):
                return False
            # Avoid acting while widgets are not yet visible/realized (early startup churn)
            try:
                if not (self._edit.isVisible() and vp.isVisible()):
                    return False
            except Exception:
                return False
            if obj not in (vp, self._edit):
                return super().eventFilter(obj, event)
            if event is None:
                return False
            et = event.type()
            # Only process expected GUI events; otherwise, ignore
            if et not in (
                QEvent.MouseButtonPress,
                QEvent.MouseMove,
                QEvent.MouseButtonRelease,
                QEvent.Leave,
                QEvent.Wheel,
            ):
                return False
            # Map position to viewport coordinates for consistent hit-testing (mouse events only)
            pos_vp = None
            if et in (QEvent.MouseButtonPress, QEvent.MouseMove, QEvent.MouseButtonRelease):
                try:
                    pos_vp = event.pos()
                    if obj is not vp and _is_alive(vp) and hasattr(obj, "mapTo"):
                        pos_vp = obj.mapTo(vp, pos_vp)
                except Exception:
                    pos_vp = None
            if et == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                if pos_vp is None:
                    return False
                info = self._image_at(pos_vp)
                if info is not None:
                    rect = info["rect"]
                    # Near right edge within 18px, allow a tiny vertical tolerance
                    edge_x = rect.x() + rect.width()
                    near_right = abs(pos_vp.x() - edge_x) <= 18 and QRect(rect.x(), rect.y()-4, rect.width(), rect.height()+8).contains(pos_vp, True)
                    if near_right:
                        self._resizing = True
                        self._start_pos = QPoint(pos_vp)
                        self._cursor_pos = info["cursor"].position()
                        self._img_src = info["name"]
                        self._orig_w = max(1, int(info["w"]))
                        self._orig_h = max(1, int(info["h"]))
                        try:
                            if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                                self._edit.viewport().setCursor(Qt.SizeHorCursor)
                        except Exception:
                            pass
                        try:
                            if _is_alive(self._overlay):
                                self._overlay.show_for_rect(rect)
                        except Exception:
                            pass
                        return True
            elif et == QEvent.MouseMove:
                if pos_vp is None:
                    return False
                if self._resizing and self._start_pos is not None and self._cursor_pos is not None:
                    dx = pos_vp.x() - self._start_pos.x()
                    new_w = max(16, self._orig_w + dx)
                    new_h = int(new_w * (self._orig_h / float(self._orig_w)))
                    # Apply new size to the image char format
                    doc = self._edit.document() if _is_alive(self._edit) else None
                    if doc is None or not _is_alive(doc):
                        return False
                    c = QTextCursor(doc)
                    c.setPosition(self._cursor_pos)
                    c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                    imgf = QTextImageFormat()
                    # Preserve src; for relative paths, name stays the rel path
                    if self._img_src:
                        imgf.setName(self._img_src)
                    imgf.setWidth(float(new_w))
                    imgf.setHeight(float(new_h))
                    c.setCharFormat(imgf)
                    try:
                        if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                            self._edit.viewport().setCursor(Qt.SizeHorCursor)
                    except Exception:
                        pass
                    # Update overlay to follow the image rect as it changes
                    try:
                        if _is_alive(self._overlay):
                            # Recompute rect from adjacent cursor rectangles for accuracy
                            cur = QTextCursor(doc)
                            cur.setPosition(self._cursor_pos)
                            r1 = self._edit.cursorRect(cur)
                            cur2 = QTextCursor(cur)
                            cur2.movePosition(QTextCursor.Right)
                            r2 = self._edit.cursorRect(cur2)
                            x_left = min(r1.left(), r2.left())
                            y_top = min(r1.top(), r2.top())
                            self._overlay.show_for_rect(QRect(int(x_left), int(y_top), int(new_w), int(new_h)))
                    except Exception:
                        pass
                    return True
                else:
                    # Hover feedback: show outline whenever over image; show resize cursor near right edge
                    info = self._image_at(pos_vp)
                    try:
                        if info is not None:
                            rect = info["rect"]
                            # Loosen hover region slightly to reduce flicker
                            hover_rect = rect.adjusted(-4, -4, 4, 4)
                            inside = hover_rect.contains(pos_vp, True)
                            near_right = abs(pos_vp.x() - (rect.x() + rect.width())) <= 18 and inside
                            if _is_alive(self._overlay) and inside:
                                self._overlay.show_for_rect(rect)
                            if near_right:
                                try:
                                    if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                                        self._edit.viewport().setCursor(Qt.SizeHorCursor)
                                except Exception:
                                    pass
                            else:
                                try:
                                    if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                                        self._edit.viewport().unsetCursor()
                                except Exception:
                                    pass
                            return True
                        else:
                            try:
                                if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                                    self._edit.viewport().unsetCursor()
                            except Exception:
                                pass
                            if _is_alive(self._overlay):
                                self._overlay.hide_handles()
                    except Exception:
                        pass
            elif et == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                if pos_vp is None:
                    # Even without pos, we can finalize a resize
                    pass
                if self._resizing:
                    self._resizing = False
                    self._start_pos = None
                    self._cursor_pos = None
                    self._img_src = None
                    self._orig_w = None
                    self._orig_h = None
                    try:
                        if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                            self._edit.viewport().unsetCursor()
                    except Exception:
                        pass
                    try:
                        if _is_alive(self._overlay):
                            self._overlay.hide_handles()
                    except Exception:
                        pass
                    return True
            elif et == QEvent.Leave:
                try:
                    if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                        self._edit.viewport().unsetCursor()
                except Exception:
                    pass
                try:
                    if _is_alive(self._overlay):
                        self._overlay.hide_handles()
                except Exception:
                    pass
            elif et in (QEvent.Wheel,):
                # On scroll, hide overlay and reset cursor to avoid stale visuals while content moves
                try:
                    if _is_alive(self._edit) and _is_alive(self._edit.viewport()):
                        self._edit.viewport().unsetCursor()
                except Exception:
                    pass
                try:
                    if _is_alive(self._overlay):
                        self._overlay.hide_handles()
                except Exception:
                    pass
        except RuntimeError:
            return False
        return super().eventFilter(obj, event)


class _ImageContextMenuHandler(QObject):
    """Adds image-specific options to the context menu: Resize…, Fit to width, Reset size."""
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit

    def _find_image_at(self, pos_vp):
        try:
            edit = self._edit
            if not _is_alive(edit):
                return None
            # Try robust shared detector first
            info = _image_info_at_position(edit, pos_vp)
            if info is not None:
                return info
            # Expand search: scan a small neighborhood around click point
            vp = edit.viewport()
            rect = vp.rect() if vp is not None else None
            for dy in (0, 4, -4, 8, -8, 12, -12):
                for dx in (0, 4, -4, 8, -8, 12, -12, 16, -16, 24, -24):
                    pt = QPoint(int(pos_vp.x() + dx), int(pos_vp.y() + dy))
                    if rect is not None and not rect.contains(pt):
                        continue
                    info = _image_info_at_position(edit, pt)
                    if info is not None:
                        return info
            # As a last resort, look at cursor at/near click position and neighbors
            cur = edit.cursorForPosition(pos_vp)
            for candidate in (QTextCursor(cur),):
                fmt = candidate.charFormat()
                if fmt is not None and ((hasattr(fmt, "isImageFormat") and fmt.isImageFormat()) or fmt.objectType() == QTextFormat.ImageObject):
                    imgf = QTextImageFormat(fmt)
                    return {"cursor_pos": candidate.position(), "name": imgf.name(), "w": imgf.width() or 0.0, "h": imgf.height() or 0.0}
                prev = QTextCursor(candidate)
                prev.movePosition(QTextCursor.Left)
                pf = prev.charFormat()
                if pf is not None and ((hasattr(pf, "isImageFormat") and pf.isImageFormat()) or pf.objectType() == QTextFormat.ImageObject):
                    imgf = QTextImageFormat(pf)
                    return {"cursor_pos": prev.position(), "name": imgf.name(), "w": imgf.width() or 0.0, "h": imgf.height() or 0.0}
        except Exception:
            return None
        return None

    def _intrinsic_size(self, name: str):
        try:
            path = name
            base = getattr(self._edit.window(), "_media_root", None)
            if base and name and not os.path.isabs(name):
                path = os.path.join(base, name)
            img = QImage(path)
            if not img.isNull():
                return img.width(), img.height()
        except Exception:
            pass
        return None, None

    def _apply_size(self, cursor_pos: int, name: str, w: float, h: float):
        try:
            doc = self._edit.document()
            c = QTextCursor(doc)
            c.setPosition(cursor_pos)
            c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            imgf = QTextImageFormat()
            if name:
                imgf.setName(name)
            imgf.setWidth(float(max(1.0, w)))
            imgf.setHeight(float(max(1.0, h)))
            c.setCharFormat(imgf)
        except Exception:
            pass

    def _fit_to_width(self, cursor_pos: int, name: str):
        iw, ih = self._intrinsic_size(name)
        if not iw or not ih:
            return
        try:
            vp = self._edit.viewport()
            avail = max(16, vp.width() - 24)
        except Exception:
            avail = max(16, iw)
        scale = avail / float(iw)
        self._apply_size(cursor_pos, name, avail, max(1.0, ih * scale))

    def _reset_size(self, cursor_pos: int, name: str):
        iw, ih = self._intrinsic_size(name)
        if iw and ih:
            self._apply_size(cursor_pos, name, float(iw), float(ih))

    def _prompt_resize(self, cursor_pos: int, name: str, cur_w: float, cur_h: float):
        # Ask for width in pixels; keep aspect ratio if intrinsic available
        iw, ih = self._intrinsic_size(name)
        aspect = (iw / ih) if (iw and ih and ih != 0) else (cur_w / cur_h if cur_h else 1.0)
        try:
            w_txt, ok = QtWidgets.QInputDialog.getText(self._edit, "Resize Image", "Width (px):", text=str(int(cur_w or (iw or 400))))
        except Exception:
            ok = False
            w_txt = None
        if not (ok and w_txt and w_txt.strip().isdigit()):
            return
        new_w = max(16.0, float(int(w_txt.strip())))
        new_h = max(1.0, float(new_w / aspect))
        self._apply_size(cursor_pos, name, new_w, new_h)

    def eventFilter(self, obj, event):
        try:
            if not _is_alive(self._edit):
                return False
            return False
        except Exception:
            return False

    def on_custom_menu(self, pos):
        try:
            if not _is_alive(self._edit):
                return
            vp = self._edit.viewport()
            # Map pos (can come from editor or viewport) into viewport coords
            if isinstance(pos, QPoint):
                pos_vp = pos
                if self.sender() is not vp and hasattr(self.sender(), "mapTo"):
                    try:
                        pos_vp = self.sender().mapTo(vp, pos)
                    except Exception:
                        pass
            else:
                # Fallback: center of cursor rect
                cur = self._edit.textCursor()
                r = self._edit.cursorRect(cur)
                pos_vp = r.center()
            # Move caret to click point to improve detection
            try:
                cur = self._edit.cursorForPosition(pos_vp)
                self._edit.setTextCursor(cur)
            except Exception:
                pass
            # Detect image
            info = _image_info_at_cursor(self._edit)
            if info is None:
                info = self._find_image_at(pos_vp)
            if info is None:
                try:
                    c_try = self._edit.cursorForPosition(pos_vp)
                    info = _image_info_near_doc_pos(self._edit, c_try.position())
                except Exception:
                    pass
            if info is None:
                try:
                    c_try = self._edit.cursorForPosition(pos_vp)
                    info = _image_info_in_block(self._edit, c_try.block(), prefer_pos=c_try.position())
                except Exception:
                    pass
            if info is None:
                # No image: show default menu
                try:
                    menu = self._edit.createStandardContextMenu()
                except Exception:
                    menu = QtWidgets.QMenu(self._edit)
                menu.exec_(self._edit.mapToGlobal(pos_vp))
                return
            # Show image-only menu
            menu = QtWidgets.QMenu(self._edit)
            cursor_pos = int(info.get("cursor_pos", 0))
            name = info.get("name", "")
            cur_w = float(info.get("w") or 0.0)
            cur_h = float(info.get("h") or 0.0)
            act_resize = menu.addAction("Image Properties…")
            menu.addSeparator()
            act_fit = menu.addAction("Fit to editor width")
            act_reset = menu.addAction("Reset to original size")

            # If this image is a video thumbnail (img wrapped in <a href=video>), add video actions
            rel_video = None
            try:
                # First try via char format anchor
                ctmp = QTextCursor(self._edit.document()); ctmp.setPosition(cursor_pos)
                csel = QTextCursor(ctmp); csel.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                fmt = csel.charFormat()
                href = None
                try:
                    if hasattr(fmt, "isAnchor") and fmt.isAnchor():
                        href = fmt.anchorHref() if hasattr(fmt, "anchorHref") else None
                except Exception:
                    href = None
                # Fallback: inspect block HTML around cursor
                if not href:
                    blk = ctmp.block(); bc = QTextCursor(blk); bc.select(QTextCursor.BlockUnderCursor)
                    frag_html = bc.selection().toHtml()
                    m = re.search(r"<a[^>]+href=\"([^\"]+)\"[^>]*>\s*<img[^>]*>\s*</a>", frag_html, re.IGNORECASE)
                    href = m.group(1) if m else None
                if isinstance(href, str) and href.strip():
                    hv = href.strip()
                    low = hv.lower()
                    if low.endswith((".mp4", ".webm", ".mov", ".avi", ".mkv")):
                        rel_video = hv
            except Exception:
                rel_video = None

            if rel_video:
                menu.addSeparator()
                act_open_video = menu.addAction("Open linked video")
                sub_regen = menu.addMenu("Regenerate video thumbnail")
                act_v1 = sub_regen.addAction("Frame at 1.0s")
                act_v3 = sub_regen.addAction("Frame at 3.0s")
                act_v5 = sub_regen.addAction("Frame at 5.0s")
                act_vc = sub_regen.addAction("Custom time…")
                sub_regen.addSeparator()
                act_vs = sub_regen.addAction("Use synthetic placeholder")

            chosen = menu.exec_(self._edit.mapToGlobal(pos_vp))
            if chosen is None:
                return
            if chosen == act_resize:
                iw, ih = self._intrinsic_size(name)
                info_d = {"cursor_pos": cursor_pos, "name": name, "w": cur_w, "h": cur_h, "iw": iw, "ih": ih}
                _image_properties_dialog_apply(self._edit, info_d)
            elif chosen == act_fit:
                self._fit_to_width(cursor_pos, name)
            elif chosen == act_reset:
                self._reset_size(cursor_pos, name)
            elif rel_video and 'act_open_video' in locals() and chosen == act_open_video:
                try:
                    base = getattr(self._edit.window(), "_media_root", None)
                    if base and not re.match(r"^[a-zA-Z]+:|^/", rel_video):
                        abs_path = os.path.normpath(os.path.join(base, rel_video))
                        QDesktopServices.openUrl(QUrl.fromLocalFile(abs_path))
                    else:
                        QDesktopServices.openUrl(QUrl(rel_video))
                except Exception:
                    pass
            elif rel_video and 'act_v1' in locals() and chosen in (act_v1, act_v3, act_v5, act_vc, act_vs):
                def _do_regen(secs=None, synthetic=False):
                    try:
                        base = getattr(self._edit.window(), "_media_root", None)
                        db_path = getattr(self._edit.window(), "_db_path", None)
                        if not (base and db_path):
                            return
                        v_abs = rel_video
                        if not os.path.isabs(v_abs):
                            v_abs = os.path.join(base, v_abs)
                        # Insert new thumbnail at current position
                        cpos = int(cursor_pos)
                        cset = QTextCursor(self._edit.document()); cset.setPosition(cpos)
                        self._edit.setTextCursor(cset)
                        _insert_video_from_path(self._edit, v_abs, capture_seconds=secs, force_synthetic=synthetic)
                        # Remove the old image char
                        cdel = QTextCursor(self._edit.document()); cdel.setPosition(cpos)
                        cdel.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                        try:
                            cdel.removeSelectedText()
                        except Exception:
                            pass
                    except Exception:
                        pass
                if chosen == act_vs:
                    _do_regen(secs=None, synthetic=True)
                elif chosen == act_vc:
                    try:
                        secs, ok = QtWidgets.QInputDialog.getDouble(self._edit, "Thumbnail time", "Seconds:", 1.0, 0.0, 36000.0, 1)
                        if ok:
                            _do_regen(secs=float(secs), synthetic=False)
                    except Exception:
                        pass
                else:
                    secs = 1.0 if chosen == act_v1 else (3.0 if chosen == act_v3 else 5.0)
                    _do_regen(secs=secs, synthetic=False)
        except Exception:
            try:
                menu = self._edit.createStandardContextMenu()
                menu.exec_(self._edit.mapToGlobal(pos if isinstance(pos, QPoint) else QPoint(0,0)))
            except Exception:
                pass


class _ImageHud(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit
        self._vp = edit.viewport()
        self._btn = QtWidgets.QToolButton(self._vp)
        self._btn.setText("Image…")
        try:
            self._btn.setIcon(_make_icon("image", QSize(16, 16)))
        except Exception:
            pass
        self._btn.setToolTip("Image Properties / Fit / Reset")
        self._btn.setVisible(False)
        self._btn.setAutoRaise(True)
        self._btn.clicked.connect(self._on_click)
        # Track moves/selection/scrolling
        try:
            edit.cursorPositionChanged.connect(self._update)
        except Exception:
            pass
        try:
            self._vp.installEventFilter(self)
        except Exception:
            pass
        self._last_info = None

    def eventFilter(self, obj, event):
        if obj is self._vp and event.type() in (QEvent.Resize, QEvent.Paint, QEvent.Wheel, QEvent.Scroll, QEvent.LayoutRequest):
            self._update()
        return super().eventFilter(obj, event)

    def _detect(self):
        edit = self._edit
        info = _image_info_at_cursor(edit)
        if info is None:
            cur = edit.textCursor()
            info = _image_info_near_doc_pos(edit, cur.position())
        if info is None:
            info = _image_info_in_block(edit, edit.textCursor().block(), prefer_pos=edit.textCursor().position())
        return info

    def _update(self):
        try:
            info = self._detect()
            self._last_info = info
            if info is None:
                self._btn.setVisible(False)
                return
            # Place button near the caret rect (top-right)
            r = self._edit.cursorRect(self._edit.textCursor())
            pt = QPoint(min(self._vp.width() - 40, r.right() + 6), max(0, r.top() - 2))
            self._btn.move(pt)
            self._btn.setVisible(True)
        except Exception:
            try:
                self._btn.setVisible(False)
            except Exception:
                pass

    def _on_click(self):
        info = self._last_info or self._detect()
        if info:
            _image_properties_dialog_apply(self._edit, info)


def _install_image_hud(text_edit: QtWidgets.QTextEdit):
    try:
        hud = _ImageHud(text_edit)
        if not hasattr(text_edit, "_imageHud"):
            text_edit._imageHud = []
        text_edit._imageHud.append(hud)
    except Exception:
        pass


def _install_image_context_menu(text_edit: QtWidgets.QTextEdit):
    try:
        if text_edit is None:
            return
        handler = _ImageContextMenuHandler(text_edit)
        # Connect custom context menu signals from both editor and viewport
        try:
            text_edit.customContextMenuRequested.connect(handler.on_custom_menu)
        except Exception:
            pass
        try:
            if text_edit.viewport() is not None:
                text_edit.viewport().setContextMenuPolicy(Qt.CustomContextMenu)
                text_edit.viewport().customContextMenuRequested.connect(handler.on_custom_menu)
        except Exception:
            pass
        if not hasattr(text_edit, "_imageContextHandlers"):
            text_edit._imageContextHandlers = []
        text_edit._imageContextHandlers.append(handler)
    except Exception:
        pass


def _install_image_resize_handler(text_edit: QtWidgets.QTextEdit):
    try:
        if text_edit is None:
            return
        handler = _ImageResizeHandler(text_edit)

        def _attach():
            try:
                if not _is_alive(text_edit):
                    return
                vp = text_edit.viewport()
                if vp is None or not _is_alive(vp):
                    return
                # Create overlay only after viewport exists and is visible
                overlay = _ImageResizeOverlay(text_edit)
                overlay.setParent(vp)
                overlay.setAttribute(Qt.WA_TransparentForMouseEvents, True)
                try:
                    overlay.setGeometry(vp.rect())
                except Exception:
                    pass
                overlay.hide()
                try:
                    overlay.raise_()
                except Exception:
                    pass
                handler.set_overlay(overlay)
                # Keep overlay sized with the viewport
                def _sync_overlay():
                    try:
                        overlay.setGeometry(vp.rect())
                        overlay.update()
                    except Exception:
                        pass
                try:
                    watcher = _ResizeViewportWatcher(overlay, _sync_overlay)
                    overlay._vpWatcher = watcher  # keep a python-side ref
                    vp.installEventFilter(watcher)
                except Exception:
                    pass
                # Ensure we get hover move events even without a button pressed
                try:
                    vp.setMouseTracking(True)
                except Exception:
                    pass
                vp.installEventFilter(handler)
                # Also enable tracking and filter on the editor
                try:
                    text_edit.setMouseTracking(True)
                    text_edit.installEventFilter(handler)
                except Exception:
                    pass
                # Drop overlay reference when viewport or editor is destroyed
                try:
                    vp.destroyed.connect(lambda *a: overlay.deleteLater())
                except Exception:
                    pass
            except Exception:
                pass

        # Delay attaching until after the current event cycle to avoid early lifecycle churn
        QTimer.singleShot(0, _attach)

        if not hasattr(text_edit, "_imageResizeHandlers"):
            text_edit._imageResizeHandlers = []
        text_edit._imageResizeHandlers.append(handler)
    except Exception:
        pass


class _ResizeViewportWatcher(QObject):
    """Internal helper to keep an overlay synced with the viewport size."""
    def __init__(self, overlay: QtWidgets.QWidget, on_change):
        super().__init__(overlay)
        self._overlay = overlay
        self._on_change = on_change

    def eventFilter(self, obj, event):
        et = event.type()
        if et in (QEvent.Resize, QEvent.Show, QEvent.Hide, QEvent.Wheel):
            try:
                self._on_change()
            except Exception:
                pass
        return False


class _ImageResizeOverlay(QtWidgets.QWidget):
    """A transparent overlay that draws visual resize handles around the hovered image."""
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit
        self._rect = None  # QRect in viewport coords
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)

    def show_for_rect(self, rect: QRect):
        try:
            self._rect = QRect(rect)
            if not self.isVisible():
                self.show()
            self.update()
        except Exception:
            pass

    def hide_handles(self):
        self._rect = None
        try:
            self.hide()
        except Exception:
            pass

    def paintEvent(self, ev):
        if self._rect is None:
            return
        try:
            p = QPainter(self)
            p.setRenderHint(QPainter.Antialiasing, True)
            # Outline (thicker, dashed for visibility) + light translucent fill
            accent = QColor(0, 120, 215)
            pen = QPen(accent)  # Windows accent blue
            pen.setWidth(2)
            pen.setStyle(Qt.DashLine)
            p.setPen(pen)
            # Semi-transparent fill to make the bounds obvious
            fill = QColor(accent)
            fill.setAlpha(40)
            p.fillRect(self._rect.adjusted(0, 0, -1, -1), fill)
            p.drawRect(self._rect.adjusted(0, 0, -1, -1))
            # Draw a single handle at right-middle (larger for hi-DPI)
            handle_size = 16
            rx = self._rect.right()
            cy = self._rect.center().y()
            handle_rect = QRect(rx - handle_size // 2, cy - handle_size // 2, handle_size, handle_size)
            # Handle with outline for contrast
            p.fillRect(handle_rect, accent)
            p.setPen(QPen(QColor(255, 255, 255), 1))
            p.drawRect(handle_rect.adjusted(0, 0, -1, -1))
            # Optional corner handles for future (not interactive yet)
            p.end()
        except Exception:
            pass


# ----------------------------- Tables -----------------------------
def _current_table(text_edit: QtWidgets.QTextEdit):
    try:
        cur = text_edit.textCursor()
        return cur.currentTable()
    except Exception:
        return None


def _enforce_uniform_table_borders(text_edit: QtWidgets.QTextEdit):
    """Ensure all tables in the editor use a single 1px solid black grid.

    Implementation:
    - Set table cellSpacing to 0 and table border to 0 (avoid doubling with cell borders)
    - Set every cell's border to 1px black, solid
    """
    if text_edit is None:
        return
    doc = text_edit.document()
    # Load theme (grid color and width)
    try:
        from settings_manager import get_table_theme
        theme = get_table_theme()
        grid_hex = theme.get("grid_color", "#000000")
        grid_w = float(theme.get("grid_width", 1.5))
    except Exception:
        grid_hex = "#000000"
        grid_w = 1.5
    cur = QTextCursor(doc)
    seen = set()
    while True:
        tbl = cur.currentTable()
        if tbl is not None:
            key = (tbl.firstPosition(), tbl.lastPosition())
            if key not in seen:
                seen.add(key)
                # Check if this table is an HR marker table; if so, skip border normalization
                skip = False
                try:
                    fmt = tbl.format()
                    # Always set zero spacing for consistency
                    try:
                        fmt.setCellSpacing(0.0)
                        tbl.setFormat(fmt)
                    except Exception:
                        pass
                    prop = fmt.property(int(QTextFormat.UserProperty) + 101)
                    skip = bool(prop)
                    # Additionally detect 1x1 top-border-only tables (HTML reload path)
                    if not skip:
                        try:
                            rows, cols = tbl.rows(), tbl.columns()
                        except Exception:
                            rows, cols = 0, 0
                        if rows == 1 and cols == 1:
                            try:
                                cell = tbl.cellAt(0, 0)
                                cf = QTextTableCellFormat(cell.format())
                                tb = float(getattr(cf, 'topBorder', lambda: 0.0)() or 0.0)
                                lb = float(getattr(cf, 'leftBorder', lambda: 0.0)() or 0.0)
                                rb = float(getattr(cf, 'rightBorder', lambda: 0.0)() or 0.0)
                                bb = float(getattr(cf, 'bottomBorder', lambda: 0.0)() or 0.0)
                                if tb > 0.0 and lb == 0.0 and rb == 0.0 and bb == 0.0:
                                    skip = True
                            except Exception:
                                pass
                except Exception:
                    skip = False
                if not skip:
                    # Normalize table format and apply per-cell borders
                    try:
                        fmt = tbl.format()
                        # Remove spacing between cells for tight single-line appearance
                        try:
                            fmt.setCellSpacing(0.0)
                        except Exception:
                            pass
                        # Use table border for the outer frame
                        fmt.setBorder(1.0)
                        try:
                            from PyQt5.QtGui import QBrush, QColor
                            fmt.setBorderBrush(QBrush(QColor(grid_hex)))
                        except Exception:
                            pass
                        try:
                            fmt.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)
                        except Exception:
                            pass
                        tbl.setFormat(fmt)
                    except Exception:
                        pass
                    # Apply per-cell borders
                    try:
                        rows, cols = tbl.rows(), tbl.columns()
                    except Exception:
                        rows, cols = 0, 0
                    for r in range(rows):
                        for c in range(cols):
                            try:
                                cell = tbl.cellAt(r, c)
                                cf = cell.format()
                                tcf = QTextTableCellFormat(cf)
                                # All sides single line at configured width/color
                                try:
                                    tcf.setBorder(float(grid_w))
                                except Exception:
                                    tcf.setBorder(1.5)
                                try:
                                    from PyQt5.QtGui import QBrush, QColor
                                    tcf.setBorderBrush(QBrush(QColor(grid_hex)))
                                except Exception:
                                    pass
                                try:
                                    tcf.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)
                                except Exception:
                                    pass
                                cell.setFormat(tcf)
                            except Exception:
                                pass
        if not cur.movePosition(QTextCursor.NextBlock):
            break


def insert_table_from_preset(text_edit: QtWidgets.QTextEdit, preset_name: str, fit_width_100: bool = True):
    """Insert a table defined by a saved preset at the current cursor position.

    Supports two schemas:
    - v2 (preferred): {"version": 2, "html": "<table>...</table>"}
      Creates an outer 1x2 container (100% width), inserts the saved HTML into the left cell,
      and inserts a blank 2-column Cost list into the right cell.
    - legacy: structural fields (rows/columns/width/etc.). Inserts a single table as before.

    fit_width_100: For legacy presets only, force table width to 100%% regardless of saved width.
    """
    if text_edit is None or not isinstance(preset_name, str) or not preset_name.strip():
        return
    try:
        from settings_manager import get_table_presets

        presets = get_table_presets()
        data = presets.get(preset_name)
    except Exception:
        data = None
    if not isinstance(data, dict):
        try:
            QtWidgets.QMessageBox.information(
                text_edit, "Insert Preset", f"Preset '{preset_name}' not found."
            )
        except Exception:
            pass
        return

    # v2 HTML-based preset path (now the only supported format)
    html = data.get("html") if isinstance(data, dict) else None
    if isinstance(html, str) and html.strip():
        cur = text_edit.textCursor()
        # If we're already inside an outer 1x2 container, reuse it to avoid nesting
        reuse_outer = False
        existing_table = cur.currentTable()
        if existing_table is not None:
            try:
                if existing_table.rows() == 1 and existing_table.columns() == 2:
                    reuse_outer = True
                    outer = existing_table
                else:
                    reuse_outer = False
            except Exception:
                reuse_outer = False

        if not reuse_outer:
            # Build outer 1x2 container at 100% width
            outer_fmt = QTextTableFormat()
            outer_fmt.setCellPadding(4)
            outer_fmt.setCellSpacing(0)
            outer_fmt.setBorder(1.0)
            try:
                outer_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                outer_fmt.setColumnWidthConstraints(
                    [
                        QTextLength(QTextLength.PercentageLength, 50.0),
                        QTextLength(QTextLength.PercentageLength, 50.0),
                    ]
                )
            except Exception:
                pass
            outer = cur.insertTable(1, 2, outer_fmt)
            # Re-apply format immediately to guard against layout quirks
            try:
                fmt_chk = outer.format()
                fmt_chk.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                fmt_chk.setColumnWidthConstraints(
                    [
                        QTextLength(QTextLength.PercentageLength, 50.0),
                        QTextLength(QTextLength.PercentageLength, 50.0),
                    ]
                )
                outer.setFormat(fmt_chk)
            except Exception:
                pass
        else:
            # Ensure the existing container is full width with 50/50 columns
            try:
                fmt = outer.format()
                fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                fmt.setColumnWidthConstraints(
                    [
                        QTextLength(QTextLength.PercentageLength, 50.0),
                        QTextLength(QTextLength.PercentageLength, 50.0),
                    ]
                )
                outer.setFormat(fmt)
            except Exception:
                pass

        # Insert saved HTML into the left cell
        try:
            left_cur = outer.cellAt(0, 0).firstCursorPosition()
            left_cur.insertHtml(html)
            # Ensure the inserted left table fills the cell and uses 50/25/25 columns
            try:
                from PyQt5.QtGui import QTextLength
                from ui_planning_register import _is_planning_register_table

                # Find first table inside the left cell range
                left_cell = outer.cellAt(0, 0)
                s_pos = left_cell.firstCursorPosition().position()
                e_pos = left_cell.lastCursorPosition().position()
                scan = QTextCursor(text_edit.document())
                scan.setPosition(s_pos)
                found_tbl = None
                iters = 0
                while scan.position() < e_pos and iters < 20000:
                    t = scan.currentTable()
                    if t is not None:
                        found_tbl = t
                        break
                    scan.movePosition(QTextCursor.NextCharacter)
                    iters += 1
                if found_tbl is not None and _is_planning_register_table(text_edit, found_tbl):
                    fmt = found_tbl.format()
                    fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                    fmt.setColumnWidthConstraints(
                        [
                            QTextLength(QTextLength.PercentageLength, 50.0),
                            QTextLength(QTextLength.PercentageLength, 25.0),
                            QTextLength(QTextLength.PercentageLength, 25.0),
                        ]
                    )
                    found_tbl.setFormat(fmt)
            except Exception:
                pass
        except Exception:
            pass

        # Insert blank right-side Costs table
        try:
            from ui_planning_register import (
                _insert_right_cost_table_in_cursor,
                ensure_planning_register_watcher,
                refresh_planning_register_styles,
                _is_planning_register_table,
                _recalc_planning_totals,
                _is_cost_list_table,
            )

            right_cell = outer.cellAt(0, 1)
            right_cur = right_cell.firstCursorPosition()
            # Scan the right cell for an inner table; only insert if a cost list table is not present
            s_pos = right_cell.firstCursorPosition().position()
            e_pos = right_cell.lastCursorPosition().position()
            scan = QTextCursor(text_edit.document())
            scan.setPosition(s_pos)
            found_inner = None
            iters = 0
            while scan.position() < e_pos and iters < 20000:
                t = scan.currentTable()
                if t is not None:
                    found_inner = t
                    break
                scan.movePosition(QTextCursor.NextCharacter)
                iters += 1
            if not (found_inner is not None and _is_cost_list_table(text_edit, found_inner)):
                _insert_right_cost_table_in_cursor(right_cur)
            # Ensure watcher active so cost formatting applies immediately
            ensure_planning_register_watcher(text_edit)
            # Reapply planning register visuals (header/totals shading, right alignment)
            try:
                refresh_planning_register_styles(text_edit)
            except Exception:
                pass
            # If left table is a planning register, ensure totals are correct now
            try:
                # Find the first table in the left cell and recalc if it matches
                left_tbl = left_cur.currentTable()
                if left_tbl is not None and _is_planning_register_table(text_edit, left_tbl):
                    _recalc_planning_totals(text_edit, left_tbl)
            except Exception:
                pass
        except Exception:
            pass

        # Place caret at end of right cell content
        try:
            after_right = outer.cellAt(0, 1).lastCursorPosition()
            text_edit.setTextCursor(after_right)
        except Exception:
            pass
        # Final enforcement: make sure the outer container is full-width with 50/50 split.
        try:
            fmt_final = outer.format()
            fmt_final.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
            fmt_final.setColumnWidthConstraints(
                [
                    QTextLength(QTextLength.PercentageLength, 50.0),
                    QTextLength(QTextLength.PercentageLength, 50.0),
                ]
            )
            outer.setFormat(fmt_final)
        except Exception:
            pass
        # Uniform borders across all tables present
        try:
            _enforce_uniform_table_borders(text_edit)
        except Exception:
            pass
        return outer
    # No HTML present -> preset unsupported
    try:
        QtWidgets.QMessageBox.information(
            text_edit,
            "Insert Preset",
            "This preset is from an older version and doesn't include table HTML. Please re-save it.",
        )
    except Exception:
        pass
    return None


def choose_and_insert_preset(text_edit: QtWidgets.QTextEdit, fit_width_100: bool = True):
    """Prompt for a preset name and insert it into the editor at the cursor."""
    if text_edit is None:
        return
    try:
        from settings_manager import list_table_preset_names

        names = list_table_preset_names()
    except Exception:
        names = []
    if not names:
        try:
            QtWidgets.QMessageBox.information(text_edit, "Insert Preset", "No presets saved yet.")
        except Exception:
            pass
        return
    try:
        name, ok = QtWidgets.QInputDialog.getItem(
            text_edit, "Insert Preset", "Preset:", names, 0, False
        )
    except Exception:
        ok = False
        name = None
    if not (ok and name and name.strip()):
        return
    insert_table_from_preset(text_edit, name.strip(), fit_width_100=fit_width_100)

def insert_planning_register_via_dialog(window: QtWidgets.QMainWindow):
    """Open a simple chooser to insert a Planning Register (new or from preset).

    This is invoked from the page context menu. It looks up the main page editor
    ('pageEdit') and then offers:
      - New Planning Register (programmatic layout)
      - Any saved presets (HTML-based)
    """
    try:
        te = window.findChild(QtWidgets.QTextEdit, "pageEdit")
    except Exception:
        te = None
    if te is None or not te.isEnabled():
        try:
            QtWidgets.QMessageBox.information(window, "Insert Planning Register", "Please open or create a page first.")
        except Exception:
            pass
        return

    # Build options: first 'New Planning Register', then saved presets (if any)
    try:
        from settings_manager import list_table_preset_names

        preset_names = list_table_preset_names()
    except Exception:
        preset_names = []
    options = ["New Planning Register"] + preset_names

    try:
        choice, ok = QtWidgets.QInputDialog.getItem(
            window, "Insert Planning Register", "Choose:", options, 0, False
        )
    except Exception:
        ok = False
        choice = None
    if not (ok and choice):
        return

    if choice == "New Planning Register":
        try:
            from ui_planning_register import insert_planning_register

            insert_planning_register(window)
        except Exception:
            pass
    else:
        # Insert from preset into the current editor; ensure full-width container
        try:
            insert_table_from_preset(te, choice, fit_width_100=True)
            try:
                _enforce_uniform_table_borders(te)
            except Exception:
                pass
        except Exception:
            pass

def save_current_table_as_preset(text_edit: QtWidgets.QTextEdit):
    """Save the table under the caret as a reusable preset (HTML-based)."""
    if text_edit is None:
        return
    tbl = _current_table(text_edit)
    if tbl is None:
        try:
            QtWidgets.QMessageBox.information(text_edit, "Save Table as Preset", "Place the caret inside a table to save it.")
        except Exception:
            pass
        return
    try:
        name, ok = QtWidgets.QInputDialog.getText(
            text_edit, "Save Table Preset", "Preset name:", text="My Table"
        )
    except Exception:
        ok = False
        name = None
    if not (ok and name and name.strip()):
        return

    def _extract_table_fragment(html_text: str) -> str:
        try:
            low = html_text.lower()
            start = low.find("<table")
            if start < 0:
                return html_text.strip()
            end = low.rfind("</table>")
            if end >= 0:
                end += len("</table>")
            else:
                end = len(html_text)
            return html_text[start:end].strip()
        except Exception:
            return html_text

    try:
        doc = text_edit.document()
        cur = QTextCursor(doc)
        cur.setPosition(tbl.firstPosition())
        cur.setPosition(tbl.lastPosition(), QTextCursor.KeepAnchor)
        raw_html = cur.selection().toHtml()
        table_html = _extract_table_fragment(raw_html)
    except Exception:
        table_html = ""
    if not table_html:
        try:
            QtWidgets.QMessageBox.information(text_edit, "Save Table as Preset", "Couldn't capture table HTML.")
        except Exception:
            pass
        return

    preset = {"version": 2, "html": table_html}
    try:
        from settings_manager import save_table_preset

        save_table_preset(name.strip(), preset)
        QtWidgets.QToolTip.showText(text_edit.mapToGlobal(text_edit.rect().center()), f"Saved preset '{name.strip()}'")
    except Exception:
        pass


def _table_insert_or_edit(text_edit: QtWidgets.QTextEdit):
    tbl = _current_table(text_edit)
    if tbl is None:
        _table_insert_dialog(text_edit)
    else:
        _table_properties_dialog(text_edit, tbl)


def _table_insert_dialog(text_edit: QtWidgets.QTextEdit):
    dlg = QtWidgets.QDialog(text_edit)
    dlg.setWindowTitle("Insert Table")
    form = QtWidgets.QFormLayout(dlg)
    sp_rows = QtWidgets.QSpinBox(dlg)
    sp_rows.setRange(1, 100)
    sp_rows.setValue(3)
    sp_cols = QtWidgets.QSpinBox(dlg)
    sp_cols.setRange(1, 20)
    sp_cols.setValue(3)
    sp_border = QtWidgets.QDoubleSpinBox(dlg)
    sp_border.setRange(0.0, 8.0)
    sp_border.setSingleStep(0.5)
    sp_border.setValue(1.0)
    sp_pad = QtWidgets.QDoubleSpinBox(dlg)
    sp_pad.setRange(0.0, 20.0)
    sp_pad.setSingleStep(0.5)
    sp_pad.setValue(2.0)
    sp_space = QtWidgets.QDoubleSpinBox(dlg)
    sp_space.setRange(0.0, 20.0)
    sp_space.setSingleStep(0.5)
    sp_space.setValue(0.0)
    sp_width = QtWidgets.QDoubleSpinBox(dlg)
    sp_width.setRange(10.0, 100.0)
    sp_width.setSingleStep(5.0)
    sp_width.setValue(100.0)
    form.addRow("Rows:", sp_rows)
    form.addRow("Columns:", sp_cols)
    form.addRow("Border width:", sp_border)
    form.addRow("Cell padding:", sp_pad)
    form.addRow("Cell spacing:", sp_space)
    form.addRow("Table width (% of editor):", sp_width)
    btns = QtWidgets.QDialogButtonBox(
        QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, parent=dlg
    )
    form.addRow(btns)
    btns.accepted.connect(dlg.accept)
    btns.rejected.connect(dlg.reject)
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return
    rows = sp_rows.value()
    cols = sp_cols.value()
    from PyQt5.QtGui import QTextLength, QTextTableFormat

    fmt = QTextTableFormat()
    fmt.setBorder(sp_border.value())
    fmt.setCellPadding(sp_pad.value())
    fmt.setCellSpacing(sp_space.value())
    # Set percentage width and distribute columns evenly
    try:
        fmt.setWidth(QTextLength(QTextLength.PercentageLength, sp_width.value()))
        if cols > 0:
            frac = 100.0 / float(cols)
            fmt.setColumnWidthConstraints(
                [QTextLength(QTextLength.PercentageLength, frac) for _ in range(cols)]
            )
    except Exception:
        pass
    cur = text_edit.textCursor()
    cur.insertTable(rows, cols, fmt)
    try:
        _enforce_uniform_table_borders(text_edit)
    except Exception:
        pass


def _table_properties_dialog(text_edit: QtWidgets.QTextEdit, table):
    fmt = table.format()
    dlg = QtWidgets.QDialog(text_edit)
    dlg.setWindowTitle("Table Properties")
    form = QtWidgets.QFormLayout(dlg)
    sp_border = QtWidgets.QDoubleSpinBox(dlg)
    sp_border.setRange(0.0, 8.0)
    sp_border.setSingleStep(0.5)
    sp_border.setValue(fmt.border())
    sp_pad = QtWidgets.QDoubleSpinBox(dlg)
    sp_pad.setRange(0.0, 20.0)
    sp_pad.setSingleStep(0.5)
    sp_pad.setValue(fmt.cellPadding())
    sp_space = QtWidgets.QDoubleSpinBox(dlg)
    sp_space.setRange(0.0, 20.0)
    sp_space.setSingleStep(0.5)
    sp_space.setValue(fmt.cellSpacing())
    sp_width = QtWidgets.QDoubleSpinBox(dlg)
    sp_width.setRange(10.0, 100.0)
    sp_width.setSingleStep(5.0)
    # Pre-fill from existing width; fallback estimate if fixed
    try:
        wlen = fmt.width()
        # PyQt QTextLength exposes type() and value() or rawValue()
        wtype = wlen.type() if hasattr(wlen, "type") else None
        if wtype == wlen.PercentageLength:
            val = wlen.rawValue() if hasattr(wlen, "rawValue") else wlen.value()
            sp_width.setValue(float(val))
        else:
            # Approximate percent from viewport width
            vp = text_edit.viewport() if hasattr(text_edit, "viewport") else None
            vw = float(vp.width()) if vp is not None else 1.0
            v = wlen.value() if hasattr(wlen, "value") else 0.0
            pct = max(10.0, min(100.0, (float(v) / vw) * 100.0 if vw > 1.0 and v else 100.0))
            sp_width.setValue(pct)
    except Exception:
        sp_width.setValue(100.0)
    form.addRow("Border width:", sp_border)
    form.addRow("Cell padding:", sp_pad)
    form.addRow("Cell spacing:", sp_space)
    form.addRow("Table width (% of editor):", sp_width)
    btns = QtWidgets.QDialogButtonBox(
        QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, parent=dlg
    )
    form.addRow(btns)
    btns.accepted.connect(dlg.accept)
    btns.rejected.connect(dlg.reject)
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return
    fmt.setBorder(sp_border.value())
    fmt.setCellPadding(sp_pad.value())
    fmt.setCellSpacing(sp_space.value())
    try:
        from PyQt5.QtGui import QTextLength

        fmt.setWidth(QTextLength(QTextLength.PercentageLength, sp_width.value()))
    except Exception:
        pass
    table.setFormat(fmt)
    try:
        _enforce_uniform_table_borders(text_edit)
    except Exception:
        pass


def _table_add_remove(text_edit: QtWidgets.QTextEdit, action: str):
    cur = text_edit.textCursor()
    tbl = cur.currentTable()
    if tbl is None:
        return
    cell = tbl.cellAt(cur)
    if not cell.isValid():
        return
    row = cell.row()
    col = cell.column()
    if action == "row_above":
        tbl.insertRows(row, 1)
    elif action == "row_below":
        tbl.insertRows(row + 1, 1)
    elif action == "col_left":
        tbl.insertColumns(col, 1)
    elif action == "col_right":
        tbl.insertColumns(col + 1, 1)
    elif action == "remove_row":
        tbl.removeRows(row, 1)
    elif action == "remove_col":
        tbl.removeColumns(col, 1)


class _TableContextMenu(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit
        try:
            self._viewport = edit.viewport()
        except Exception:
            self._viewport = None
        # Track destruction to avoid using deleted C++ objects
        try:
            edit.destroyed.connect(self._on_dead)
        except Exception:
            pass
        try:
            if self._viewport is not None:
                self._viewport.destroyed.connect(self._on_dead)
        except Exception:
            pass

    def _on_dead(self, *args, **kwargs):
        self._edit = None
        self._viewport = None

    def eventFilter(self, obj, event):
        if self._edit is None:
            return False
        if (
            obj is self._edit or (self._viewport is not None and obj is self._viewport)
        ) and event.type() == QEvent.ContextMenu:
            pos = event.pos()
            # Position reported is in obj coords; map to the edit for cursor and to global for menu
            try:
                if obj is self._edit:
                    widget_pos = pos
                    global_pos = self._edit.mapToGlobal(pos)
                else:
                    # obj is viewport
                    widget_pos = pos  # QTextEdit accepts viewport coords for cursorForPosition
                    global_pos = obj.mapToGlobal(pos)
            except Exception:
                widget_pos = pos
                try:
                    global_pos = self._edit.mapToGlobal(pos)
                except Exception:
                    return False
            # Capture original selection and table BEFORE making any cursor changes
            orig_cur = self._edit.textCursor()
            orig_tbl = orig_cur.currentTable()
            # Compute selection rectangle from original selection (do not disturb selection)
            orig_rect = _table_selection_rect(self._edit, orig_tbl)

            # First priority: if click is on/near an image, show image menu and consume
            try:
                # Try detection chain (prioritize clicked position to avoid disturbing selection)
                info = _image_info_at_position(self._edit, widget_pos)
                if info is None:
                    # Fallbacks that don't require changing the current selection
                    try:
                        c_try = self._edit.cursorForPosition(widget_pos)
                        info = _image_info_near_doc_pos(self._edit, c_try.position())
                    except Exception:
                        pass
                if info is None:
                    try:
                        info = _image_info_in_block(self._edit, self._edit.textCursor().block(), prefer_pos=self._edit.textCursor().position())
                    except Exception:
                        pass
                if info is not None:
                    menu = QtWidgets.QMenu(self._edit)
                    cursor_pos = int(info.get("cursor_pos", 0))
                    name = info.get("name", "")
                    cur_w = float(info.get("w") or 0.0)
                    cur_h = float(info.get("h") or 0.0)
                    act_resize = menu.addAction("Image Properties…")
                    menu.addSeparator()
                    act_fit = menu.addAction("Fit to editor width")
                    act_reset = menu.addAction("Reset to original size")
                    chosen = menu.exec_(global_pos)
                    if chosen is not None:
                        if chosen == act_resize:
                            iw, ih = (None, None)
                            try:
                                iw, ih = _ImageContextMenuHandler(self._edit)._intrinsic_size(name)
                            except Exception:
                                pass
                            info_d = {"cursor_pos": cursor_pos, "name": name, "w": cur_w, "h": cur_h, "iw": iw, "ih": ih}
                            _image_properties_dialog_apply(self._edit, info_d)
                        elif chosen == act_fit:
                            try:
                                _ImageContextMenuHandler(self._edit)._fit_to_width(cursor_pos, name)
                            except Exception:
                                pass
                        elif chosen == act_reset:
                            try:
                                _ImageContextMenuHandler(self._edit)._reset_size(cursor_pos, name)
                            except Exception:
                                pass
                    return True
            except Exception:
                pass
            # Also capture clicked cell context
            try:
                clicked_cur = self._edit.cursorForPosition(widget_pos)
            except Exception:
                clicked_cur = orig_cur
            clicked_tbl = clicked_cur.currentTable()
            # Choose active table for the menu: prefer the one with a valid selection rect
            tbl = orig_tbl if (orig_tbl is not None and orig_rect is not None) else clicked_tbl
            # When not over a table, show a simple menu (Paste + Insert)
            if tbl is None:
                menu = QtWidgets.QMenu(self._edit)
                act_paste = menu.addAction("Paste")
                sub_ins = menu.addMenu("Insert")
                act_ins_table = sub_ins.addAction("Table…")
                # Insert Planning Register (dialog) under Insert submenu
                act_ins_pr_dialog = sub_ins.addAction("Planning Register…")
                chosen = menu.exec_(global_pos)
                if chosen is None:
                    return True
                if chosen == act_paste:
                    try:
                        from settings_manager import get_default_paste_mode

                        mode = get_default_paste_mode() or "rich"
                    except Exception:
                        mode = "rich"
                    try:
                        if mode == "text-only":
                            from ui_richtext import paste_text_only

                            paste_text_only(self._edit)
                        elif mode == "match-style":
                            from ui_richtext import paste_match_style

                            paste_match_style(self._edit)
                        elif mode == "clean":
                            from ui_richtext import paste_clean_formatting

                            paste_clean_formatting(self._edit)
                        else:
                            self._edit.paste()
                    except Exception:
                        try:
                            self._edit.paste()
                        except Exception:
                            pass
                    return True
                if chosen == act_ins_table:
                    _table_insert_dialog(self._edit)
                    return True
                if chosen == act_ins_pr_dialog:
                    try:
                        from ui_richtext import insert_planning_register_via_dialog

                        # window is parent of the editor; walk up to QMainWindow
                        w = self._edit.window()
                        insert_planning_register_via_dialog(w)
                    except Exception:
                        pass
                    return True
                return True
            # Otherwise, build the full table menu
            menu = QtWidgets.QMenu(self._edit)
            # Precompute selection rectangle early (for multi-column operations)
            sel_rect = _table_selection_rect(self._edit, tbl)
            # Insert submenu with Table and Planning Register
            sub_ins = menu.addMenu("Insert")
            act_ins = sub_ins.addAction("Table…")
            act_prop = menu.addAction("Table Properties…")
            act_fit = menu.addAction("Fit Table to Width")
            act_dist = menu.addAction("Distribute Columns Evenly")
            act_set_col = menu.addAction("Set Current Column Width…")
            menu.addSeparator()
            act_save_preset = menu.addAction("Save Table as Preset…")
            # Insert Planning Register (dialog) under Insert submenu
            act_ins_pr_dialog = sub_ins.addAction("Planning Register…")
            menu.addSeparator()
            # Currency column helpers
            act_mark_currency = menu.addAction("Mark Column(s) as Currency + Total")
            act_unmark_currency = menu.addAction("Clear Currency Formatting")
            menu.addSeparator()
            # Determine multi-cell selection rectangle (within chosen table)
            sel_rows = (sel_rect[2] - sel_rect[0] + 1) if sel_rect is not None else 1
            # Dynamic labels reflecting selection count
            act_row_above = menu.addAction(
                f"Insert Row{'s' if sel_rows>1 else ''} Above ({sel_rows})"
            )
            act_row_below = menu.addAction(
                f"Insert Row{'s' if sel_rows>1 else ''} Below ({sel_rows})"
            )
            act_col_left = menu.addAction("Insert Column Left")
            act_col_right = menu.addAction("Insert Column Right")
            act_rm_row = menu.addAction(f"Remove Selected Row{'s' if sel_rows>1 else ''}")
            act_rm_col = menu.addAction("Remove Column")
            act_clear_cells = menu.addAction("Clear Selected Cells")
            # Enable/disable depending on context
            has_tbl = tbl is not None
            act_prop.setEnabled(has_tbl)
            act_fit.setEnabled(has_tbl)
            act_dist.setEnabled(has_tbl)
            act_set_col.setEnabled(has_tbl)
            act_row_above.setEnabled(has_tbl)
            act_row_below.setEnabled(has_tbl)
            act_col_left.setEnabled(has_tbl)
            act_col_right.setEnabled(has_tbl)
            act_rm_row.setEnabled(has_tbl)
            act_rm_col.setEnabled(has_tbl)
            act_clear_cells.setEnabled(has_tbl and sel_rect is not None)
            chosen = menu.exec_(global_pos)
            if chosen is None:
                return True
            if chosen == act_ins:
                _table_insert_dialog(self._edit)
            elif chosen == act_ins_pr_dialog:
                try:
                    from ui_richtext import insert_planning_register_via_dialog

                    w = self._edit.window()
                    insert_planning_register_via_dialog(w)
                except Exception:
                    pass
            elif chosen == act_mark_currency and has_tbl:
                try:
                    _table_mark_currency_columns(self._edit, tbl, sel_rect)
                except Exception:
                    pass
            elif chosen == act_unmark_currency and has_tbl:
                try:
                    _table_unmark_currency_columns(self._edit, tbl, sel_rect)
                except Exception:
                    pass
            elif chosen == act_prop and has_tbl:
                _table_properties_dialog(self._edit, tbl)
            elif chosen == act_fit and has_tbl:
                _table_fit_width(tbl)
            elif chosen == act_dist and has_tbl:
                _table_distribute_columns(tbl)
            elif chosen == act_set_col and has_tbl:
                # For column-based actions, position the caret to clicked cell to define the column
                try:
                    self._edit.setTextCursor(clicked_cur)
                except Exception:
                    pass
                _table_set_current_column_width(self._edit, tbl)
            elif chosen == act_save_preset and has_tbl:
                # Use the centralized HTML-based saver so data and styles are preserved
                try:
                    save_current_table_as_preset(self._edit)
                except Exception:
                    pass
            elif has_tbl:
                if chosen == act_row_above:
                    _table_insert_rows_from_selection(self._edit, tbl, sel_rect, above=True)
                elif chosen == act_row_below:
                    _table_insert_rows_from_selection(self._edit, tbl, sel_rect, above=False)
                elif chosen == act_col_left:
                    try:
                        self._edit.setTextCursor(clicked_cur)
                    except Exception:
                        pass
                    _table_add_remove(self._edit, "col_left")
                elif chosen == act_col_right:
                    try:
                        self._edit.setTextCursor(clicked_cur)
                    except Exception:
                        pass
                    _table_add_remove(self._edit, "col_right")
                elif chosen == act_rm_row:
                    _table_remove_rows_from_selection(self._edit, tbl, sel_rect)
                elif chosen == act_rm_col:
                    try:
                        self._edit.setTextCursor(clicked_cur)
                    except Exception:
                        pass
                    _table_add_remove(self._edit, "remove_col")
                elif chosen == act_clear_cells and sel_rect is not None:
                    _table_clear_selected_cells(self._edit, tbl, sel_rect)
            return True
        return super().eventFilter(obj, event)


def _install_table_context_menu(text_edit: QtWidgets.QTextEdit):
    handler = _TableContextMenu(text_edit)
    text_edit.installEventFilter(handler)
    try:
        vp = text_edit.viewport()
        if vp is not None:
            vp.installEventFilter(handler)
    except Exception:
        pass
    if not hasattr(text_edit, "_tableCtx"):  # keep references
        text_edit._tableCtx = []
    text_edit._tableCtx.append(handler)


def _table_selection_rect(text_edit: QtWidgets.QTextEdit, table):
    """Return (r0,c0,r1,c1) rectangle for current selection within table; None if selection not in table."""
    try:
        if table is None:
            return None
        cur = text_edit.textCursor()
        a = cur.anchor()
        p = cur.position()
        c1 = table.cellAt(min(a, p))
        c2 = table.cellAt(max(a, p))
        if not c1.isValid() or not c2.isValid():
            # If no selection, use current cell
            c = table.cellAt(cur)
            if not c.isValid():
                return None
            r = c.row()
            ccol = c.column()
            return (r, ccol, r, ccol)
        r0 = min(c1.row(), c2.row())
        r1 = max(c1.row(), c2.row())
        c0 = min(c1.column(), c2.column())
        c1i = max(c1.column(), c2.column())
        return (r0, c0, r1, c1i)
    except Exception:
        return None


def _table_clear_selected_cells(text_edit: QtWidgets.QTextEdit, table, rect):
    if rect is None:
        return
    r0, c0, r1, c1 = rect
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = table.cellAt(r, c)
            if not cell.isValid():
                continue
            tc = cell.firstCursorPosition()
            # Select the cell range
            last = cell.lastCursorPosition()
            tc.setPosition(last.position(), QTextCursor.KeepAnchor)
            try:
                tc.removeSelectedText()
            except Exception:
                pass


def _table_insert_rows_from_selection(text_edit: QtWidgets.QTextEdit, table, rect, above: bool):
    cur = text_edit.textCursor()
    cell = table.cellAt(cur)
    if not cell.isValid():
        return
    base_row = cell.row()
    count = 1
    if rect is not None:
        r0, _c0, r1, _c1 = rect
        count = max(1, r1 - r0 + 1)
        base_row = r0 if above else (r1 + 1)
    # Capture context about header/totals position before insert
    try:
        rows_before = table.rows()
    except Exception:
        rows_before = None
    totals_before = (rows_before - 1) if rows_before is not None and rows_before > 0 else None
    try:
        table.insertRows(base_row, count)
        # If inserting immediately before header (row 0) or immediately before previous totals row,
        # clear background on the newly inserted rows so they appear as normal data rows.
        if base_row == 0 or (totals_before is not None and base_row == totals_before):
            try:
                cols = table.columns()
                for rr in range(base_row, base_row + count):
                    for cc in range(cols):
                        c = table.cellAt(rr, cc)
                        if c.isValid():
                            cf = c.format()
                            try:
                                # Clear background to transparent so the row looks like an interior data row
                                cf.setBackground(QColor(0, 0, 0, 0))
                            except Exception:
                                try:
                                    cf.setBackground(Qt.transparent)
                                except Exception:
                                    pass
                            c.setFormat(cf)
            except Exception:
                pass
        # For Planning Register tables and Cost List tables, ensure numeric columns in the
        # newly inserted rows are right-aligned so the caret appears on the right in empty cells.
        try:
            from ui_planning_register import _is_planning_register_table, _is_cost_list_table, _is_protected_cell

            bf = QTextBlockFormat()
            bf.setAlignment(Qt.AlignRight)
            rows_total = table.rows()
            r_start = max(0, base_row)
            r_end = min(rows_total - 1, base_row + count - 1)
            if _is_planning_register_table(text_edit, table):
                # Skip header (row 0) and totals (last row)
                for rr in range(r_start, r_end + 1):
                    if rr == 0 or rr == (rows_total - 1):
                        continue
                    for cc in (1, 2):
                        try:
                            if _is_protected_cell(table, rr, cc):
                                continue
                        except Exception:
                            pass
                        cell = table.cellAt(rr, cc)
                        if cell.isValid():
                            tcur = cell.firstCursorPosition()
                            tcur.mergeBlockFormat(bf)
            elif _is_cost_list_table(text_edit, table):
                for rr in range(r_start, r_end + 1):
                    # Skip header row 0
                    if rr == 0:
                        continue
                    cc = 1
                    cell = table.cellAt(rr, cc)
                    if cell.isValid():
                        tcur = cell.firstCursorPosition()
                        tcur.mergeBlockFormat(bf)
        except Exception:
            pass
    except Exception:
        pass


def _table_remove_rows_from_selection(text_edit: QtWidgets.QTextEdit, table, rect):
    cur = text_edit.textCursor()
    cell = table.cellAt(cur)
    if rect is None or not cell.isValid():
        # Remove current row
        try:
            table.removeRows(cell.row(), 1)
        except Exception:
            pass
        _table_delete_if_empty(text_edit, table)
        return
    r0, _c0, r1, _c1 = rect
    try:
        table.removeRows(r0, r1 - r0 + 1)
    except Exception:
        pass
    _table_delete_if_empty(text_edit, table)


def _table_delete_if_empty(text_edit: QtWidgets.QTextEdit, table):
    try:
        if table.rows() > 0:
            return
        # Select the entire table frame and replace with a blank block
        start = table.firstPosition()
        end = table.lastPosition()
        c = QTextCursor(text_edit.document())
        c.setPosition(start)
        c.setPosition(end, QTextCursor.KeepAnchor)
        c.removeSelectedText()
        # Ensure there's a paragraph to continue typing
        c.insertBlock()
    except Exception:
        pass


def _table_fit_width(table):
    from PyQt5.QtGui import QTextLength, QTextTableFormat

    try:
        fmt = table.format()
        fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
        table.setFormat(fmt)
    except Exception:
        pass


def _table_distribute_columns(table):
    from PyQt5.QtGui import QTextLength

    try:
        cols = table.columns()
        if cols <= 0:
            return
        frac = 100.0 / float(cols)
        fmt = table.format()
        fmt.setColumnWidthConstraints(
            [QTextLength(QTextLength.PercentageLength, frac) for _ in range(cols)]
        )
        table.setFormat(fmt)
    except Exception:
        pass


def _table_set_current_column_width(text_edit: QtWidgets.QTextEdit, table):
    from PyQt5.QtGui import QTextLength

    cur = text_edit.textCursor()
    cell = table.cellAt(cur)
    if not cell.isValid():
        return
    col_idx = cell.column()
    cols = table.columns()
    fmt = table.format()
    constraints = list(fmt.columnWidthConstraints()) or []
    if not constraints or len(constraints) != cols:
        constraints = [QTextLength(QTextLength.PercentageLength, 100.0 / cols) for _ in range(cols)]
    # Ask user for percentage
    dlg = QtWidgets.QInputDialog(text_edit)
    dlg.setWindowTitle("Set Column Width")
    dlg.setLabelText(f"Width for column {col_idx+1} (% of table width):")
    dlg.setInputMode(QtWidgets.QInputDialog.DoubleInput)
    dlg.setDoubleRange(1.0, 100.0)
    dlg.setDoubleDecimals(1)
    dlg.setDoubleValue(
        constraints[col_idx].rawValue()
        if hasattr(constraints[col_idx], "rawValue")
        else constraints[col_idx].value()
    )
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return
    new_pct = dlg.doubleValue()
    # Rebalance other columns to keep sum ~100
    other_indices = [i for i in range(cols) if i != col_idx]
    current_sum = sum((c.rawValue() if hasattr(c, "rawValue") else c.value()) for c in constraints)
    remaining = max(1.0, 100.0 - new_pct)
    if not other_indices:
        constraints[col_idx] = QTextLength(QTextLength.PercentageLength, new_pct)
    else:
        # Distribute remaining proportionally to their existing sizes
        other_sum = max(
            1e-6,
            sum(
                (
                    constraints[i].rawValue()
                    if hasattr(constraints[i], "rawValue")
                    else constraints[i].value()
                )
                for i in other_indices
            ),
        )
        for i in other_indices:
            base = (
                constraints[i].rawValue()
                if hasattr(constraints[i], "rawValue")
                else constraints[i].value()
            )
            pct = remaining * (base / other_sum)
            constraints[i] = QTextLength(QTextLength.PercentageLength, pct)
        constraints[col_idx] = QTextLength(QTextLength.PercentageLength, new_pct)
    fmt.setColumnWidthConstraints(constraints)
    table.setFormat(fmt)


# ----------------------------- Table formulas (SUM) -----------------------------
def _letters_to_index(letters: str) -> int:
    """Convert spreadsheet-like column letters (A, B, ... Z, AA, AB, ...) to 0-based index."""
    s = (letters or "").strip().upper()
    if not s or not s.isalpha():
        return -1
    idx = 0
    for ch in s:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _parse_cell_address(addr: str):
    """Parse like A1 -> (row_idx, col_idx) 0-based. Returns (r, c) or (None, None) if invalid."""
    if not isinstance(addr, str):
        return None, None
    m = re.match(r"^\s*([A-Za-z]+)(\d+)\s*$", addr)
    if not m:
        return None, None
    letters, row_str = m.group(1), m.group(2)
    col = _letters_to_index(letters)
    try:
        row = int(row_str) - 1
    except Exception:
        row = -1
    if row < 0 or col < 0:
        return None, None
    return row, col


def _table_cell_plain_text(table, row: int, col: int) -> str:
    try:
        cell = table.cellAt(row, col)
        if not cell.isValid():
            return ""
        c = cell.firstCursorPosition()
        # select to last
        last = cell.lastCursorPosition()
        c.setPosition(last.position(), QTextCursor.KeepAnchor)
        return c.selectedText()
    except Exception:
        return ""


def _table_set_cell_plain_text(text_edit: QtWidgets.QTextEdit, table, row: int, col: int, text: str):
    try:
        cell = table.cellAt(row, col)
        if not cell.isValid():
            return
        c = cell.firstCursorPosition()
        last = cell.lastCursorPosition()
        c.beginEditBlock()
        try:
            c.setPosition(last.position(), QTextCursor.KeepAnchor)
            try:
                c.removeSelectedText()
            except Exception:
                pass
            c.insertText(str(text))
        finally:
            c.endEditBlock()
    except Exception:
        pass


def _sum_range_in_table(table, start_addr: str, end_addr: str) -> float:
    r0, c0 = _parse_cell_address(start_addr)
    r1, c1 = _parse_cell_address(end_addr)
    if r0 is None or c0 is None or r1 is None or c1 is None:
        return 0.0
    r0, r1 = min(r0, r1), max(r0, r1)
    c0, c1 = min(c0, c1), max(c0, c1)
    total = 0.0
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            txt = _table_cell_plain_text(table, r, c).strip()
            # If a referenced cell contains a formula, ignore it during SUM
            if txt.startswith("="):
                continue
            # Try parsing as float, permissive of commas
            try:
                num = float(txt.replace(",", "")) if txt else 0.0
                total += num
            except Exception:
                # ignore non-numeric
                pass
    return total


# (Removed legacy inline SUM formula recalculation support.)

# --- Currency Column Helpers ---
_CURRENCY_SUFFIX = " (Currency)"

def _format_currency(value: float) -> str:
    try:
        # Show negative values with leading minus (simple style)
        return ("-$%s" % f"{abs(value):,.2f}") if value < 0 else f"${value:,.2f}"
    except Exception:
        return str(value)

def _detect_currency_columns(table) -> set:
    cols = set()
    try:
        if table.rows() == 0:
            return cols
        cols_count = table.columns()
        for c in range(cols_count):
            hdr_txt = _table_cell_plain_text(table, 0, c)
            if isinstance(hdr_txt, str) and hdr_txt.endswith(_CURRENCY_SUFFIX):
                cols.add(c)
    except Exception:
        pass
    return cols

def _table_mark_currency_columns(text_edit: QtWidgets.QTextEdit, table, sel_rect):
    if table is None:
        return
    # Determine columns to mark: from selection if multi-cell, else current cell
    cols_to_mark = set()
    try:
        if sel_rect is not None:
            r0, r1, c0, c1 = sel_rect
            for c in range(c0, c1 + 1):
                cols_to_mark.add(c)
        else:
            cur = text_edit.textCursor()
            cell = table.cellAt(cur)
            if cell.isValid():
                cols_to_mark.add(cell.column())
    except Exception:
        pass
    if not cols_to_mark:
        return
    # Append suffix to header cells and right-align numeric cells; also format existing numeric entries
    rows = table.rows()
    for c in cols_to_mark:
        # Header cell suffix
        try:
            hdr_txt = _table_cell_plain_text(table, 0, c)
            if hdr_txt and not hdr_txt.endswith(_CURRENCY_SUFFIX):
                _table_set_cell_plain_text(text_edit, table, 0, c, hdr_txt + _CURRENCY_SUFFIX)
        except Exception:
            pass
        # Align all cells in column (excluding header maybe) to right
        for r in range(1, rows):
            try:
                cell = table.cellAt(r, c)
                if not cell.isValid():
                    continue
                cur = cell.firstCursorPosition()
                bf = cur.blockFormat()
                from PyQt5.QtCore import Qt as _Qt
                bf.setAlignment(_Qt.AlignRight)
                cur.setBlockFormat(bf)
                # Format numeric cell content as currency if parseable
                raw = _table_cell_plain_text(table, r, c)
                if raw:
                    cleaned = raw.replace("$", "").replace(",", "").strip()
                    # Allow leading minus
                    try:
                        val = float(cleaned)
                        _table_set_cell_plain_text(text_edit, table, r, c, _format_currency(val))
                    except Exception:
                        pass
            except Exception:
                pass
    # Ensure a total row exists (last row). If last row header cell text equals 'Total', reuse.
    try:
        need_total_row = True
        last_row_idx = rows - 1 if rows > 0 else -1
        if last_row_idx >= 0:
            first_cell_txt = _table_cell_plain_text(table, last_row_idx, 0)
            if isinstance(first_cell_txt, str) and first_cell_txt.strip().lower() == "total":
                need_total_row = False
        if need_total_row:
            table.appendRows(1)
            last_row_idx = table.rows() - 1
            _table_set_cell_plain_text(text_edit, table, last_row_idx, 0, "Total")
        # Compute totals for newly marked columns immediately
        for c in cols_to_mark:
            total = 0.0
            for r in range(1, last_row_idx):  # exclude header and total row
                try:
                    raw = _table_cell_plain_text(table, r, c)
                    if not raw:
                        continue
                    # Strip currency symbols/commas
                    cleaned = raw.replace("$", "").replace(",", "").strip()
                    if cleaned.endswith(_CURRENCY_SUFFIX):
                        cleaned = cleaned[:-len(_CURRENCY_SUFFIX)]
                    val = float(cleaned) if cleaned else 0.0
                    total += val
                except Exception:
                    pass
            _table_set_cell_plain_text(text_edit, table, last_row_idx, c, _format_currency(total))
            # Right-align total cell
            try:
                cell = table.cellAt(last_row_idx, c)
                if cell.isValid():
                    cur = cell.firstCursorPosition()
                    bf = cur.blockFormat()
                    from PyQt5.QtCore import Qt as _Qt
                    bf.setAlignment(_Qt.AlignRight)
                    cur.setBlockFormat(bf)
            except Exception:
                pass
    except Exception:
        pass

def _table_unmark_currency_columns(text_edit: QtWidgets.QTextEdit, table, sel_rect):
    if table is None or table.rows() == 0:
        return
    cols_all = _detect_currency_columns(table)
    if not cols_all:
        return
    # Determine target columns via selection; if no selection, unmark all currency columns
    target = set()
    try:
        if sel_rect is not None:
            r0, r1, c0, c1 = sel_rect
            for c in range(c0, c1 + 1):
                if c in cols_all:
                    target.add(c)
        else:
            target = cols_all
    except Exception:
        target = cols_all
    if not target:
        return
    rows = table.rows()
    for c in target:
        # Remove suffix from header
        try:
            hdr_txt = _table_cell_plain_text(table, 0, c)
            if hdr_txt and hdr_txt.endswith(_CURRENCY_SUFFIX):
                base = hdr_txt[:-len(_CURRENCY_SUFFIX)]
                _table_set_cell_plain_text(text_edit, table, 0, c, base)
        except Exception:
            pass
    # Optionally clear totals (blank last row cells for these columns)
    try:
        last_row_idx = rows - 1
        if last_row_idx > 0:
            first_cell_txt = _table_cell_plain_text(table, last_row_idx, 0)
            if isinstance(first_cell_txt, str) and first_cell_txt.strip().lower() == "total":
                for c in target:
                    _table_set_cell_plain_text(text_edit, table, last_row_idx, c, "")
    except Exception:
        pass

def _table_recompute_currency_columns(text_edit: QtWidgets.QTextEdit, table):
    if table is None or table.rows() < 2:
        return
    cols = _detect_currency_columns(table)
    if not cols:
        return
    last_row_idx = table.rows() - 1
    first_cell_txt = _table_cell_plain_text(table, last_row_idx, 0)
    has_total_row = isinstance(first_cell_txt, str) and first_cell_txt.strip().lower() == "total"
    if not has_total_row:
        # Append total row if missing
        table.appendRows(1)
        last_row_idx = table.rows() - 1
        _table_set_cell_plain_text(text_edit, table, last_row_idx, 0, "Total")
    # Recompute totals
    for c in cols:
        total = 0.0
        for r in range(1, last_row_idx):
            try:
                raw = _table_cell_plain_text(table, r, c)
                if not raw:
                    continue
                cleaned = raw.replace("$", "").replace(",", "").strip()
                val = float(cleaned) if cleaned else 0.0
                total += val
                # Reformat cell display as currency
                _table_set_cell_plain_text(text_edit, table, r, c, _format_currency(val))
            except Exception:
                pass
        _table_set_cell_plain_text(text_edit, table, last_row_idx, c, _format_currency(total))
        try:
            cell = table.cellAt(last_row_idx, c)
            if cell.isValid():
                cur = cell.firstCursorPosition()
                bf = cur.blockFormat()
                from PyQt5.QtCore import Qt as _Qt
                bf.setAlignment(_Qt.AlignRight)
                cur.setBlockFormat(bf)
        except Exception:
            pass


def ensure_currency_columns_watcher(text_edit: QtWidgets.QTextEdit):
    if text_edit is None or not isinstance(text_edit, QtWidgets.QTextEdit):
        return
    if hasattr(text_edit, "_currency_columns_watcher_active"):
        return
    text_edit._currency_columns_watcher_active = True
    text_edit._currency_last_cell = None

    def _on_cursor_changed():
        try:
            cur = text_edit.textCursor()
            tbl = cur.currentTable()
            prev = text_edit._currency_last_cell
            if tbl is not None:
                cell = tbl.cellAt(cur)
                if cell.isValid():
                    coord = (tbl, cell.row(), cell.column())
                    if prev is None:
                        text_edit._currency_last_cell = coord
                    elif prev != coord:
                        prev_tbl = prev[0]
                        if prev_tbl is not None and _detect_currency_columns(prev_tbl):
                            _table_recompute_currency_columns(text_edit, prev_tbl)
                        text_edit._currency_last_cell = coord
                else:
                    if prev is not None:
                        prev_tbl = prev[0]
                        if prev_tbl is not None and _detect_currency_columns(prev_tbl):
                            _table_recompute_currency_columns(text_edit, prev_tbl)
                    text_edit._currency_last_cell = None
            else:
                if prev is not None:
                    prev_tbl = prev[0]
                    if prev_tbl is not None and _detect_currency_columns(prev_tbl):
                        _table_recompute_currency_columns(text_edit, prev_tbl)
                text_edit._currency_last_cell = None
        except Exception:
            pass

    text_edit.cursorPositionChanged.connect(_on_cursor_changed)


def _apply_selection_colors(text_edit: QtWidgets.QTextEdit, bg: QColor, fg: QColor):
    pal = text_edit.palette()
    for group in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
        pal.setColor(group, QPalette.Highlight, bg)
        pal.setColor(group, QPalette.HighlightedText, fg)
    text_edit.setPalette(pal)
    # Stylesheet fallback (some styles prefer stylesheet roles)
    try:
        text_edit.setStyleSheet(
            "QTextEdit { selection-background-color: %s; selection-color: %s; }"
            % (bg.name(), fg.name())
        )
    except Exception:
        pass


# ----------------------------- Formatting helpers -----------------------------
def _apply_font_family(text_edit: QtWidgets.QTextEdit, family: str):
    if not family:
        return
    fmt = QTextCharFormat()
    fmt.setFontFamily(str(family))
    cursor = text_edit.textCursor()
    if not cursor.hasSelection():
        cursor.select(cursor.WordUnderCursor)
    cursor.mergeCharFormat(fmt)
    text_edit.mergeCurrentCharFormat(fmt)


def _apply_font_size(text_edit: QtWidgets.QTextEdit, size_pt: float):
    try:
        if not size_pt:
            return
        size_f = float(size_pt)
    except Exception:
        return
    fmt = QTextCharFormat()
    fmt.setFontPointSize(size_f)
    cursor = text_edit.textCursor()
    if not cursor.hasSelection():
        cursor.select(cursor.WordUnderCursor)
    cursor.mergeCharFormat(fmt)
    text_edit.mergeCurrentCharFormat(fmt)


def _apply_text_color(text_edit: QtWidgets.QTextEdit, foreground: bool = True):
    """Open a color dialog and apply the chosen color to foreground or background of selection."""
    dlg = QtWidgets.QColorDialog(text_edit)
    try:
        dlg.setOption(QtWidgets.QColorDialog.ShowAlphaChannel, False)
    except Exception:
        pass
    if dlg.exec_() == QtWidgets.QDialog.Accepted:
        color = dlg.selectedColor()
        fmt = QTextCharFormat()
        if foreground:
            fmt.setForeground(color)
        else:
            fmt.setBackground(color)
        cursor = text_edit.textCursor()
        if not cursor.hasSelection():
            cursor.select(cursor.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        text_edit.mergeCurrentCharFormat(fmt)


def _clear_background(text_edit: QtWidgets.QTextEdit):
    fmt = QTextCharFormat()
    try:
        fmt.setBackground(Qt.transparent)
    except Exception:
        fmt.clearBackground()
    cursor = text_edit.textCursor()
    if not cursor.hasSelection():
        cursor.select(cursor.WordUnderCursor)
    cursor.mergeCharFormat(fmt)
    # Also update typing format so new text doesn't carry highlight
    text_edit.mergeCurrentCharFormat(fmt)


# ----------------------------- Lists and indentation -----------------------------
def _toggle_list(text_edit: QtWidgets.QTextEdit, ordered: bool):
    cursor = text_edit.textCursor()
    block = cursor.block()
    cur_list = block.textList()
    if cur_list is not None:
        # If same list kind, remove list formatting
        is_ordered = _is_ordered_style(cur_list.format().style())
        if (ordered and is_ordered) or ((not ordered) and (not is_ordered)):
            bf = block.blockFormat()
            bf.setObjectIndex(-1)
            cursor.mergeBlockFormat(bf)
            return
    # Otherwise create list of requested type at level 1 or keep existing level
    level = 1
    if cur_list is not None:
        try:
            level = max(1, cur_list.format().indent())
        except Exception:
            level = 1
    lf = QTextListFormat()
    lf.setIndent(level)
    lf.setStyle(_ordered_style_for_level(level) if ordered else _unordered_style_for_level(level))
    new_list = cursor.createList(lf)
    _merge_with_adjacent_lists(cursor, new_list)


def _change_list_indent(text_edit: QtWidgets.QTextEdit, delta: int):
    cursor = text_edit.textCursor()
    block = cursor.block()
    cur_list = block.textList()
    if cur_list is None:
        # Fallback: adjust plain paragraph left margin when not in a list
        try:
            step = INDENT_STEP_PX
        except Exception:
            step = 24.0
        _change_block_left_margin(text_edit, float(delta) * float(step))
        return
    cur_level = max(1, cur_list.format().indent())
    new_level = max(1, cur_level + int(delta))
    # Build new list format for this block only
    ordered = _is_ordered_style(cur_list.format().style())
    lf = QTextListFormat()
    lf.setIndent(new_level)
    lf.setStyle(
        _ordered_style_for_level(new_level) if ordered else _unordered_style_for_level(new_level)
    )
    # Remove this block from its current list by clearing object index
    bf = block.blockFormat()
    bf.setObjectIndex(-1)
    cursor.beginEditBlock()
    cursor.mergeBlockFormat(bf)
    # Recreate as its own list with new indent, then merge
    cursor.createList(lf)
    cursor.endEditBlock()
    # Merge contiguous lists to maintain numbering
    nb = cursor.block().textList()
    if nb is not None:
        _merge_with_adjacent_lists(cursor, nb)


class _ListTabHandler(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit

    def eventFilter(self, obj, event):
        if obj is self._edit and event.type() == QEvent.KeyPress:
            key = event.key()
            if key in (Qt.Key_Tab, Qt.Key_Backtab):
                cur = self._edit.textCursor()
                if cur.block().textList() is not None:
                    is_backtab = (key == Qt.Key_Backtab) or bool(
                        event.modifiers() & Qt.ShiftModifier
                    )
                    _change_list_indent(self._edit, -1 if is_backtab else +1)
                    return True  # consume to avoid inserting a tab char
        return super().eventFilter(obj, event)


def _install_list_tab_handler(text_edit: QtWidgets.QTextEdit):
    handler = _ListTabHandler(text_edit)
    text_edit.installEventFilter(handler)
    # Keep a reference to prevent GC
    if not hasattr(text_edit, "_listTabHandler"):
        text_edit._listTabHandler = []
    text_edit._listTabHandler.append(handler)


class _TableTabHandler(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit

    def eventFilter(self, obj, event):
        if obj is self._edit and event.type() == QEvent.KeyPress:
            key = event.key()
            if key in (Qt.Key_Tab, Qt.Key_Backtab):
                shift = (key == Qt.Key_Backtab) or bool(event.modifiers() & Qt.ShiftModifier)
                cur = self._edit.textCursor()
                tbl = cur.currentTable()
                if tbl is None:
                    return False
                cell = tbl.cellAt(cur)
                if not cell.isValid():
                    return False
                row = cell.row()
                col = cell.column()
                rows = tbl.rows()
                cols = tbl.columns()
                if shift:
                    # Move backward
                    prev_col = col - 1
                    prev_row = row
                    if prev_col < 0:
                        prev_row -= 1
                        if prev_row < 0:
                            return True  # swallow at very start
                        prev_col = cols - 1
                    target = tbl.cellAt(prev_row, prev_col)
                    self._edit.setTextCursor(target.firstCursorPosition())
                    return True
                # Forward
                next_col = col + 1
                next_row = row
                if next_col >= cols:
                    next_col = 0
                    next_row += 1
                    if next_row >= rows:
                        # Append a new row
                        try:
                            tbl.insertRows(rows, 1)
                            rows += 1
                        except Exception:
                            pass
                if next_row < rows:
                    target = tbl.cellAt(next_row, next_col)
                    self._edit.setTextCursor(target.firstCursorPosition())
                return True
        return super().eventFilter(obj, event)


def _install_table_tab_handler(text_edit: QtWidgets.QTextEdit):
    handler = _TableTabHandler(text_edit)
    text_edit.installEventFilter(handler)
    if not hasattr(text_edit, "_tableTabHandlers"):
        text_edit._tableTabHandlers = []
    text_edit._tableTabHandlers.append(handler)


# ----------------------------- Plain paragraph indent with Tab/Shift+Tab -----------------------------
INDENT_STEP_PX = 24.0


def _change_block_left_margin(text_edit: QtWidgets.QTextEdit, delta_px: float):
    cur = text_edit.textCursor()
    start = cur.selectionStart()
    end = cur.selectionEnd()
    work = QTextCursor(cur)
    work.beginEditBlock()
    try:
        # If no selection, apply to current block
        if start == end:
            block = cur.block()
            bf = block.blockFormat()
            bf.setLeftMargin(max(0.0, float(bf.leftMargin()) + float(delta_px)))
            cur.mergeBlockFormat(bf)
        else:
            c = QTextCursor(text_edit.document())
            c.setPosition(start)
            while True:
                block = c.block()
                bf = block.blockFormat()
                bf.setLeftMargin(max(0.0, float(bf.leftMargin()) + float(delta_px)))
                # Select this block and apply
                bc = QTextCursor(block)
                bc.select(QTextCursor.BlockUnderCursor)
                bc.mergeBlockFormat(bf)
                # Stop if we've passed end
                if block.position() + block.length() >= end:
                    break
                if not c.movePosition(QTextCursor.NextBlock):
                    break
    finally:
        work.endEditBlock()


class _PlainIndentTabHandler(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit

    def eventFilter(self, obj, event):
        if obj is self._edit and event.type() == QEvent.KeyPress:
            key = event.key()
            if key in (Qt.Key_Tab, Qt.Key_Backtab):
                cur = self._edit.textCursor()
                # Skip if inside table or list; handled by other filters
                if cur.currentTable() is not None or cur.block().textList() is not None:
                    return False
                is_out = (key == Qt.Key_Backtab) or bool(event.modifiers() & Qt.ShiftModifier)
                _change_block_left_margin(
                    self._edit, -INDENT_STEP_PX if is_out else +INDENT_STEP_PX
                )
                return True
        return super().eventFilter(obj, event)


def _install_plain_indent_tab_handler(text_edit: QtWidgets.QTextEdit):
    handler = _PlainIndentTabHandler(text_edit)
    text_edit.installEventFilter(handler)
    if not hasattr(text_edit, "_plainIndentHandlers"):
        text_edit._plainIndentHandlers = []
    text_edit._plainIndentHandlers.append(handler)
    # Load indent step from settings if available
    try:
        from settings_manager import get_plain_indent_px

        global INDENT_STEP_PX
        INDENT_STEP_PX = float(get_plain_indent_px())
    except Exception:
        pass


def _select_combo_value(combo: QtWidgets.QComboBox, value: int):
    try:
        target = int(value)
    except Exception:
        return
    for i in range(combo.count()):
        try:
            if int(combo.itemData(i)) == target:
                combo.setCurrentIndex(i)
                return
        except Exception:
            continue


def _install_default_paste_override(text_edit: QtWidgets.QTextEdit):
    class _PasteHandler(QObject):
        def __init__(self, edit):
            super().__init__(edit)
            self._edit = edit

        def _mode(self) -> str:
            try:
                from settings_manager import get_default_paste_mode

                return get_default_paste_mode() or "rich"
            except Exception:
                return "rich"

        def eventFilter(self, obj, event):
            if obj is self._edit and event.type() == QEvent.KeyPress:
                if (
                    event.key() == Qt.Key_V
                    and (event.modifiers() & Qt.ControlModifier)
                    and not (event.modifiers() & Qt.AltModifier)
                ):
                    mode = self._mode()
                    if mode == "text-only":
                        paste_text_only(self._edit)
                    elif mode == "match-style":
                        paste_match_style(self._edit)
                    elif mode == "clean":
                        paste_clean_formatting(self._edit)
                    else:
                        self._edit.paste()
                    return True
            return super().eventFilter(obj, event)

    handler = _PasteHandler(text_edit)
    text_edit.installEventFilter(handler)
    if not hasattr(text_edit, "_pasteHandler"):
        text_edit._pasteHandler = []
    text_edit._pasteHandler.append(handler)


def paste_clean_formatting(text_edit: QtWidgets.QTextEdit):
    """Paste rich text but drop most inline styles/classes and normalize to current font family/size.
    Keeps structure like paragraphs, links, images, and lists.
    """
    cb = QtWidgets.QApplication.clipboard()
    md = cb.mimeData()
    if md is None:
        return
    html = None
    if md.hasHtml():
        html = md.html()
    elif md.hasText():
        txt = md.text().strip()
        if _looks_like_url(txt):
            url = _normalize_url_scheme(txt)
            html = f'<a href="{url}">{txt}</a>'
        else:
            html = txt.replace("\n", "<br/>")
    if not html:
        return
    try:
        s = _strip_match_style_html(html)
        # Additionally drop any remaining style/class attributes outright
        import re

        s = re.sub(r'\sstyle\s*=\s*"[^"]*"', "", s, flags=re.IGNORECASE)
        s = re.sub(r"\sstyle\s*=\s*'[^']*'", "", s, flags=re.IGNORECASE)
        s = re.sub(r'\sclass\s*=\s*"[^"]*"', "", s, flags=re.IGNORECASE)
        s = re.sub(r"\sclass\s*=\s*'[^']*'", "", s, flags=re.IGNORECASE)
        cleaned = s
    except Exception:
        cleaned = html
    cursor = text_edit.textCursor()
    pre_fmt = text_edit.currentCharFormat()
    before = cursor.position()
    cursor.insertHtml(cleaned)
    after = cursor.position()
    # Normalize inserted range to current family/size and transparent background
    doc_font = _effective_default_font(text_edit)
    fmt = QTextCharFormat(pre_fmt)
    fam = fmt.fontFamily() or doc_font.family()
    if fam:
        fmt.setFontFamily(fam)
    sz = fmt.fontPointSize()
    if not sz or sz <= 0:
        try:
            sz = doc_font.pointSizeF() if doc_font.pointSizeF() > 0 else float(doc_font.pointSize())
        except Exception:
            sz = 12.0
    fmt.setFontPointSize(float(sz))
    try:
        fmt.setBackground(Qt.transparent)
    except Exception:
        fmt.clearBackground()
    rng = QTextCursor(text_edit.document())
    rng.setPosition(before)
    rng.setPosition(after, QTextCursor.KeepAnchor)
    rng.mergeCharFormat(fmt)
    # Restore typing format
    restored = QTextCharFormat(pre_fmt)
    try:
        restored.setBackground(Qt.transparent)
    except Exception:
        restored.clearBackground()
    text_edit.setCurrentCharFormat(restored)
    cursor.setPosition(after)
    text_edit.setTextCursor(cursor)


# ----------------------------- Link handling -----------------------------
def _looks_like_url(text: str) -> bool:
    if not text:
        return False
    t = text.strip()
    # Quick heuristics for URLs and mailto
    if t.lower().startswith(("http://", "https://", "mailto:")):
        return True
    if t.lower().startswith("www.") and " " not in t:
        return True
    # basic domain.tld pattern without spaces
    return ("." in t and " " not in t and "/" in t) or (
        "." in t and t.count(".") >= 1 and " " not in t
    )


def _normalize_url_scheme(text: str) -> str:
    t = (text or "").strip()
    if not t:
        return t
    if t.lower().startswith(("http://", "https://", "mailto:")):
        return t
    if t.lower().startswith("www."):
        return "http://" + t
    return t


class _LinkClickHandler(QObject):
    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit
        self._viewport = edit.viewport() if edit is not None else None
        self._pressed_anchor = None
        # Mark dead when either the edit or its viewport is destroyed
        try:
            edit.destroyed.connect(self._on_dead)
        except Exception:
            pass
        try:
            if self._viewport is not None:
                self._viewport.destroyed.connect(self._on_dead)
        except Exception:
            pass

    def _on_dead(self, *args, **kwargs):
        self._edit = None
        self._viewport = None
        self._pressed_anchor = None

    def eventFilter(self, obj, event):
        # We install this on the viewport so positions are in viewport coords
        if self._viewport is not None and obj is self._viewport:
            if event.type() == QEvent.MouseMove:
                try:
                    href = self._edit.anchorAt(event.pos()) if self._edit is not None else ""
                    self._viewport.setCursor(Qt.PointingHandCursor if href else Qt.IBeamCursor)
                except RuntimeError:
                    return False
            elif event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                try:
                    self._pressed_anchor = (
                        self._edit.anchorAt(event.pos()) if self._edit is not None else None
                    )
                except RuntimeError:
                    self._pressed_anchor = None
            elif event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                try:
                    href = self._edit.anchorAt(event.pos()) if self._edit is not None else None
                except RuntimeError:
                    href = None
                if href and href == self._pressed_anchor:
                    try:
                        # Resolve relative paths (like media/...) against document base for external open
                        if href and not re.match(r"^[a-zA-Z]+:|^/", href):
                            base = self._edit.document().baseUrl().toLocalFile() if self._edit else ""
                            if base:
                                # Join using OS path, then convert to file URL
                                abs_path = os.path.normpath(os.path.join(base, href))
                                QDesktopServices.openUrl(QUrl.fromLocalFile(abs_path))
                                return True
                        QDesktopServices.openUrl(QUrl(_normalize_url_scheme(href)))
                    except Exception:
                        pass
                    # Prevent the click from also moving the caret
                    return True
                self._pressed_anchor = None
        return super().eventFilter(obj, event)


def _install_link_click_handler(text_edit: QtWidgets.QTextEdit):
    # Allow links to be interactable with the mouse while still editing
    try:
        flags = text_edit.textInteractionFlags()
        flags |= Qt.LinksAccessibleByMouse
        text_edit.setTextInteractionFlags(flags)
    except Exception:
        pass
    try:
        if text_edit is None:
            return
        handler = _LinkClickHandler(text_edit)
        vp = text_edit.viewport() if text_edit is not None else None
        if vp is not None:
            vp.installEventFilter(handler)
        # Keep strong reference on the text_edit object
        if not hasattr(text_edit, "_linkHandler"):
            text_edit._linkHandler = []
        text_edit._linkHandler.append(handler)
    except RuntimeError:
        # Widget likely already destroyed; ignore
        return
