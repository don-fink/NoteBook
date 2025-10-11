"""
ui_richtext.py
Adds a lightweight rich text toolbar to a QTextEdit inside a tab.
Supports: Undo/Redo, Bold/Italic/Underline/Strike, Font family/size,
Text color/Highlight, Align L/C/R/Justify, Bullets/Numbers, Clear formatting,
Insert image, Horizontal rule.
"""

from PyQt5 import QtWidgets
import re
from PyQt5.QtCore import QEvent, QObject, QPoint, QRect, QSize, Qt, QUrl
from PyQt5.QtGui import (
    QColor,
    QDesktopServices,
    QFont,
    QIcon,
    QKeySequence,
    QPainter,
    QPalette,
    QPen,
    QPixmap,
    QTextCharFormat,
    QTextCursor,
    QTextList,
    QTextListFormat,
    QTextLength,
    QTextTableFormat,
)


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
    # Visual polish: icon-only toolbar, larger icons
    toolbar.setIconSize(QSize(24, 24))
    toolbar.setStyleSheet(
        """
        QToolBar {
            background: #f6f6f6;
            border-bottom: 1px solid #d0d0d0;
            spacing: 4px;
        }
        QToolButton {
            padding: 2px 6px;
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
        font_box.setMaximumWidth(200)
        font_box.setMinimumContentsLength(10)
    except Exception:
        pass
    font_box.currentFontChanged.connect(lambda f: _apply_font_family(text_edit, f.family()))
    font_box.setToolTip("Font family")
    toolbar.addWidget(font_box)

    size_box = QtWidgets.QComboBox(toolbar)
    for sz in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32]:
        size_box.addItem(str(sz), sz)
    size_box.setEditable(False)
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
    btn_color.clicked.connect(lambda: _pick_color_and_apply(text_edit, foreground=True))
    toolbar.addWidget(btn_color)

    btn_bg = QtWidgets.QToolButton(toolbar)
    btn_bg.setText("Bg")
    btn_bg.setToolTip("Highlight")
    btn_bg.clicked.connect(lambda: _pick_color_and_apply(text_edit, foreground=False))
    toolbar.addWidget(btn_bg)

    # Clear only background highlight (keep bold/italic/etc.)
    btn_bg_clear = QtWidgets.QToolButton(toolbar)
    btn_bg_clear.setText("NoBg")
    btn_bg_clear.setToolTip("Remove highlight (background)")
    btn_bg_clear.clicked.connect(lambda: _clear_background(text_edit))
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
    act_indent.setToolTip("Indent (Tab, Ctrl+])")
    act_outdent = toolbar.addAction(
        _make_icon("outdent"), "", lambda: _change_list_indent(text_edit, -1)
    )
    act_outdent.setShortcut(QKeySequence("Ctrl+["))
    act_outdent.setToolTip("Outdent (Shift+Tab, Ctrl+[)")

    # Enable Tab/Shift+Tab to control list levels
    _install_list_tab_handler(text_edit)
    # Enable table cell Tab navigation
    _install_table_tab_handler(text_edit)
    # Enable plain paragraph indent/outdent with Tab/Shift+Tab when not in lists/tables
    _install_plain_indent_tab_handler(text_edit)

    toolbar.addSeparator()

    # Clear formatting, HR, Insert image
    act_clear = toolbar.addAction(
        _make_icon("color"), "", lambda: text_edit.setCurrentCharFormat(QTextCharFormat())
    )
    act_clear.setToolTip("Clear formatting")
    act_hr = toolbar.addAction(
        _make_icon("hr"), "", lambda: text_edit.textCursor().insertHtml("<hr/>")
    )
    act_hr.setToolTip("Insert horizontal rule")
    act_img = toolbar.addAction(
        _make_icon("image"), "", lambda: _insert_image_via_dialog(text_edit)
    )
    act_img.setToolTip("Insert image from file")

    # Paste Text Only quick action
    act_paste_plain = toolbar.addAction(_make_icon("color"), "", lambda: paste_text_only(text_edit))
    act_paste_plain.setToolTip("Paste Text Only (Ctrl+Shift+V)")

    # Table: insert or edit if caret is inside a table
    toolbar.addSeparator()
    btn_table = QtWidgets.QToolButton(toolbar)
    btn_table.setIcon(_make_icon("table"))
    btn_table.setToolTip("Insert/edit table")
    btn_table.setEnabled(True)
    btn_table.clicked.connect(lambda: _table_insert_or_edit(text_edit))
    toolbar.addWidget(btn_table)

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

    # Install Ctrl+V override to honor Default Paste Mode without relying on a window-level shortcut
    _install_default_paste_override(text_edit)
    # Enable Ctrl+Click to open links in the system browser
    _install_link_click_handler(text_edit)
    # Enable right-click table context menu
    _install_table_context_menu(text_edit)
    # Ensure default context menu events propagate so our filter can handle them
    try:
        text_edit.setContextMenuPolicy(Qt.DefaultContextMenu)
    except Exception:
        pass

    # Improve selection visibility: use a clearer highlight and text color
    try:
        _apply_selection_colors(text_edit, QColor("#4d84b7"), QColor("#000000"))
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
                if lk in ("class", "bgcolor", "color", "face", "size"):
                    continue
                if lk == "style":
                    # Keep only Qt list-related declarations to preserve indent/numbering across reloads
                    # Also keep paragraph left margin to support Tab/Shift+Tab indent for plain text
                    if tag_l in ("ol", "ul", "li", "p"):
                        try:
                            decls = [d.strip() for d in str(v).split(";") if d.strip()]
                            kept = []
                            for d in decls:
                                key = d.split(":", 1)[0].strip().lower()
                                if (
                                    key.startswith("-qt-list-")
                                    or key == "-qt-paragraph-type"
                                    or (tag_l == "p" and key in ("margin-left",))
                                ):
                                    kept.append(d)
                            if kept:
                                buffered_style = "; ".join(kept)
                        except Exception:
                            pass
                    continue
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
                elif tag_l == "img" and lk in ("src", "alt", "title"):
                    allowed.append((k, v))
                elif lk.startswith("data-"):
                    continue
                elif lk in (
                    "width",
                    "height",
                    "cellpadding",
                    "cellspacing",
                    "border",
                ) and tag_l in ("table", "td", "th", "tr"):
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
                if lk in ("class", "bgcolor", "color", "face", "size"):
                    continue
                if lk == "style":
                    if tag_l in ("ol", "ul", "li", "p"):
                        try:
                            decls = [d.strip() for d in str(v).split(";") if d.strip()]
                            kept = []
                            for d in decls:
                                key = d.split(":", 1)[0].strip().lower()
                                if (
                                    key.startswith("-qt-list-")
                                    or key == "-qt-paragraph-type"
                                    or (tag_l == "p" and key in ("margin-left",))
                                ):
                                    kept.append(d)
                            if kept:
                                buffered_style = "; ".join(kept)
                        except Exception:
                            pass
                    continue
                if tag_l == "img" and lk in ("src", "alt", "title"):
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


def _insert_image_via_dialog(text_edit: QtWidgets.QTextEdit):
    path, _ = QtWidgets.QFileDialog.getOpenFileName(
        text_edit, "Insert Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp);;All Files (*)"
    )
    if not path:
        return
    cursor = text_edit.textCursor()
    cursor.insertImage(path)


# ----------------------------- Tables -----------------------------
def _current_table(text_edit: QtWidgets.QTextEdit):
    try:
        cur = text_edit.textCursor()
        return cur.currentTable()
    except Exception:
        return None


def insert_table_from_preset(text_edit: QtWidgets.QTextEdit, preset_name: str, fit_width_100: bool = True):
    """Insert a table defined by a saved preset at the current cursor position.

    fit_width_100: If True, force table width to 100%% regardless of the preset's saved width.
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
    rows = max(1, int(data.get("rows", 3)))
    cols = max(1, int(data.get("columns", 3)))
    fmt = QTextTableFormat()
    try:
        fmt.setBorder(float(data.get("border", 1.0)))
        fmt.setCellPadding(float(data.get("cell_padding", 2.0)))
        fmt.setCellSpacing(float(data.get("cell_spacing", 0.0)))
    except Exception:
        pass
    width_pct = float(data.get("width_pct", 100.0))
    try:
        fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0 if fit_width_100 else width_pct))
    except Exception:
        pass
    # Column widths (percentages)
    col_widths = data.get("column_widths_pct") or []
    try:
        if isinstance(col_widths, (list, tuple)) and len(col_widths) == cols:
            fmt.setColumnWidthConstraints(
                [QTextLength(QTextLength.PercentageLength, float(p)) for p in col_widths]
            )
        else:
            frac = 100.0 / float(cols)
            fmt.setColumnWidthConstraints(
                [QTextLength(QTextLength.PercentageLength, frac) for _ in range(cols)]
            )
    except Exception:
        pass
    # Header row count if provided
    try:
        hrc = int(data.get("header_row_count", 0))
        if hrc > 0:
            fmt.setHeaderRowCount(hrc)
    except Exception:
        pass
    cur = text_edit.textCursor()
    table = cur.insertTable(rows, cols, fmt)
    # Populate headers if available
    headers = data.get("headers") or []
    if isinstance(headers, (list, tuple)) and headers:
        row0 = 0
        # Apply bold + light gray background to the header row
        header_fmt = QTextCharFormat()
        header_fmt.setFontWeight(QFont.Bold)
        try:
            bg = QColor(245, 245, 245)
        except Exception:
            bg = None
        for c in range(min(cols, len(headers))):
            cell = table.cellAt(row0, c)
            if bg is not None:
                cf = cell.format()
                cf.setBackground(bg)
                cell.setFormat(cf)
            tcur = cell.firstCursorPosition()
            tcur.mergeCharFormat(header_fmt)
            try:
                tcur.insertText(str(headers[c]))
            except Exception:
                pass
    return table


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
            # Capture original selection and table
            orig_cur = self._edit.textCursor()
            orig_tbl = orig_cur.currentTable()
            # Compute selection rectangle from original selection (do not disturb selection)
            orig_rect = _table_selection_rect(self._edit, orig_tbl)
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
                # Insert Preset submenu
                sub_preset = menu.addMenu("Insert Preset")
                try:
                    from settings_manager import list_table_preset_names

                    names = list_table_preset_names()
                except Exception:
                    names = []
                _preset_actions = {}
                if names:
                    for nm in names:
                        a = sub_preset.addAction(nm)
                        _preset_actions[a] = nm
                else:
                    sub_preset.setEnabled(False)
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
                if chosen in _preset_actions:
                    # Insert selected preset at cursor, force 100% width per user preference
                    insert_table_from_preset(self._edit, _preset_actions[chosen], fit_width_100=True)
                    return True
                return True
            # Otherwise, build the full table menu
            menu = QtWidgets.QMenu(self._edit)
            act_ins = menu.addAction("Insert Table…")
            act_prop = menu.addAction("Table Properties…")
            act_fit = menu.addAction("Fit Table to Width")
            act_dist = menu.addAction("Distribute Columns Evenly")
            act_set_col = menu.addAction("Set Current Column Width…")
            menu.addSeparator()
            act_save_preset = menu.addAction("Save Table as Preset…")
            # Insert Preset submenu while inside a table
            sub_insert_preset = menu.addMenu("Insert Preset")
            try:
                from settings_manager import list_table_preset_names

                names = list_table_preset_names()
            except Exception:
                names = []
            _ins_preset_actions = {}
            if names:
                for nm in names:
                    a = sub_insert_preset.addAction(nm)
                    _ins_preset_actions[a] = nm
            else:
                sub_insert_preset.setEnabled(False)
            menu.addSeparator()
            act_recalc = menu.addAction("Recalculate Formulas (SUM)")
            menu.addSeparator()
            # Determine multi-cell selection rectangle (within chosen table)
            sel_rect = _table_selection_rect(self._edit, tbl)
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
                try:
                    name, ok = QtWidgets.QInputDialog.getText(
                        self._edit, "Save Table Preset", "Preset name:", text="My Table"
                    )
                except Exception:
                    ok = False
                    name = None
                if ok and name and name.strip():
                    # Capture table structure and headers
                    fmt = tbl.format()
                    try:
                        wlen = fmt.width()
                        width_pct = (
                            (wlen.rawValue() if hasattr(wlen, "rawValue") else wlen.value())
                            if (hasattr(wlen, "type") and wlen.type() == wlen.PercentageLength)
                            else 100.0
                        )
                    except Exception:
                        width_pct = 100.0
                    try:
                        constraints = list(fmt.columnWidthConstraints()) or []
                        col_widths = [
                            (c.rawValue() if hasattr(c, "rawValue") else c.value()) for c in constraints
                        ] if constraints else []
                    except Exception:
                        col_widths = []
                    rows = tbl.rows()
                    cols = tbl.columns()
                    headers = []
                    try:
                        for c in range(cols):
                            headers.append(_table_cell_plain_text(tbl, 0, c).strip())
                    except Exception:
                        headers = []
                    preset = {
                        "columns": int(cols),
                        "rows": int(rows),
                        "width_pct": float(width_pct),
                        "border": float(fmt.border()),
                        "cell_padding": float(fmt.cellPadding()),
                        "cell_spacing": float(fmt.cellSpacing()),
                        "column_widths_pct": col_widths,
                        "header_row_count": int(fmt.headerRowCount() if hasattr(fmt, "headerRowCount") else 0),
                        "headers": headers,
                    }
                    try:
                        from settings_manager import save_table_preset

                        save_table_preset(name.strip(), preset)
                        QtWidgets.QToolTip.showText(global_pos, f"Saved preset '{name.strip()}'")
                    except Exception:
                        pass
            elif chosen in _ins_preset_actions:
                insert_table_from_preset(self._edit, _ins_preset_actions[chosen], fit_width_100=True)
            elif chosen == act_recalc and has_tbl:
                try:
                    _table_recalculate_formulas(self._edit, tbl)
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
    try:
        table.insertRows(base_row, count)
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


_SUM_RE = re.compile(r"^\s*=\s*SUM\(\s*([A-Za-z]+\d+)\s*:\s*([A-Za-z]+\d+)\s*\)\s*$")


def _table_recalculate_formulas(text_edit: QtWidgets.QTextEdit, table):
    """Scan table for cells that contain '=SUM(A1:B10)' (case-insensitive) and replace with the computed sum.
    Note: formulas are not persisted to storage; this is an in-session convenience feature.
    """
    try:
        rows = table.rows()
        cols = table.columns()
        for r in range(rows):
            for c in range(cols):
                raw = _table_cell_plain_text(table, r, c)
                if not raw:
                    continue
                m = _SUM_RE.match(raw)
                if not m:
                    continue
                start_addr, end_addr = m.group(1), m.group(2)
                val = _sum_range_in_table(table, start_addr, end_addr)
                # Format without trailing .0 for integers
                out = ("%d" % int(val)) if abs(val - int(val)) < 1e-9 else ("%s" % val)
                _table_set_cell_plain_text(text_edit, table, r, c, out)
    except Exception:
        pass


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
    fmt.setFontFamily(family)
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


def _pick_color_and_apply(text_edit: QtWidgets.QTextEdit, foreground: bool = True):
    dlg = QtWidgets.QColorDialog(text_edit)
    dlg.setOption(QtWidgets.QColorDialog.ShowAlphaChannel, False)
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
