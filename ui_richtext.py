"""
ui_richtext.py
Adds a lightweight rich text toolbar to a QTextEdit inside a tab.
Supports: Undo/Redo, Bold/Italic/Underline/Strike, Font family/size,
Text color/Highlight, Align L/C/R/Justify, Bullets/Numbers, Clear formatting,
Insert image, Horizontal rule.
"""
from PyQt5 import QtWidgets
from PyQt5.QtGui import QTextCharFormat, QFont, QColor, QKeySequence, QIcon, QPixmap, QPainter, QPen, QTextListFormat, QTextCursor, QTextList, QDesktopServices, QPalette
from PyQt5.QtCore import Qt, QSize, QRect, QPoint, QEvent, QObject, QUrl


def _ensure_layout(widget: QtWidgets.QWidget) -> QtWidgets.QVBoxLayout:
    layout = widget.layout()
    if isinstance(layout, QtWidgets.QVBoxLayout):
        return layout
    layout = QtWidgets.QVBoxLayout(widget)
    widget.setLayout(layout)
    return layout

# Defaults you can change
DEFAULT_FONT_FAMILY = "Arial"  # e.g., "Arial", "Calibri", "Times New Roman"
DEFAULT_FONT_SIZE_PT = 12          # in points

# List scheme configuration (can be changed at runtime from main menu)
_ORDERED_SCHEME = 'classic'  # 'classic' or 'decimal'
_UNORDERED_SCHEME = 'disc-circle-square'  # 'disc-circle-square' or 'disc-only'


def _make_icon(kind: str, size: QSize = QSize(24, 24), fg: QColor = QColor('#303030')) -> QIcon:
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
            p.drawLine(int(w*0.25), y, int(w*0.75), y)
        if strike:
            y = int(h * 0.5)
            p.drawLine(int(w*0.25), y, int(w*0.75), y)

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
            if mode == 'left':
                x1, x2 = 3, 3 + int((w-6) * lw)
            elif mode == 'center':
                span = int((w-6) * lw)
                x1 = (w - span) // 2
                x2 = x1 + span
            elif mode == 'right':
                span = int((w-6) * lw)
                x2 = w-3
                x1 = x2 - span
            else:  # justify
                x1, x2 = 3, w-3
            p.drawLine(x1, yy, x2, yy)

    def _draw_list(bullets: bool):
        y = 5
        for i in range(3):
            yy = y + i * int(h * 0.24)
            if bullets:
                p.drawEllipse(QPoint(6, yy), 2, 2)
            else:
                # tiny '1.' like marker
                p.drawText(QRect(2, yy-6, 10, 12), Qt.AlignLeft | Qt.AlignVCenter, str(i+1)+'.')
            p.drawLine(12, yy, w-4, yy)

    if kind == 'undo':
        # left arrow
        p.drawLine(int(w*0.75), int(h*0.3), int(w*0.35), int(h*0.3))
        p.drawLine(int(w*0.35), int(h*0.3), int(w*0.45), int(h*0.2))
        p.drawLine(int(w*0.35), int(h*0.3), int(w*0.45), int(h*0.4))
    elif kind == 'redo':
        p.drawLine(int(w*0.25), int(h*0.3), int(w*0.65), int(h*0.3))
        p.drawLine(int(w*0.65), int(h*0.3), int(w*0.55), int(h*0.2))
        p.drawLine(int(w*0.65), int(h*0.3), int(w*0.55), int(h*0.4))
    elif kind == 'bold':
        _draw_text('B', bold=True)
    elif kind == 'italic':
        _draw_text('I', italic=True)
    elif kind == 'underline':
        _draw_text('U', underline=True)
    elif kind == 'strike':
        _draw_text('S', strike=True)
    elif kind == 'align_left':
        _draw_align('left')
    elif kind == 'align_center':
        _draw_align('center')
    elif kind == 'align_right':
        _draw_align('right')
    elif kind == 'align_justify':
        _draw_align('justify')
    elif kind == 'list_bullets':
        _draw_list(True)
    elif kind == 'list_numbers':
        _draw_list(False)
    elif kind == 'indent':
        # right pointing arrow/step
        p.drawLine(4, h//2, w-8, h//2)
        p.drawLine(w-10, h//2 - 6, w-4, h//2)
        p.drawLine(w-10, h//2 + 6, w-4, h//2)
        p.drawLine(4, 6, 4, h-6)
    elif kind == 'outdent':
        # left pointing arrow/step
        p.drawLine(w-4, h//2, 8, h//2)
        p.drawLine(10, h//2 - 6, 4, h//2)
        p.drawLine(10, h//2 + 6, 4, h//2)
        p.drawLine(w-4, 6, w-4, h-6)
    elif kind == 'hr':
        y = h//2
        p.drawLine(4, y, w-4, y)
    elif kind == 'image':
        # simple picture: frame + mountain + sun
        p.drawRect(3, 5, w-6, h-10)
        p.drawLine(6, h-8, w-10, h-14)
        p.drawEllipse(QPoint(w-10, 9), 2, 2)
    elif kind == 'table':
        # 3x3 grid
        p.drawRect(3, 5, w-6, h-10)
        for i in range(1, 3):
            x = 3 + i * (w-6)//3
            p.drawLine(x, 5, x, h-5)
            y = 5 + i * (h-10)//3
            p.drawLine(3, y, w-3, y)
    elif kind == 'color':
        _draw_text('A')
        p.drawRect(5, h-8, w-10, 4)
    elif kind == 'highlight':
        p.fillRect(QRect(4, h-10, w-8, 6), QColor('#ffe680'))
        _draw_text('A')

    p.end()
    return QIcon(pm)


def add_rich_text_toolbar(parent_tab: QtWidgets.QWidget, text_edit: QtWidgets.QTextEdit, before_widget: QtWidgets.QWidget = None):
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
    act_undo = toolbar.addAction(_make_icon('undo'), "", text_edit.undo)
    act_undo.setShortcut(QKeySequence.Undo)
    act_undo.setToolTip("Undo (Ctrl+Z)")
    act_redo = toolbar.addAction(_make_icon('redo'), "", text_edit.redo)
    act_redo.setShortcut(QKeySequence.Redo)
    act_redo.setToolTip("Redo (Ctrl+Y)")
    toolbar.addSeparator()

    # Bold/Italic/Underline/Strike
    def toggle_format(flag_attr: str, on: bool):
        fmt = QTextCharFormat()
        if flag_attr == 'bold':
            fmt.setFontWeight(QFont.Bold if on else QFont.Normal)
        elif flag_attr == 'italic':
            fmt.setFontItalic(on)
        elif flag_attr == 'underline':
            fmt.setFontUnderline(on)
        elif flag_attr == 'strike':
            fmt.setFontStrikeOut(on)
        cursor = text_edit.textCursor()
        if not cursor.hasSelection():
            # Apply to current word/cursor moving forward
            cursor.select(cursor.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        text_edit.mergeCurrentCharFormat(fmt)

    act_bold = QtWidgets.QAction(_make_icon('bold'), "", toolbar)
    act_bold.setCheckable(True)
    act_bold.setShortcut(QKeySequence.Bold)
    act_bold.setToolTip("Bold (Ctrl+B)")
    act_bold.triggered.connect(lambda on: toggle_format('bold', on))
    toolbar.addAction(act_bold)

    act_italic = QtWidgets.QAction(_make_icon('italic'), "", toolbar)
    act_italic.setCheckable(True)
    act_italic.setShortcut(QKeySequence.Italic)
    act_italic.setToolTip("Italic (Ctrl+I)")
    act_italic.triggered.connect(lambda on: toggle_format('italic', on))
    toolbar.addAction(act_italic)

    act_underline = QtWidgets.QAction(_make_icon('underline'), "", toolbar)
    act_underline.setCheckable(True)
    act_underline.setShortcut(QKeySequence.Underline)
    act_underline.setToolTip("Underline (Ctrl+U)")
    act_underline.triggered.connect(lambda on: toggle_format('underline', on))
    toolbar.addAction(act_underline)

    act_strike = QtWidgets.QAction(_make_icon('strike'), "", toolbar)
    act_strike.setCheckable(True)
    act_strike.setToolTip("Strikethrough")
    act_strike.triggered.connect(lambda on: toggle_format('strike', on))
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
    size_box.currentIndexChanged.connect(lambda _i: _apply_font_size(text_edit, size_box.currentData()))
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

    toolbar.addSeparator()

    # Alignment
    group_align = QtWidgets.QActionGroup(toolbar)
    group_align.setExclusive(True)
    act_align_left = QtWidgets.QAction(_make_icon('align_left'), "", toolbar, checkable=True)
    act_align_left.setToolTip("Align Left")
    act_align_center = QtWidgets.QAction(_make_icon('align_center'), "", toolbar, checkable=True)
    act_align_center.setToolTip("Align Center")
    act_align_right = QtWidgets.QAction(_make_icon('align_right'), "", toolbar, checkable=True)
    act_align_right.setToolTip("Align Right")
    act_align_justify = QtWidgets.QAction(_make_icon('align_justify'), "", toolbar, checkable=True)
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
    act_bullets = toolbar.addAction(_make_icon('list_bullets'), "", lambda: _toggle_list(text_edit, ordered=False))
    act_bullets.setToolTip("Bulleted list")
    act_numbers = toolbar.addAction(_make_icon('list_numbers'), "", lambda: _toggle_list(text_edit, ordered=True))
    act_numbers.setToolTip("Numbered list")

    # Indent/Outdent for list nesting
    act_indent = toolbar.addAction(_make_icon('indent'), "", lambda: _change_list_indent(text_edit, +1))
    act_indent.setShortcut(QKeySequence("Ctrl+]"))
    act_indent.setToolTip("Indent (Tab, Ctrl+])")
    act_outdent = toolbar.addAction(_make_icon('outdent'), "", lambda: _change_list_indent(text_edit, -1))
    act_outdent.setShortcut(QKeySequence("Ctrl+["))
    act_outdent.setToolTip("Outdent (Shift+Tab, Ctrl+[)")

    # Enable Tab/Shift+Tab to control list levels
    _install_list_tab_handler(text_edit)

    toolbar.addSeparator()

    # Clear formatting, HR, Insert image
    act_clear = toolbar.addAction(_make_icon('color'), "", lambda: text_edit.setCurrentCharFormat(QTextCharFormat()))
    act_clear.setToolTip("Clear formatting")
    act_hr = toolbar.addAction(_make_icon('hr'), "", lambda: text_edit.textCursor().insertHtml("<hr/>"))
    act_hr.setToolTip("Insert horizontal rule")
    act_img = toolbar.addAction(_make_icon('image'), "", lambda: _insert_image_via_dialog(text_edit))
    act_img.setToolTip("Insert image from file")

    # Paste Text Only quick action
    act_paste_plain = toolbar.addAction(_make_icon('color'), "", lambda: paste_text_only(text_edit))
    act_paste_plain.setToolTip("Paste Text Only (Ctrl+Shift+V)")

    # Placeholder: Table (to be implemented with insert/edit actions)
    toolbar.addSeparator()
    btn_table = QtWidgets.QToolButton(toolbar)
    btn_table.setIcon(_make_icon('table'))
    btn_table.setToolTip("Insert/edit table (coming soon)")
    btn_table.setEnabled(False)
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

    # Improve selection visibility: use a clearer highlight and text color
    try:
        _apply_selection_colors(text_edit, QColor("#4d84b7"), QColor('#000000'))
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
        text = md.data('text/plain').data().decode('utf-8', errors='replace') if md.hasFormat('text/plain') else ''
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
            html = txt.replace('\n', '<br/>')
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
    s = re.sub(r'<\s*style\b[^>]*>.*?<\s*/\s*style\s*>', '', s, flags=re.IGNORECASE | re.DOTALL)
    # Remove bgcolor attribute
    s = re.sub(r'\sbgcolor\s*=\s*"[^"]*"', '', s, flags=re.IGNORECASE)
    s = re.sub(r"\sbgcolor\s*=\s*'[^']*'", '', s, flags=re.IGNORECASE)
    s = re.sub(r'\sbgcolor\s*=\s*[^\s>]+', '', s, flags=re.IGNORECASE)
    # Replace deprecated <font> tags with span
    s = re.sub(r'<\s*font\b[^>]*>', '<span>', s, flags=re.IGNORECASE)
    s = re.sub(r'<\s*/\s*font\s*>', '</span>', s, flags=re.IGNORECASE)
    # Drop face/size/color attributes
    s = re.sub(r'\s(face|size|color)\s*=\s*"[^"]*"', '', s, flags=re.IGNORECASE)
    s = re.sub(r"\s(face|size|color)\s*=\s*'[^']*'", '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s(face|size|color)\s*=\s*[^\s>]+', '', s, flags=re.IGNORECASE)
    # Clean style attributes: remove background*, font-size, font-family, shorthand font, and line-height
    def _clean_style(m):
        inner = m.group(1)
        parts = [p.strip() for p in inner.split(';') if p.strip()]
        kept = []
        for p in parts:
            key = p.split(':',1)[0].strip().lower()
            if key.startswith('background') or key in ('font-size','font-family','font','line-height'):
                continue
            kept.append(p)
        if not kept:
            return ''
        return ' style="' + '; '.join(kept) + '"'
    s = re.sub(r'\sstyle\s*=\s*"([^"]*)"', _clean_style, s, flags=re.IGNORECASE)
    s = re.sub(r"\sstyle\s*=\s*'([^']*)'", lambda m: _clean_style(m).replace('"','\"'), s, flags=re.IGNORECASE)
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
            if tag_l == 'style':
                # Skip entire style blocks
                self._skip_style = True
                return
            # Convert deprecated <font> to span
            if tag_l == 'font':
                tag_l = 'span'
            allowed = []
            buffered_style = None
            for k, v in attrs:
                lk = k.lower()
                if lk in ('class', 'bgcolor', 'color', 'face', 'size'):
                    continue
                if lk == 'style':
                    # Keep only Qt list-related declarations to preserve indent/numbering across reloads
                    # Applies to list containers and list paragraphs
                    if tag_l in ('ol','ul','li','p'):
                        try:
                            decls = [d.strip() for d in str(v).split(';') if d.strip()]
                            kept = []
                            for d in decls:
                                key = d.split(':',1)[0].strip().lower()
                                if key.startswith('-qt-list-') or key == '-qt-paragraph-type':
                                    kept.append(d)
                            if kept:
                                buffered_style = '; '.join(kept)
                        except Exception:
                            pass
                    continue
                # Preserve list semantics
                if tag_l in ('ol','ul') and lk in ('type','start'):  # type may be set by Qt for list appearance
                    allowed.append((k, v))
                elif tag_l == 'li' and lk in ('value',):  # value allows continuing numbering
                    allowed.append((k, v))
                elif tag_l == 'a' and lk in ('href', 'title'):
                    allowed.append((k, v))
                elif tag_l == 'img' and lk in ('src', 'alt', 'title'):
                    allowed.append((k, v))
                elif lk.startswith('data-'):
                    continue
                elif lk in ('width','height','cellpadding','cellspacing','border') and tag_l in ('table','td','th','tr'):
                    allowed.append((k, v))
                # drop everything else
            if buffered_style:
                allowed.append(('style', buffered_style))
            attrs_txt = ''.join(f' {k}="{v}"' for k,v in allowed)
            self.out.append(f'<{tag_l}{attrs_txt}>' )

        def handle_endtag(self, tag):
            tag_l = tag.lower()
            if tag_l == 'style':
                self._skip_style = False
                return
            if tag_l == 'font':
                tag_l = 'span'
            self.out.append(f'</{tag_l}>' )

        def handle_startendtag(self, tag, attrs):
            tag_l = tag.lower()
            if tag_l == 'style':
                return
            allowed = []
            buffered_style = None
            for k, v in attrs:
                lk = k.lower()
                if lk in ('class', 'bgcolor', 'color', 'face', 'size'):
                    continue
                if lk == 'style':
                    tag_l = tag_l
                    if tag_l in ('ol','ul','li','p'):
                        try:
                            decls = [d.strip() for d in str(v).split(';') if d.strip()]
                            kept = []
                            for d in decls:
                                key = d.split(':',1)[0].strip().lower()
                                if key.startswith('-qt-list-') or key == '-qt-paragraph-type':
                                    kept.append(d)
                            if kept:
                                buffered_style = '; '.join(kept)
                        except Exception:
                            pass
                    continue
                if tag_l == 'img' and lk in ('src', 'alt', 'title'):
                    allowed.append((k, v))
            if buffered_style:
                allowed.append(('style', buffered_style))
            attrs_txt = ''.join(f' {k}="{v}"' for k,v in allowed)
            self.out.append(f'<{tag_l}{attrs_txt}/>' )

        def handle_data(self, data):
            if not self._skip_style:
                self.out.append(data)

    try:
        cl = _StoreCleaner()
        cl.feed(raw_html)
        return ''.join(cl.out)
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
    if _ORDERED_SCHEME == 'decimal':
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
    if _UNORDERED_SCHEME == 'disc-only':
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
    if ordered in (None, ''):
        pass
    elif ordered in ('classic', 'decimal'):
        _ORDERED_SCHEME = ordered
    if unordered in (None, ''):
        pass
    elif unordered in ('disc-circle-square', 'disc-only'):
        _UNORDERED_SCHEME = unordered


def get_list_schemes():
    return _ORDERED_SCHEME, _UNORDERED_SCHEME


def _insert_image_via_dialog(text_edit: QtWidgets.QTextEdit):
    path, _ = QtWidgets.QFileDialog.getOpenFileName(text_edit, "Insert Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp);;All Files (*)")
    if not path:
        return
    cursor = text_edit.textCursor()
    cursor.insertImage(path)


def _apply_selection_colors(text_edit: QtWidgets.QTextEdit, bg: QColor, fg: QColor):
    pal = text_edit.palette()
    for group in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
        pal.setColor(group, QPalette.Highlight, bg)
        pal.setColor(group, QPalette.HighlightedText, fg)
    text_edit.setPalette(pal)
    # Stylesheet fallback (some styles prefer stylesheet roles)
    try:
        text_edit.setStyleSheet(
            "QTextEdit { selection-background-color: %s; selection-color: %s; }" % (bg.name(), fg.name())
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
    lf.setStyle(_ordered_style_for_level(new_level) if ordered else _unordered_style_for_level(new_level))
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
                    is_backtab = (key == Qt.Key_Backtab) or bool(event.modifiers() & Qt.ShiftModifier)
                    _change_list_indent(self._edit, -1 if is_backtab else +1)
                    return True
        return super().eventFilter(obj, event)


def _install_list_tab_handler(text_edit: QtWidgets.QTextEdit):
    handler = _ListTabHandler(text_edit)
    text_edit.installEventFilter(handler)
    # Keep a reference to prevent GC
    if not hasattr(text_edit, "_listTabHandler"):
        text_edit._listTabHandler = []
    text_edit._listTabHandler.append(handler)


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
                return get_default_paste_mode() or 'rich'
            except Exception:
                return 'rich'

        def eventFilter(self, obj, event):
            if obj is self._edit and event.type() == QEvent.KeyPress:
                if event.key() == Qt.Key_V and (event.modifiers() & Qt.ControlModifier) and not (event.modifiers() & Qt.AltModifier):
                    mode = self._mode()
                    if mode == 'text-only':
                        paste_text_only(self._edit)
                    elif mode == 'match-style':
                        paste_match_style(self._edit)
                    elif mode == 'clean':
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
            html = txt.replace('\n', '<br/>')
    if not html:
        return
    try:
        s = _strip_match_style_html(html)
        # Additionally drop any remaining style/class attributes outright
        import re
        s = re.sub(r'\sstyle\s*=\s*"[^"]*"', '', s, flags=re.IGNORECASE)
        s = re.sub(r"\sstyle\s*=\s*'[^']*'", '', s, flags=re.IGNORECASE)
        s = re.sub(r'\sclass\s*=\s*"[^"]*"', '', s, flags=re.IGNORECASE)
        s = re.sub(r"\sclass\s*=\s*'[^']*'", '', s, flags=re.IGNORECASE)
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
    if t.lower().startswith(('http://', 'https://', 'mailto:')):
        return True
    if t.lower().startswith('www.') and ' ' not in t:
        return True
    # basic domain.tld pattern without spaces
    return ('.' in t and ' ' not in t and '/' in t) or ('.' in t and t.count('.') >= 1 and ' ' not in t)


def _normalize_url_scheme(text: str) -> str:
    t = (text or '').strip()
    if not t:
        return t
    if t.lower().startswith(('http://', 'https://', 'mailto:')):
        return t
    if t.lower().startswith('www.'):
        return 'http://' + t
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
                    href = self._edit.anchorAt(event.pos()) if self._edit is not None else ''
                    self._viewport.setCursor(Qt.PointingHandCursor if href else Qt.IBeamCursor)
                except RuntimeError:
                    return False
            elif event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                try:
                    self._pressed_anchor = self._edit.anchorAt(event.pos()) if self._edit is not None else None
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
        if not hasattr(text_edit, '_linkHandler'):
            text_edit._linkHandler = []
        text_edit._linkHandler.append(handler)
    except RuntimeError:
        # Widget likely already destroyed; ignore
        return
