"""
ui_richtext.py
Adds a lightweight rich text toolbar to a QTextEdit inside a tab.
Supports: Undo/Redo, Bold/Italic/Underline/Strike, Font family/size,
Text color/Highlight, Align L/C/R/Justify, Bullets/Numbers, Clear formatting,
Insert image, Horizontal rule.
"""
from PyQt5 import QtWidgets
from PyQt5.QtGui import QTextCharFormat, QFont, QColor
from PyQt5.QtCore import Qt


def _ensure_layout(widget: QtWidgets.QWidget) -> QtWidgets.QVBoxLayout:
    layout = widget.layout()
    if isinstance(layout, QtWidgets.QVBoxLayout):
        return layout
    layout = QtWidgets.QVBoxLayout(widget)
    widget.setLayout(layout)
    return layout


def add_rich_text_toolbar(parent_tab: QtWidgets.QWidget, text_edit: QtWidgets.QTextEdit, before_widget: QtWidgets.QWidget = None):
    if text_edit is None or parent_tab is None:
        return None
    layout = _ensure_layout(parent_tab)
    toolbar = QtWidgets.QToolBar(parent_tab)
    toolbar.setIconSize(QtWidgets.QSize(16, 16)) if hasattr(QtWidgets, 'QSize') else None

    # Undo/Redo
    act_undo = toolbar.addAction("Undo", text_edit.undo)
    act_redo = toolbar.addAction("Redo", text_edit.redo)
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

    act_bold = QtWidgets.QAction("B", toolbar)
    act_bold.setCheckable(True)
    act_bold.triggered.connect(lambda on: toggle_format('bold', on))
    toolbar.addAction(act_bold)

    act_italic = QtWidgets.QAction("I", toolbar)
    act_italic.setCheckable(True)
    act_italic.triggered.connect(lambda on: toggle_format('italic', on))
    toolbar.addAction(act_italic)

    act_underline = QtWidgets.QAction("U", toolbar)
    act_underline.setCheckable(True)
    act_underline.triggered.connect(lambda on: toggle_format('underline', on))
    toolbar.addAction(act_underline)

    act_strike = QtWidgets.QAction("S", toolbar)
    act_strike.setCheckable(True)
    act_strike.triggered.connect(lambda on: toggle_format('strike', on))
    toolbar.addAction(act_strike)

    toolbar.addSeparator()

    # Font family and size
    font_box = QtWidgets.QFontComboBox(toolbar)
    font_box.currentFontChanged.connect(lambda f: _apply_font_family(text_edit, f.family()))
    toolbar.addWidget(font_box)

    size_box = QtWidgets.QComboBox(toolbar)
    for sz in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32]:
        size_box.addItem(str(sz), sz)
    size_box.setEditable(False)
    size_box.currentIndexChanged.connect(lambda _i: _apply_font_size(text_edit, size_box.currentData()))
    toolbar.addWidget(size_box)

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
    act_align_left = QtWidgets.QAction("L", toolbar, checkable=True)
    act_align_center = QtWidgets.QAction("C", toolbar, checkable=True)
    act_align_right = QtWidgets.QAction("R", toolbar, checkable=True)
    act_align_justify = QtWidgets.QAction("J", toolbar, checkable=True)
    for a in (act_align_left, act_align_center, act_align_right, act_align_justify):
        group_align.addAction(a)
        toolbar.addAction(a)
    act_align_left.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignLeft))
    act_align_center.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignHCenter))
    act_align_right.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignRight))
    act_align_justify.triggered.connect(lambda: text_edit.setAlignment(Qt.AlignJustify))

    toolbar.addSeparator()

    # Lists
    act_bullets = toolbar.addAction("• List", lambda: _toggle_list(text_edit, ordered=False))
    act_numbers = toolbar.addAction("1. List", lambda: _toggle_list(text_edit, ordered=True))

    toolbar.addSeparator()

    # Clear formatting, HR, Insert image
    toolbar.addAction("Clear", lambda: text_edit.setCurrentCharFormat(QTextCharFormat()))
    toolbar.addAction("HR", lambda: text_edit.textCursor().insertHtml("<hr/>"))
    toolbar.addAction("Image", lambda: _insert_image_via_dialog(text_edit))

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
    return toolbar


def _select_combo_value(combo: QtWidgets.QComboBox, value: int):
    for i in range(combo.count()):
        if combo.itemData(i) == value:
            combo.setCurrentIndex(i)
            return


def _apply_font_family(text_edit: QtWidgets.QTextEdit, family: str):
    fmt = QTextCharFormat()
    fmt.setFontFamily(family)
    _merge_format(text_edit, fmt)


def _apply_font_size(text_edit: QtWidgets.QTextEdit, size: int):
    if not size:
        return
    fmt = QTextCharFormat()
    fmt.setFontPointSize(float(size))
    _merge_format(text_edit, fmt)


def _pick_color_and_apply(text_edit: QtWidgets.QTextEdit, foreground: bool):
    color = QtWidgets.QColorDialog.getColor(parent=text_edit)
    if not color.isValid():
        return
    fmt = QTextCharFormat()
    if foreground:
        fmt.setForeground(color)
    else:
        fmt.setBackground(color)
    _merge_format(text_edit, fmt)


def _merge_format(text_edit: QtWidgets.QTextEdit, fmt: QTextCharFormat):
    cursor = text_edit.textCursor()
    if not cursor.hasSelection():
        cursor.select(cursor.WordUnderCursor)
    cursor.mergeCharFormat(fmt)
    text_edit.mergeCurrentCharFormat(fmt)


def _toggle_list(text_edit: QtWidgets.QTextEdit, ordered: bool):
    cursor = text_edit.textCursor()
    cursor.beginEditBlock()
    try:
        if ordered:
            cursor.insertList(QtWidgets.QTextListFormat.ListDecimal)
        else:
            cursor.insertList(QtWidgets.QTextListFormat.ListDisc)
    finally:
        cursor.endEditBlock()


def _insert_image_via_dialog(text_edit: QtWidgets.QTextEdit):
    path, _ = QtWidgets.QFileDialog.getOpenFileName(text_edit, "Insert Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp);;All Files (*)")
    if not path:
        return
    cursor = text_edit.textCursor()
    cursor.insertImage(path)
