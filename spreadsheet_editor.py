"""
spreadsheet_editor.py
Minimal in-app spreadsheet editor used by the spreadsheet management flow.
"""

from __future__ import annotations

import csv
import importlib
import os
import re

from qt_compat import QtCore, QtGui, QtWidgets


_ITEM_STYLE_ROLE = int(QtCore.Qt.ItemDataRole.UserRole) + 101


def _normalized_ext(path_or_ext: str) -> str:
    value = str(path_or_ext or "").strip().lower()
    if not value:
        return ""
    if value.startswith(".") and value.count(".") == 1 and os.path.sep not in value and "/" not in value:
        return value
    if os.path.sep in value or "/" in value or "." in os.path.basename(value):
        value = os.path.splitext(value)[1].lower()
    elif not value.startswith("."):
        value = f".{value}"
    return value


def _has_openpyxl() -> bool:
    try:
        importlib.import_module("openpyxl")

        return True
    except Exception:
        return False


def _openpyxl_module():
    return importlib.import_module("openpyxl")


def _cell_style_key(row_index: int, col_index: int) -> str:
    return f"{row_index},{col_index}"


def _xlsx_color_to_hex(color) -> str:
    try:
        rgb = str(getattr(color, "rgb", "") or "").strip()
        if len(rgb) == 8:
            rgb = rgb[-6:]
        if len(rgb) == 6 and rgb.upper() != "000000":
            return f"#{rgb.upper()}"
    except Exception:
        pass
    return ""


def _xlsx_style_from_cell(cell) -> dict:
    style: dict = {}
    try:
        if bool(getattr(cell.font, "bold", False)):
            style["bold"] = True
        if bool(getattr(cell.font, "italic", False)):
            style["italic"] = True
        if bool(getattr(cell.font, "underline", False)):
            style["underline"] = True
    except Exception:
        pass
    try:
        fill_type = str(getattr(cell.fill, "fill_type", "") or "")
        fill_color = _xlsx_color_to_hex(getattr(cell.fill, "fgColor", None))
        if fill_type == "solid" and fill_color:
            style["fill_color"] = fill_color
    except Exception:
        pass
    try:
        align = str(getattr(cell.alignment, "horizontal", "") or "").lower()
        if align in ("left", "center", "right"):
            style["align"] = align
    except Exception:
        pass
    try:
        number_format = str(getattr(cell, "number_format", "") or "")
        if number_format and number_format != "General":
            style["number_format"] = number_format
    except Exception:
        pass
    return style


def _apply_xlsx_style_to_cell(cell, style: dict):
    openpyxl = _openpyxl_module()
    font = openpyxl.styles.Font(
        bold=bool(style.get("bold", False)),
        italic=bool(style.get("italic", False)),
        underline="single" if bool(style.get("underline", False)) else None,
    )
    cell.font = font

    fill_color = str(style.get("fill_color") or "").strip()
    if fill_color:
        cell.fill = openpyxl.styles.PatternFill(
            fill_type="solid",
            fgColor=fill_color.lstrip("#"),
        )

    align = str(style.get("align") or "").strip().lower()
    if align in ("left", "center", "right"):
        cell.alignment = openpyxl.styles.Alignment(horizontal=align)

    number_format = str(style.get("number_format") or "").strip()
    if number_format:
        cell.number_format = number_format


def _coerce_xlsx_cell_value(value, style: dict | None = None):
    text = "" if value is None else str(value)
    if text == "":
        return None
    if text.startswith("'"):
        return text[1:]
    if text.startswith("="):
        return text

    number_format = str((style or {}).get("number_format") or "").strip()
    if number_format == "@":
        return text

    lowered = text.lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False

    if text == text.strip():
        if re.fullmatch(r"[+-]?[0-9]+", text):
            digits = text.lstrip("+-")
            if not (len(digits) > 1 and digits.startswith("0")):
                try:
                    return int(text)
                except Exception:
                    pass
        if re.fullmatch(r"[+-]?(?:[0-9]*\.[0-9]+|[0-9]+\.[0-9]*|[0-9]+[eE][+-]?[0-9]+|[0-9]*\.[0-9]+[eE][+-]?[0-9]+)", text):
            try:
                return float(text)
            except Exception:
                pass
    return text


def _item_style_meta(item: QtWidgets.QTableWidgetItem | None) -> dict:
    if item is None:
        return {}
    try:
        data = item.data(_ITEM_STYLE_ROLE)
        return dict(data) if isinstance(data, dict) else {}
    except Exception:
        return {}


def _set_item_style_meta(item: QtWidgets.QTableWidgetItem, style: dict):
    item.setData(_ITEM_STYLE_ROLE, dict(style or {}))


def _apply_item_style(item: QtWidgets.QTableWidgetItem, style: dict):
    font = item.font()
    font.setBold(bool(style.get("bold", False)))
    font.setItalic(bool(style.get("italic", False)))
    font.setUnderline(bool(style.get("underline", False)))
    item.setFont(font)

    fill_color = str(style.get("fill_color") or "").strip()
    if fill_color:
        item.setBackground(QtGui.QColor(fill_color))
    else:
        item.setBackground(QtGui.QBrush())

    align = str(style.get("align") or "").strip().lower()
    if align == "center":
        alignment = QtCore.Qt.AlignmentFlag.AlignCenter
    elif align == "right":
        alignment = QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter
    elif align == "left":
        alignment = QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter
    else:
        alignment = QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter
    item.setTextAlignment(int(alignment))

    _set_item_style_meta(item, style)


def supports_in_app_edit(path_or_ext: str) -> bool:
    ext = _normalized_ext(path_or_ext)
    if ext == ".csv":
        return True
    if ext == ".xlsx":
        return _has_openpyxl()
    return False


def _read_xlsx_grid(xlsx_path: str) -> list[list[str]]:
    rows: list[list[str]] = []
    if not os.path.isfile(xlsx_path):
        return rows
    workbook = _openpyxl_module().load_workbook(xlsx_path)
    sheet = workbook.active
    for row in sheet.iter_rows(values_only=True):
        values = ["" if cell is None else str(cell) for cell in row]
        while values and values[-1] == "":
            values.pop()
        if values:
            rows.append(values)
    return rows


def _write_xlsx_grid(xlsx_path: str, rows: list[list[str]]):
    folder = os.path.dirname(os.path.abspath(xlsx_path))
    if folder:
        os.makedirs(folder, exist_ok=True)

    openpyxl = _openpyxl_module()
    if os.path.isfile(xlsx_path):
        workbook = openpyxl.load_workbook(xlsx_path)
    else:
        workbook = openpyxl.Workbook()
    sheet = workbook.active

    max_row = max(sheet.max_row, 1)
    max_col = max(sheet.max_column, 1)
    for row_index in range(1, max_row + 1):
        for col_index in range(1, max_col + 1):
            sheet.cell(row=row_index, column=col_index).value = None

    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=col_index).value = _coerce_xlsx_cell_value(value)

    workbook.save(xlsx_path)


def read_spreadsheet_document(path: str) -> tuple[list[dict], int]:
    ext = _normalized_ext(path)
    if ext == ".xlsx":
        return _read_xlsx_document(path)
    return ([{"name": "Sheet1", "rows": read_csv_grid(path)}], 0)


def write_spreadsheet_document(path: str, sheets: list[dict], active_index: int = 0):
    ext = _normalized_ext(path)
    if ext == ".xlsx":
        _write_xlsx_document(path, sheets, active_index=active_index)
        return
    rows = sheets[0]["rows"] if sheets else []
    write_csv_grid(path, rows)


def read_csv_grid(csv_path: str) -> list[list[str]]:
    rows: list[list[str]] = []
    if not os.path.isfile(csv_path):
        return rows
    with open(csv_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            rows.append([str(cell) for cell in row])
    return rows


def write_csv_grid(csv_path: str, rows: list[list[str]]):
    folder = os.path.dirname(os.path.abspath(csv_path))
    if folder:
        os.makedirs(folder, exist_ok=True)
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(["" if cell is None else str(cell) for cell in row])


def read_spreadsheet_grid(path: str) -> list[list[str]]:
    ext = _normalized_ext(path)
    if ext == ".xlsx":
        return _read_xlsx_grid(path)
    return read_csv_grid(path)


def write_spreadsheet_grid(path: str, rows: list[list[str]]):
    ext = _normalized_ext(path)
    if ext == ".xlsx":
        _write_xlsx_grid(path, rows)
        return
    write_csv_grid(path, rows)


def _read_xlsx_document(xlsx_path: str) -> tuple[list[dict], int]:
    sheets: list[dict] = []
    active_index = 0
    if not os.path.isfile(xlsx_path):
        return ([{"name": "Sheet1", "rows": []}], 0)

    workbook = _openpyxl_module().load_workbook(xlsx_path)
    active_title = workbook.active.title if workbook.active is not None else None
    for idx, sheet in enumerate(workbook.worksheets):
        rows: list[list[str]] = []
        styles: dict = {}
        for row_index in range(1, sheet.max_row + 1):
            row_values: list[str] = []
            row_styles: dict[int, dict] = {}
            last_meaningful_col = 0
            for col_index in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row_index, column=col_index)
                value = "" if cell.value is None else str(cell.value)
                style = _xlsx_style_from_cell(cell)
                row_values.append(value)
                if value != "" or style:
                    last_meaningful_col = col_index
                if style:
                    row_styles[col_index - 1] = style
            if last_meaningful_col:
                output_row_index = len(rows)
                rows.append(row_values[:last_meaningful_col])
                for col_index, style in row_styles.items():
                    styles[_cell_style_key(output_row_index, col_index)] = style
        sheets.append({"name": sheet.title, "rows": rows, "styles": styles})
        if sheet.title == active_title:
            active_index = idx
    if not sheets:
        sheets = [{"name": "Sheet1", "rows": []}]
        active_index = 0
    return sheets, active_index


def _write_xlsx_document(xlsx_path: str, sheets: list[dict], active_index: int = 0):
    folder = os.path.dirname(os.path.abspath(xlsx_path))
    if folder:
        os.makedirs(folder, exist_ok=True)

    openpyxl = _openpyxl_module()
    workbook = openpyxl.Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    if not sheets:
        sheets = [{"name": "Sheet1", "rows": []}]

    created = []
    for sheet_meta in sheets:
        title = str(sheet_meta.get("name") or f"Sheet{len(created) + 1}")
        worksheet = workbook.create_sheet(title=title)
        created.append(worksheet)
        styles = dict(sheet_meta.get("styles") or {})
        for row_index, row in enumerate(sheet_meta.get("rows") or [], start=1):
            for col_index, value in enumerate(row, start=1):
                cell = worksheet.cell(row=row_index, column=col_index)
                style = dict(styles.get(_cell_style_key(row_index - 1, col_index - 1)) or {})
                cell.value = _coerce_xlsx_cell_value(value, style=style)
                if style:
                    _apply_xlsx_style_to_cell(cell, style)

    try:
        workbook.active = max(0, min(int(active_index), len(created) - 1))
    except Exception:
        pass
    workbook.save(xlsx_path)


class SpreadsheetEditorDialog(QtWidgets.QDialog):
    def __init__(self, spreadsheet_path: str, parent=None):
        super().__init__(parent)
        self._spreadsheet_path = spreadsheet_path
        self._ext = _normalized_ext(spreadsheet_path)
        self._supports_sheets = self._ext == ".xlsx"
        self._sheet_docs: list[dict] = []
        self._current_sheet_index = 0
        self._dirty = False
        self.setWindowTitle(f"Edit Spreadsheet - {os.path.basename(spreadsheet_path)}")
        self.resize(900, 520)

        layout = QtWidgets.QVBoxLayout(self)

        kind = "XLSX" if self._ext == ".xlsx" else "CSV"
        header = QtWidgets.QLabel(
            f"Editing {kind} file:\n{spreadsheet_path}",
            self,
        )
        header.setWordWrap(True)
        layout.addWidget(header)

        formula_row = QtWidgets.QHBoxLayout()
        self._cell_label = QtWidgets.QLabel("Cell: -", self)
        self._formula_edit = QtWidgets.QLineEdit(self)
        self._formula_edit.setPlaceholderText("Formula or value for the selected cell")
        formula_row.addWidget(self._cell_label)
        formula_row.addWidget(self._formula_edit, 1)
        layout.addLayout(formula_row)

        sheet_row = QtWidgets.QHBoxLayout()
        self._sheet_label = QtWidgets.QLabel("Sheet:", self)
        self._sheet_combo = QtWidgets.QComboBox(self)
        self._add_sheet_btn = QtWidgets.QPushButton("Add Sheet", self)
        self._rename_sheet_btn = QtWidgets.QPushButton("Rename Sheet", self)
        sheet_row.addWidget(self._sheet_label)
        sheet_row.addWidget(self._sheet_combo, 1)
        sheet_row.addWidget(self._add_sheet_btn)
        sheet_row.addWidget(self._rename_sheet_btn)
        layout.addLayout(sheet_row)

        format_row = QtWidgets.QHBoxLayout()
        self._bold_btn = QtWidgets.QPushButton("B", self)
        self._italic_btn = QtWidgets.QPushButton("I", self)
        self._underline_btn = QtWidgets.QPushButton("U", self)
        self._fill_btn = QtWidgets.QPushButton("Fill", self)
        self._clear_fill_btn = QtWidgets.QPushButton("Clear Fill", self)
        self._align_left_btn = QtWidgets.QPushButton("Align Left", self)
        self._align_center_btn = QtWidgets.QPushButton("Align Center", self)
        self._align_right_btn = QtWidgets.QPushButton("Align Right", self)
        for btn in (self._bold_btn, self._italic_btn, self._underline_btn):
            btn.setCheckable(True)
            btn.setMaximumWidth(40)
        for btn in (
            self._bold_btn,
            self._italic_btn,
            self._underline_btn,
            self._fill_btn,
            self._clear_fill_btn,
            self._align_left_btn,
            self._align_center_btn,
            self._align_right_btn,
        ):
            format_row.addWidget(btn)
        layout.addLayout(format_row)

        self._table = QtWidgets.QTableWidget(self)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
            | QtWidgets.QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.verticalHeader().setVisible(True)
        self._table.itemChanged.connect(self._mark_dirty)
        layout.addWidget(self._table)

        tools = QtWidgets.QHBoxLayout()
        self._add_row_btn = QtWidgets.QPushButton("Add Row", self)
        self._add_col_btn = QtWidgets.QPushButton("Add Column", self)
        self._save_btn = QtWidgets.QPushButton("Save", self)
        self._save_btn.setDefault(True)
        self._close_btn = QtWidgets.QPushButton("Close", self)
        for btn in (self._add_row_btn, self._add_col_btn, self._save_btn, self._close_btn):
            tools.addWidget(btn)
        layout.addLayout(tools)

        self._add_row_btn.clicked.connect(self._add_row)
        self._add_col_btn.clicked.connect(self._add_column)
        self._sheet_combo.currentIndexChanged.connect(self._switch_sheet)
        self._add_sheet_btn.clicked.connect(self._add_sheet)
        self._rename_sheet_btn.clicked.connect(self._rename_sheet)
        self._formula_edit.editingFinished.connect(self._apply_formula_edit)
        self._bold_btn.toggled.connect(lambda checked: self._apply_font_style("bold", checked))
        self._italic_btn.toggled.connect(lambda checked: self._apply_font_style("italic", checked))
        self._underline_btn.toggled.connect(lambda checked: self._apply_font_style("underline", checked))
        self._fill_btn.clicked.connect(self._choose_fill_color)
        self._clear_fill_btn.clicked.connect(self._clear_fill_color)
        self._align_left_btn.clicked.connect(lambda: self._apply_alignment("left"))
        self._align_center_btn.clicked.connect(lambda: self._apply_alignment("center"))
        self._align_right_btn.clicked.connect(lambda: self._apply_alignment("right"))
        self._table.currentCellChanged.connect(self._update_cell_controls)
        self._save_btn.clicked.connect(self.save)
        self._close_btn.clicked.connect(self.accept)

        self._load()

    def _load(self):
        self._sheet_docs, active_index = read_spreadsheet_document(self._spreadsheet_path)
        self._current_sheet_index = active_index
        self._refresh_sheet_controls()
        self._load_sheet(self._current_sheet_index)
        self._dirty = False

    def _refresh_sheet_controls(self):
        self._sheet_combo.blockSignals(True)
        self._sheet_combo.clear()
        for sheet_meta in self._sheet_docs:
            self._sheet_combo.addItem(str(sheet_meta.get("name") or "Sheet"))
        if self._sheet_docs:
            self._sheet_combo.setCurrentIndex(max(0, min(self._current_sheet_index, len(self._sheet_docs) - 1)))
        self._sheet_combo.blockSignals(False)

        visible = self._supports_sheets
        self._sheet_label.setVisible(visible)
        self._sheet_combo.setVisible(visible)
        self._add_sheet_btn.setVisible(visible)
        self._rename_sheet_btn.setVisible(visible)
        self._bold_btn.setVisible(visible)
        self._italic_btn.setVisible(visible)
        self._underline_btn.setVisible(visible)
        self._fill_btn.setVisible(visible)
        self._clear_fill_btn.setVisible(visible)
        self._align_left_btn.setVisible(visible)
        self._align_center_btn.setVisible(visible)
        self._align_right_btn.setVisible(visible)

    def _load_sheet(self, sheet_index: int):
        sheet_index = max(0, min(sheet_index, len(self._sheet_docs) - 1)) if self._sheet_docs else 0
        rows = self._sheet_docs[sheet_index]["rows"] if self._sheet_docs else []
        styles = dict(self._sheet_docs[sheet_index].get("styles") or {}) if self._sheet_docs else {}
        row_count = max(len(rows), 10)
        col_count = max((len(row) for row in rows), default=0)
        col_count = max(col_count, 5)

        self._table.blockSignals(True)
        self._table.setRowCount(row_count)
        self._table.setColumnCount(col_count)
        self._table.setHorizontalHeaderLabels([self._column_name(i) for i in range(col_count)])

        for row_index in range(row_count):
            source_row = rows[row_index] if row_index < len(rows) else []
            for col_index in range(col_count):
                value = source_row[col_index] if col_index < len(source_row) else ""
                item = QtWidgets.QTableWidgetItem(value)
                style = dict(styles.get(_cell_style_key(row_index, col_index)) or {})
                if style:
                    _apply_item_style(item, style)
                self._table.setItem(row_index, col_index, item)
        self._table.blockSignals(False)
        self._current_sheet_index = sheet_index
        if row_count and col_count:
            self._table.setCurrentCell(0, 0)
        self._update_cell_controls()

    def _sync_current_sheet_from_grid(self):
        if not self._sheet_docs:
            return
        rows, styles = self._current_sheet_payload()
        self._sheet_docs[self._current_sheet_index]["rows"] = rows
        if self._supports_sheets:
            self._sheet_docs[self._current_sheet_index]["styles"] = styles

    def _switch_sheet(self, sheet_index: int):
        if not self._sheet_docs:
            return
        if sheet_index < 0 or sheet_index >= len(self._sheet_docs):
            return
        if sheet_index == self._current_sheet_index:
            return
        self._sync_current_sheet_from_grid()
        self._load_sheet(sheet_index)

    def _add_sheet(self):
        base = "Sheet"
        existing = {str(sheet.get("name") or "") for sheet in self._sheet_docs}
        counter = 1
        while f"{base}{counter}" in existing:
            counter += 1
        self._sheet_docs.append({"name": f"{base}{counter}", "rows": []})
        self._refresh_sheet_controls()
        self._sheet_combo.setCurrentIndex(len(self._sheet_docs) - 1)
        self._dirty = True

    def _rename_sheet(self):
        if not self._sheet_docs:
            return
        current = self._sheet_docs[self._current_sheet_index]
        new_name, ok = QtWidgets.QInputDialog.getText(
            self,
            "Rename Sheet",
            "Sheet name:",
            text=str(current.get("name") or "Sheet"),
        )
        if not ok:
            return
        new_name = str(new_name or "").strip() or str(current.get("name") or "Sheet")
        current["name"] = new_name[:31] if self._supports_sheets else new_name
        self._refresh_sheet_controls()
        self._dirty = True

    def _mark_dirty(self, *_args):
        self._dirty = True

    def _add_row(self):
        self._sync_current_sheet_from_grid()
        self._table.insertRow(self._table.rowCount())
        self._dirty = True

    def _add_column(self):
        self._sync_current_sheet_from_grid()
        next_col = self._table.columnCount()
        self._table.insertColumn(next_col)
        self._table.setHorizontalHeaderItem(next_col, QtWidgets.QTableWidgetItem(self._column_name(next_col)))
        self._dirty = True

    def _ensure_item(self, row_index: int, col_index: int) -> QtWidgets.QTableWidgetItem:
        item = self._table.item(row_index, col_index)
        if item is None:
            item = QtWidgets.QTableWidgetItem("")
            self._table.setItem(row_index, col_index, item)
        return item

    def _selected_coordinates(self) -> list[tuple[int, int]]:
        coords: list[tuple[int, int]] = []
        seen = set()
        for selected_range in self._table.selectedRanges():
            for row_index in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                for col_index in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    key = (row_index, col_index)
                    if key not in seen:
                        seen.add(key)
                        coords.append(key)
        if not coords and self._table.currentRow() >= 0 and self._table.currentColumn() >= 0:
            coords.append((self._table.currentRow(), self._table.currentColumn()))
        return coords

    def _current_sheet_payload(self) -> tuple[list[list[str]], dict]:
        rows: list[list[str]] = []
        styles: dict = {}
        for row_index in range(self._table.rowCount()):
            row_values: list[str] = []
            row_styles: dict[int, dict] = {}
            last_meaningful_col = 0
            for col_index in range(self._table.columnCount()):
                item = self._table.item(row_index, col_index)
                value = item.text() if item is not None else ""
                style = _item_style_meta(item)
                if value != "" or style:
                    last_meaningful_col = col_index + 1
                row_values.append(value)
                if style:
                    row_styles[col_index] = style
            if last_meaningful_col:
                output_row_index = len(rows)
                rows.append(row_values[:last_meaningful_col])
                for col_index, style in row_styles.items():
                    styles[_cell_style_key(output_row_index, col_index)] = style
        return rows, styles

    def _update_cell_controls(self, *_args):
        row_index = self._table.currentRow()
        col_index = self._table.currentColumn()
        if row_index < 0 or col_index < 0:
            self._cell_label.setText("Cell: -")
            self._formula_edit.blockSignals(True)
            self._formula_edit.setText("")
            self._formula_edit.blockSignals(False)
            return

        item = self._ensure_item(row_index, col_index)
        self._cell_label.setText(f"Cell: {self._column_name(col_index)}{row_index + 1}")
        self._formula_edit.blockSignals(True)
        self._formula_edit.setText(item.text())
        self._formula_edit.blockSignals(False)

        style = _item_style_meta(item)
        for button, key in (
            (self._bold_btn, "bold"),
            (self._italic_btn, "italic"),
            (self._underline_btn, "underline"),
        ):
            button.blockSignals(True)
            button.setChecked(bool(style.get(key, False)))
            button.blockSignals(False)

    def _apply_formula_edit(self):
        row_index = self._table.currentRow()
        col_index = self._table.currentColumn()
        if row_index < 0 or col_index < 0:
            return
        item = self._ensure_item(row_index, col_index)
        item.setText(self._formula_edit.text())
        self._dirty = True

    def _apply_font_style(self, key: str, enabled: bool):
        if not self._supports_sheets:
            return
        changed = False
        for row_index, col_index in self._selected_coordinates():
            item = self._ensure_item(row_index, col_index)
            style = _item_style_meta(item)
            if enabled:
                style[key] = True
            else:
                style.pop(key, None)
            _apply_item_style(item, style)
            changed = True
        if changed:
            self._dirty = True
            self._update_cell_controls()

    def _choose_fill_color(self):
        if not self._supports_sheets:
            return
        color = QtWidgets.QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        changed = False
        for row_index, col_index in self._selected_coordinates():
            item = self._ensure_item(row_index, col_index)
            style = _item_style_meta(item)
            style["fill_color"] = color.name().upper()
            _apply_item_style(item, style)
            changed = True
        if changed:
            self._dirty = True
            self._update_cell_controls()

    def _clear_fill_color(self):
        if not self._supports_sheets:
            return
        changed = False
        for row_index, col_index in self._selected_coordinates():
            item = self._ensure_item(row_index, col_index)
            style = _item_style_meta(item)
            if "fill_color" in style:
                style.pop("fill_color", None)
                _apply_item_style(item, style)
                changed = True
        if changed:
            self._dirty = True
            self._update_cell_controls()

    def _apply_alignment(self, align: str):
        if not self._supports_sheets:
            return
        changed = False
        for row_index, col_index in self._selected_coordinates():
            item = self._ensure_item(row_index, col_index)
            style = _item_style_meta(item)
            style["align"] = align
            _apply_item_style(item, style)
            changed = True
        if changed:
            self._dirty = True
            self._update_cell_controls()

    def save(self):
        self._sync_current_sheet_from_grid()
        write_spreadsheet_document(
            self._spreadsheet_path,
            self._sheet_docs,
            active_index=self._current_sheet_index,
        )
        self._dirty = False
        QtWidgets.QMessageBox.information(self, "Spreadsheet Saved", "Changes saved.")

    def accept(self):
        if self._dirty:
            answer = QtWidgets.QMessageBox.question(
                self,
                "Unsaved Changes",
                "Save changes before closing?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel,
            )
            if answer == QtWidgets.QMessageBox.Cancel:
                return
            if answer == QtWidgets.QMessageBox.Yes:
                self.save()
        super().accept()

    @staticmethod
    def _column_name(index: int) -> str:
        name = ""
        index += 1
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            name = chr(ord("A") + remainder) + name
        return name


def edit_csv_file(csv_path: str, parent=None) -> int:
    return edit_spreadsheet_file(csv_path, parent=parent)


def edit_spreadsheet_file(spreadsheet_path: str, parent=None) -> int:
    dialog = SpreadsheetEditorDialog(spreadsheet_path, parent=parent)
    return dialog.exec()