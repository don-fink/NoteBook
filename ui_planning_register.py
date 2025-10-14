"""
ui_planning_register.py
Insert a Planning Register layout into the current page's QTextEdit.

Layout requested:
- Outer table: 1 row x 2 columns
- Inner table in the left cell: 7 rows x 3 columns
  - Header (top row) labels: "Description", "Estimated Cost", "Actual Cost"
  - Left column width = 50% of inner table; remaining 50% split across the other two columns
"""

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import (
    QTextTableFormat,
    QTextLength,
    QTextCharFormat,
    QFont,
    QTextCursor,
    QTextTableCellFormat,
    QTextBlockFormat,
    QColor,
    QBrush,
)
from PyQt5.QtCore import Qt


def _insert_inner_table_in_cursor(cursor):
    """Insert the 3x7 inner table at the given cursor.

    Column widths: [50%, 25%, 25%]. Header row populated.
    """
    inner_fmt = QTextTableFormat()
    inner_fmt.setCellPadding(4)
    inner_fmt.setCellSpacing(2)
    inner_fmt.setBorder(0.8)
    inner_fmt.setHeaderRowCount(1)
    inner_fmt.setColumnWidthConstraints(
        [
            QTextLength(QTextLength.PercentageLength, 50.0),
            QTextLength(QTextLength.PercentageLength, 25.0),
            QTextLength(QTextLength.PercentageLength, 25.0),
        ]
    )

    inner = cursor.insertTable(7, 3, inner_fmt)

    headers = ["Description", "Estimated Cost", "Actual Cost"]
    header_fmt = QTextCharFormat()
    header_fmt.setFontWeight(QFont.Bold)
    header_bg = QBrush(QColor(245, 245, 245))  # slight gray

    # Populate header cells
    for col, label in enumerate(headers):
        hcell = inner.cellAt(0, col)
        # Background for header row
        cfmt = hcell.format()
        cfmt.setBackground(header_bg)
        hcell.setFormat(cfmt)
        # Bold header text
        hcur = hcell.firstCursorPosition()
        hcur.mergeCharFormat(header_fmt)
        hcur.insertText(label)
        # Right-align numeric column headers (cols 1 and 2)
        if col in (1, 2):
            bfmt = QTextBlockFormat()
            bfmt.setAlignment(Qt.AlignRight)
            hcur = hcell.firstCursorPosition()
            hcur.mergeBlockFormat(bfmt)

    # Bottom row label: "Total" in left cell
    total_row_index = inner.rows() - 1
    total_label_cell = inner.cellAt(total_row_index, 0)
    total_label_cur = total_label_cell.firstCursorPosition()
    total_label_cur.insertText("Total")
    # Background for totals row
    totals_bg = QBrush(QColor(245, 245, 245))
    for c in range(inner.columns()):
        tcell = inner.cellAt(total_row_index, c)
        tfmt = tcell.format()
        tfmt.setBackground(totals_bg)
        tcell.setFormat(tfmt)

    # Right-align numeric columns for all rows (including totals)
    for r in range(inner.rows()):
        for c in (1, 2):
            bfmt = QTextBlockFormat()
            bfmt.setAlignment(Qt.AlignRight)
            ccur = inner.cellAt(r, c).firstCursorPosition()
            ccur.mergeBlockFormat(bfmt)

    return inner


def _insert_right_cost_table_in_cursor(cursor):
    """Insert a 2-column table with a shaded header row: headers 'Description' and 'Costs'.

    Column widths ~70%/30%. Right column is right-aligned. No totals row.
    """
    fmt = QTextTableFormat()
    fmt.setCellPadding(4)
    fmt.setCellSpacing(2)
    fmt.setBorder(0.8)
    fmt.setHeaderRowCount(1)
    fmt.setColumnWidthConstraints(
        [
            QTextLength(QTextLength.PercentageLength, 70.0),
            QTextLength(QTextLength.PercentageLength, 30.0),
        ]
    )
    table = cursor.insertTable(7, 2, fmt)

    headers = ["Description", "Costs"]
    header_fmt = QTextCharFormat()
    header_fmt.setFontWeight(QFont.Bold)
    header_bg = QBrush(QColor(245, 245, 245))

    for col, label in enumerate(headers):
        hcell = table.cellAt(0, col)
        cfmt = hcell.format()
        cfmt.setBackground(header_bg)
        hcell.setFormat(cfmt)
        hcur = hcell.firstCursorPosition()
        hcur.mergeCharFormat(header_fmt)
        hcur.insertText(label)
        # Right-align numeric header for 'Costs'
        if col == 1:
            bfmt = QTextBlockFormat()
            bfmt.setAlignment(Qt.AlignRight)
            hcur = hcell.firstCursorPosition()
            hcur.mergeBlockFormat(bfmt)

    # Right-align all rows in the cost column
    try:
        bfmt = QTextBlockFormat()
        bfmt.setAlignment(Qt.AlignRight)
        for r in range(table.rows()):
            ccur = table.cellAt(r, 1).firstCursorPosition()
            ccur.mergeBlockFormat(bfmt)
    except Exception:
        pass

    return table


def _cell_plain_text(text_edit: QtWidgets.QTextEdit, table, row: int, col: int) -> str:
    cell = table.cellAt(row, col)
    if not cell.isValid():
        return ""
    # Determine the selection range for the cell without modifying the editor state
    start = cell.firstCursorPosition().position()
    end = cell.lastCursorPosition().position()
    tmp = QTextCursor(text_edit.document())
    tmp.setPosition(start)
    tmp.setPosition(end, QTextCursor.KeepAnchor)
    return tmp.selection().toPlainText().strip()


def _cell_set_plain_text(text_edit: QtWidgets.QTextEdit, table, row: int, col: int, text: str):
    cell = table.cellAt(row, col)
    if not cell.isValid():
        return
    cur = QTextCursor(text_edit.document())
    cur.beginEditBlock()
    try:
        start = cell.firstCursorPosition().position()
        end = cell.lastCursorPosition().position()
        cur.setPosition(start)
        cur.setPosition(end, QTextCursor.KeepAnchor)
        cur.removeSelectedText()
        cur.insertText(text)
    finally:
        cur.endEditBlock()


def _parse_number(value: str) -> float:
    # Strip currency symbols and thousands separators; allow minus and dot
    if not value:
        return 0.0
    import re

    cleaned = re.sub(r"[^0-9.\-]", "", value)
    if cleaned.count(".") > 1:
        # If multiple dots, keep last as decimal separator
        parts = cleaned.split(".")
        cleaned = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(cleaned)
    except Exception:
        return 0.0


def _try_parse_number(value: str):
    """Return (float_value, True) if value looks numeric; else (0.0, False)."""
    if value is None:
        return 0.0, False
    s = value.strip()
    if not s:
        return 0.0, False
    try:
        v = _parse_number(s)
        # Consider it numeric if there's at least one digit in the text
        has_digit = any(ch.isdigit() for ch in s)
        return (v, True) if has_digit else (0.0, False)
    except Exception:
        return 0.0, False


def _currency_symbol() -> str:
    # Simple default; can be made configurable later
    return "$"


def _format_currency(value: float) -> str:
    return f"{_currency_symbol()}{value:,.2f}"


def _is_planning_register_table(text_edit: QtWidgets.QTextEdit, table) -> bool:
    try:
        if table.columns() < 3 or table.rows() < 3:
            return False
        # Check header row labels (best-effort)
        h0 = _cell_plain_text(text_edit, table, 0, 0).lower()
        h1 = _cell_plain_text(text_edit, table, 0, 1).lower()
        h2 = _cell_plain_text(text_edit, table, 0, 2).lower()
        if not (
            ("description" in h0)
            and ("estimated" in h1)
            and ("actual" in h2)
        ):
            return False
        # Bottom-left must be "Total"
        bl = _cell_plain_text(text_edit, table, table.rows() - 1, 0).strip().lower()
        return bl == "total"
    except Exception:
        return False


def _is_cost_list_table(text_edit: QtWidgets.QTextEdit, table) -> bool:
    """Detect the right-cell 2-column cost list table by headers."""
    try:
        if table.columns() != 2 or table.rows() < 2:
            return False
        h0 = _cell_plain_text(text_edit, table, 0, 0).lower()
        h1 = _cell_plain_text(text_edit, table, 0, 1).lower()
        return ("description" in h0) and ("cost" in h1)
    except Exception:
        return False


def _recalc_planning_totals(text_edit: QtWidgets.QTextEdit, table):
    try:
        rows = table.rows()
        if rows < 3:
            return
        # Sum rows 1..rows-2 for columns 1 and 2
        start_row = 1
        end_row = rows - 2
        sum_est = 0.0
        sum_act = 0.0
        for r in range(start_row, end_row + 1):
            v_est = _parse_number(_cell_plain_text(text_edit, table, r, 1))
            v_act = _parse_number(_cell_plain_text(text_edit, table, r, 2))
            sum_est += v_est
            sum_act += v_act
        # Write totals to bottom row (only if changed to reduce signal churn)
        total_row = rows - 1
        new_est = _format_currency(sum_est)
        new_act = _format_currency(sum_act)
        cur_est = _cell_plain_text(text_edit, table, total_row, 1)
        cur_act = _cell_plain_text(text_edit, table, total_row, 2)
        if cur_est != new_est:
            _cell_set_plain_text(text_edit, table, total_row, 1, new_est)
        if cur_act != new_act:
            _cell_set_plain_text(text_edit, table, total_row, 2, new_act)
    except Exception:
        pass


def _is_protected_cell(table, row: int, col: int) -> bool:
    if table is None:
        return False
    # Protect header row and bottom row entirely
    return row == 0 or row == (table.rows() - 1)


def _format_cost_cell_on_exit(text_edit: QtWidgets.QTextEdit, table, row: int, col: int):
    if table is None or col not in (1, 2) or _is_protected_cell(table, row, col):
        return
    raw = _cell_plain_text(text_edit, table, row, col)
    val, is_numeric = _try_parse_number(raw)
    if is_numeric:
        formatted = _format_currency(val)
        if raw != formatted:
            _cell_set_plain_text(text_edit, table, row, col, formatted)


class _PlanningRegisterWatcher(QtCore.QObject):
    """Watches a QTextEdit for cell-exit in cost columns and triggers totals recalculation."""

    def __init__(self, edit: QtWidgets.QTextEdit):
        super().__init__(edit)
        self._edit = edit
        self._prev = None  # (table, row, col)
        self._updating = False  # reentrancy guard to avoid recursive signal handling
        edit.cursorPositionChanged.connect(self._on_cursor_changed)
        edit.installEventFilter(self)

    def eventFilter(self, obj, event):
        try:
            if obj is self._edit:
                et = event.type()
                if et == QtCore.QEvent.FocusOut:
                    # On editor focus loss, normalize the current cost cell (if any)
                    # and recalc totals for the planning table.
                    if self._prev is not None and not self._updating:
                        table, row, col = self._prev
                        if table is not None and _is_planning_register_table(self._edit, table):
                            self._updating = True
                            try:
                                # Format currency if leaving a cost cell
                                _format_cost_cell_on_exit(self._edit, table, row, col)
                                _recalc_planning_totals(self._edit, table)
                            finally:
                                self._updating = False
                        elif table is not None and _is_cost_list_table(self._edit, table):
                            # For cost list tables, just format the cost column on exit
                            if col == 1 and row != 0:  # protect header row
                                self._updating = True
                                try:
                                    raw = _cell_plain_text(self._edit, table, row, col)
                                    val, is_num = _try_parse_number(raw)
                                    if is_num:
                                        fmt_val = _format_currency(val)
                                        if raw != fmt_val:
                                            _cell_set_plain_text(self._edit, table, row, col, fmt_val)
                                finally:
                                    self._updating = False
                elif et == QtCore.QEvent.KeyPress:
                    key = event.key()
                    mods = event.modifiers()
                    cur = self._edit.textCursor()
                    table = cur.currentTable()
                    if table is not None and _is_planning_register_table(self._edit, table):
                        cell = table.cellAt(cur)
                        row, col = cell.row(), cell.column()
                        # Block editing in protected cells
                        if _is_protected_cell(table, row, col):
                            if (
                                key in (Qt.Key_Backspace, Qt.Key_Delete, Qt.Key_Return, Qt.Key_Enter)
                                or (event.text() and event.text().strip())
                                or ((mods & Qt.ControlModifier) and key in (Qt.Key_V, Qt.Key_X, Qt.Key_Insert))
                                or ((mods & Qt.ShiftModifier) and key == Qt.Key_Insert)
                            ):
                                return True  # eat the key
                        # Tab on last editable row & right-most cost cell inserts a new row
                        if key in (Qt.Key_Tab,):
                            last_editable_row = table.rows() - 2
                            if row == last_editable_row and col == 2:
                                self._updating = True
                                try:
                                    # First normalize the current cost cell and recalc so value is formatted immediately
                                    try:
                                        _format_cost_cell_on_exit(self._edit, table, row, col)
                                    except Exception:
                                        pass
                                    # Insert a new row before the total row
                                    table.insertRows(table.rows() - 1, 1)
                                    # New row index (data row just above totals)
                                    new_row = last_editable_row + 1
                                    # Move caret to first cell of the new row
                                    ncell = table.cellAt(new_row, 0)
                                    self._edit.setTextCursor(ncell.firstCursorPosition())
                                    # Ensure numeric columns are right-aligned in the new row
                                    bfmt = QTextBlockFormat()
                                    bfmt.setAlignment(Qt.AlignRight)
                                    for new_c in (1, 2):
                                        ccur = table.cellAt(new_row, new_c).firstCursorPosition()
                                        ccur.mergeBlockFormat(bfmt)
                                    # Clear any inherited background on the new data row
                                    for clr_c in range(table.columns()):
                                        c = table.cellAt(new_row, clr_c)
                                        cfmt = c.format()
                                        cfmt.setBackground(QBrush(Qt.NoBrush))
                                        c.setFormat(cfmt)
                                    # Re-apply background on the (new) totals row
                                    totals_bg = QBrush(QColor(245, 245, 245))
                                    total_row_index = table.rows() - 1
                                    for cidx in range(table.columns()):
                                        tcell = table.cellAt(total_row_index, cidx)
                                        tfmt = tcell.format()
                                        tfmt.setBackground(totals_bg)
                                        tcell.setFormat(tfmt)
                                    # Re-apply background on the header row to ensure it stays shaded
                                    header_bg = QBrush(QColor(245, 245, 245))
                                    for hcol in range(table.columns()):
                                        hcell = table.cellAt(0, hcol)
                                        hfmt = hcell.format()
                                        hfmt.setBackground(header_bg)
                                        hcell.setFormat(hfmt)
                                    # Recalculate totals immediately
                                    _recalc_planning_totals(self._edit, table)
                                    # Update internal previous cell tracker to the new row to avoid double-format on next move
                                    try:
                                        self._prev = (table, new_row, 0)
                                    except Exception:
                                        pass
                                finally:
                                    self._updating = False
                                # Keep totals intact (they'll recalc on exit later); consume the key
                                return True
                    elif table is not None and _is_cost_list_table(self._edit, table):
                        cell = table.cellAt(cur)
                        row, col = cell.row(), cell.column()
                        # Protect header row from edits/paste/enter
                        if row == 0:
                            if (
                                key in (Qt.Key_Backspace, Qt.Key_Delete, Qt.Key_Return, Qt.Key_Enter)
                                or (event.text() and event.text().strip())
                                or ((mods & Qt.ControlModifier) and key in (Qt.Key_V, Qt.Key_X, Qt.Key_Insert))
                                or ((mods & Qt.ShiftModifier) and key == Qt.Key_Insert)
                            ):
                                return True
            return super().eventFilter(obj, event)
        except KeyboardInterrupt:
            # Swallow spurious interrupts that can be injected during debugging/interactive stops
            return True
        except Exception:
            # Avoid surfacing exceptions in the event filter; default processing
            return super().eventFilter(obj, event)

    def _current_cell(self):
        cur = self._edit.textCursor()
        table = cur.currentTable()
        if table is None:
            return None
        cell = table.cellAt(cur)
        return table, cell.row(), cell.column()

    def _on_cursor_changed(self):
        try:
            if self._updating:
                return
            now = self._current_cell()
            prev = self._prev
            # If we are leaving the previous cell (either to a different cell or out of the table entirely)
            if prev is not None:
                prev_table, prev_row, prev_col = prev
                if prev_table is not None and _is_planning_register_table(self._edit, prev_table):
                    left_prev_cell = False
                    if now is None:
                        left_prev_cell = True
                    else:
                        now_table, now_row, now_col = now
                        left_prev_cell = (
                            now_table != prev_table or now_row != prev_row or now_col != prev_col
                        )
                    if left_prev_cell and (prev_col in (1, 2)):
                        self._updating = True
                        try:
                            _format_cost_cell_on_exit(self._edit, prev_table, prev_row, prev_col)
                            _recalc_planning_totals(self._edit, prev_table)
                        finally:
                            self._updating = False
                elif prev_table is not None and _is_cost_list_table(self._edit, prev_table):
                    left_prev_cell = False
                    if now is None:
                        left_prev_cell = True
                    else:
                        now_table, now_row, now_col = now
                        left_prev_cell = (
                            now_table != prev_table or now_row != prev_row or now_col != prev_col
                        )
                    if left_prev_cell and (prev_col == 1) and prev_row != 0:
                        self._updating = True
                        try:
                            raw = _cell_plain_text(self._edit, prev_table, prev_row, prev_col)
                            val, is_num = _try_parse_number(raw)
                            if is_num:
                                fmt_val = _format_currency(val)
                                if raw != fmt_val:
                                    _cell_set_plain_text(self._edit, prev_table, prev_row, prev_col, fmt_val)
                        finally:
                            self._updating = False
            # Update previous reference after handling
            self._prev = now
        except Exception:
            pass


def insert_planning_register(window: QtWidgets.QMainWindow):
    """Insert the Planning Register tables into the active QTextEdit.

    If no page is open/editable, show a friendly message.
    """
    te: QtWidgets.QTextEdit = window.findChild(QtWidgets.QTextEdit, "pageEdit")
    if te is None or not te.isEnabled():
        QtWidgets.QMessageBox.information(
            window,
            "Planning Register",
            "Please open or create a page first.",
        )
        return

    cursor = te.textCursor()
    # Ensure we're at an insertion point (avoid splitting header fields etc.)
    if not cursor.atBlockStart():
        cursor.insertBlock()

    # Outer table: 1 row x 2 columns; force full-width (100%) container
    outer_fmt = QTextTableFormat()
    outer_fmt.setCellPadding(4)
    outer_fmt.setCellSpacing(3)
    outer_fmt.setBorder(1.0)
    outer_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
    # Evenly split the two columns
    outer_fmt.setColumnWidthConstraints(
        [
            QTextLength(QTextLength.PercentageLength, 50.0),
            QTextLength(QTextLength.PercentageLength, 50.0),
        ]
    )
    outer = cursor.insertTable(1, 2, outer_fmt)

    # Insert inner table into left cell (row 0, col 0)
    left_cell_cursor = outer.cellAt(0, 0).firstCursorPosition()
    inner = _insert_inner_table_in_cursor(left_cell_cursor)

    # Initialize totals once (will be kept up-to-date by watcher)
    _recalc_planning_totals(te, inner)

    # Insert cost list table into right cell (row 0, col 1)
    right_cell_cursor = outer.cellAt(0, 1).firstCursorPosition()
    _insert_right_cost_table_in_cursor(right_cell_cursor)

    # Optional: place the caret at the end of the right cell content (after the inserted table)
    try:
        after_outer = outer.cellAt(0, 1).lastCursorPosition()
        te.setTextCursor(after_outer)
    except Exception:
        pass

    # Install a single watcher per editor to keep totals dynamic on cell exit
    if not hasattr(te, "_planning_register_watcher"):
        te._planning_register_watcher = _PlanningRegisterWatcher(te)


def refresh_planning_register_styles(text_edit: QtWidgets.QTextEdit):
    """Reapply header/totals background and numeric right alignment for Planning Register tables in the editor.

    This is useful after loading HTML from storage to restore expected visuals
    in case a previous save stripped some styles.
    """
    if text_edit is None:
        return
    doc = text_edit.document()
    cur = QTextCursor(doc)
    seen = set()
    try:
        bg = QColor(245, 245, 245)
    except Exception:
        bg = None
    # Iterate blocks and collect unique tables by firstPosition
    while True:
        tbl = cur.currentTable()
        if tbl is not None:
            key = (tbl.firstPosition(), tbl.lastPosition())
            if key not in seen:
                seen.add(key)
                # If this is an outer 1x2 container, enforce 100% width with 50/50 split as well
                try:
                    if tbl.rows() == 1 and tbl.columns() == 2:
                        fmt_o = tbl.format()
                        fmt_o.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                        fmt_o.setColumnWidthConstraints(
                            [
                                QTextLength(QTextLength.PercentageLength, 50.0),
                                QTextLength(QTextLength.PercentageLength, 50.0),
                            ]
                        )
                        tbl.setFormat(fmt_o)
                except Exception:
                    pass
                if _is_planning_register_table(text_edit, tbl):
                    try:
                        rows, cols = tbl.rows(), tbl.columns()
                    except Exception:
                        rows, cols = 0, 0
                    # Ensure table fills its container and column widths are 50/25/25
                    try:
                        from PyQt5.QtGui import QTextLength, QTextTableFormat

                        fmt = tbl.format()
                        fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                        if cols >= 3:
                            fmt.setColumnWidthConstraints(
                                [
                                    QTextLength(QTextLength.PercentageLength, 50.0),
                                    QTextLength(QTextLength.PercentageLength, 25.0),
                                    QTextLength(QTextLength.PercentageLength, 25.0),
                                ]
                            )
                        tbl.setFormat(fmt)
                    except Exception:
                        pass
                    # Header row background and bold stays as-is; set background if missing
                    if rows >= 1 and cols >= 1 and bg is not None:
                        for c in range(cols):
                            cell = tbl.cellAt(0, c)
                            cf = cell.format()
                            cf.setBackground(bg)
                            cell.setFormat(cf)
                    # Totals row background
                    if rows >= 2 and bg is not None:
                        tr = rows - 1
                        for c in range(cols):
                            cell = tbl.cellAt(tr, c)
                            cf = cell.format()
                            cf.setBackground(bg)
                            cell.setFormat(cf)
                    # Right-align numeric columns across all rows
                    try:
                        bf = QTextBlockFormat()
                        bf.setAlignment(Qt.AlignRight)
                        for r in range(rows):
                            for c in (1, 2):
                                if c < cols:
                                    tcur = tbl.cellAt(r, c).firstCursorPosition()
                                    tcur.mergeBlockFormat(bf)
                    except Exception:
                        pass
                elif _is_cost_list_table(text_edit, tbl):
                    # For the right-side cost list tables, ensure width 100% and columns 70/30
                    try:
                        from PyQt5.QtGui import QTextLength

                        rows, cols = tbl.rows(), tbl.columns()
                        fmt = tbl.format()
                        fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100.0))
                        if cols >= 2:
                            fmt.setColumnWidthConstraints(
                                [
                                    QTextLength(QTextLength.PercentageLength, 70.0),
                                    QTextLength(QTextLength.PercentageLength, 30.0),
                                ]
                            )
                        tbl.setFormat(fmt)
                    except Exception:
                        pass
                    # Ensure header background and right alignment on the numeric column
                    try:
                        if bg is not None and tbl.rows() >= 1:
                            for c in range(tbl.columns()):
                                cell = tbl.cellAt(0, c)
                                cf = cell.format()
                                cf.setBackground(bg)
                                cell.setFormat(cf)
                        bf = QTextBlockFormat()
                        bf.setAlignment(Qt.AlignRight)
                        for r in range(tbl.rows()):
                            if tbl.columns() >= 2:
                                tcur = tbl.cellAt(r, 1).firstCursorPosition()
                                tcur.mergeBlockFormat(bf)
                    except Exception:
                        pass
        # Move to next block; stop at end
        if not cur.movePosition(QTextCursor.NextBlock):
            break


def ensure_planning_register_watcher(text_edit: QtWidgets.QTextEdit):
    """Ensure the planning register watcher is installed for the given QTextEdit.

    This enables currency formatting on cell-exit, totals recalculation, row-add on Tab,
    and protected rows even if the register was not inserted in the current session.
    """
    if text_edit is None:
        return
    if not hasattr(text_edit, "_planning_register_watcher"):
        text_edit._planning_register_watcher = _PlanningRegisterWatcher(text_edit)
