import sys
import re
from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem

def cell_to_index(cell):
    col = ord(cell[0].upper()) - ord('A')
    row = int(cell[1:]) - 1
    return row, col

def sum_range(table, start_cell, end_cell):
    start_row, start_col = cell_to_index(start_cell)
    end_row, end_col = cell_to_index(end_cell)
    total = 0
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            item = table.item(r, c)
            if item:
                try:
                    total += float(item.text())
                except ValueError:
                    pass
    return total

class Spreadsheet(QTableWidget):
    def __init__(self, rows, cols):
        super().__init__(rows, cols)
        self.setWindowTitle("PyQt Spreadsheet with SUM Support")
        self.formulas = {}  # {(row, col): "=SUM(A1:A12)"}
        self.cellChanged.connect(self.on_cell_changed)

    def on_cell_changed(self, row, col):
        item = self.item(row, col)
        if item:
            text = item.text()
            if text.startswith('=SUM('):
                self.formulas[(row, col)] = text
            elif (row, col) in self.formulas:
                # If user overwrites formula with plain value, remove it
                del self.formulas[(row, col)]
        self.recalculate_formulas()

    def recalculate_formulas(self):
        self.blockSignals(True)
        for (row, col), formula in self.formulas.items():
            match = re.match(r'=SUM\((\w+\d+):(\w+\d+)\)', formula)
            if match:
                start, end = match.groups()
                total = sum_range(self, start, end)
                display = f"{total:.2f}"  # Show result
                self.item(row, col).setText(display)
        self.blockSignals(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    sheet = Spreadsheet(20, 5)  # 20 rows, 5 columns (A–E)
    sheet.show()
    sys.exit(app.exec_())