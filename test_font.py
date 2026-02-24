"""Quick test of font application."""
import sys
from PyQt5 import QtWidgets
from PyQt5.QtGui import QFont, QTextCharFormat

app = QtWidgets.QApplication(sys.argv)

te = QtWidgets.QTextEdit()
te.setPlainText("Hello World")
te.show()

# Select all text
cursor = te.textCursor()
cursor.select(cursor.Document)
te.setTextCursor(cursor)

print(f"Selection: '{cursor.selectedText()}'")
print(f"Has selection: {cursor.hasSelection()}")

# Apply font
fmt = QTextCharFormat()
fmt.setFontFamily("Courier New")
print(f"Applying font: Courier New")

cursor.mergeCharFormat(fmt)
te.setTextCursor(cursor)

# Check result
cursor = te.textCursor()
cursor.select(cursor.Document)
result_fmt = cursor.charFormat()
print(f"Result font family: {result_fmt.fontFamily()}")

# Check HTML output
html = te.toHtml()
print(f"\nHTML contains 'Courier': {'Courier' in html}")
print(f"HTML snippet: {html[:500]}...")

sys.exit(0)
