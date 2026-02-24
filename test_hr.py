import sys
import os
os.environ['QT_QPA_PLATFORM'] = 'minimal'

from PyQt5 import QtWidgets
from PyQt5.QtGui import QTextCursor

app = QtWidgets.QApplication(sys.argv)
te = QtWidgets.QTextEdit()

# Simulate inserting HR like the code does
hr_html = '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse;"><tr><td style="border-top: 1px solid #000000; padding: 0; height: 1px;"></td></tr></table><p></p>'

te.insertHtml(hr_html)

# Get back the HTML Qt produces
output_html = te.toHtml()
print('=== RAW OUTPUT FROM QT ===')
print(output_html)
print()

# Now test the sanitizer
from ui_richtext import sanitize_html_for_storage
sanitized = sanitize_html_for_storage(output_html)
print('=== AFTER SANITIZE ===')
print(sanitized)
print()
print('=== BORDER-TOP IN SANITIZED? ===')
if 'border-top' in sanitized.lower():
    print('YES - border-top preserved')
else:
    print('NO - border-top STRIPPED!')
