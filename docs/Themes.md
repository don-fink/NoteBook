# Making a Theme

A theme is a Qt stylesheet (`.qss`) file. NoteBook lists every `.qss` file in its own `themes` folder and in a `themes` folder next to your `settings.json`. The file name becomes the theme's name: `dark-slate.qss` appears as **Dark Slate** under **Tools → Settings → Themes**. A file with the same name as a theme that comes with NoteBook replaces it.

The easiest way to start is to copy `default.qss` from NoteBook's `themes` folder, rename it, and change its colors.

## Icons and other files

A `url(...)` that isn't a full path is read relative to the theme file's folder, so a theme can bring its own icons. The themes that come with NoteBook keep theirs in `themes/icons/<theme name>/`.

## What a theme can style

Besides ordinary Qt widgets (`QMenu`, `QPushButton`, `QTreeView` and so on), NoteBook names these parts so a theme can style them:

| Part | Selector | Notes |
|------|----------|-------|
| Message windows | `QMessageBox#nbMessage` | Every pop-up message and confirmation. |
| Message kind | `QMessageBox#nbMessage[kind="info"]` | `kind` is `info`, `warning`, `error`, `question` or `about`. NoteBook's themes use it to color the bar down the left edge. |
| Message headline | `QMessageBox#nbMessage QLabel#qt_msgbox_label` | The first, short line. |
| Message explanation | `QMessageBox#nbMessage QLabel#qt_msgbox_informativelabel` | The line under the headline. |
| Message buttons | `QMessageBox#nbMessage QPushButton[role="primary"]` | `role` is `primary` (the main action), `destructive` (deletes or overwrites) or `secondary`. |
| Message icons | `QMessageBox { messagebox-warning-icon: url(...); }` | Also `messagebox-information-icon`, `messagebox-question-icon` and `messagebox-critical-icon`. |
| Name and number prompts | `QInputDialog#nbPrompt` | Rename, page title, width and similar prompts. Set the room around and between their controls with `qproperty-boxPadding: 18; qproperty-boxSpacing: 10;` (pixels). |
| Notices | `QLabel#nbToast` | The short notices near the bottom of the window. |
| Tree drag line | `QFrame#leftTreeDropIndicator` and `QTreeView::drop-indicator` | Shown while dragging binders and pages. |
| Image resize handles | `QWidget#nbResizeHandles { qproperty-accentColor: #2f6fb5; }` | The outline and corner handles on images. |
| Search box | `QWidget#nbSearchBar`, `QLineEdit#nbSearchBox`, `QLabel#nbSearchCount`, `QLabel#nbSearchScopeLabel`, `QToolButton#nbSearchEverywhere` | `QLineEdit#nbSearchBox[noResults="true"]` styles the box when nothing is found. Set the highlight on the page with `QWidget#nbSearchBar { qproperty-highlightColor: #ffd54f; }`. |
| Formatting toolbar | `QWidget#richTextToolbar` and `QWidget#richTextToolbar QToolButton` | The bar above the page. Its icons are drawn in the buttons' `color`. |

If a theme doesn't style the notices, the tree drag line or the formatting toolbar, NoteBook uses its own built-in look for them.

## Dark themes

Page content keeps the colors it was written with, and most text in pages is stored as black or near-black. A dark theme should therefore keep `QTextEdit` (the page) a light color, as the Dark theme does with its dimmed paper color. The same goes for `QTextBrowser` (help), `QPlainTextEdit` (HTML source) and `QTableView` (spreadsheets).
