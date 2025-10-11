# NoteBook

A PyQt5 desktop notebook app with binders, sections, and pages.

## Two‑Pane UI (current)

The app now uses a two‑column layout by default:
- Left: Binder → Sections → Pages tree (`notebookName`).
- Right: Page title (editable, bold) + rich text editor (`pageTitleEdit`, `pageEdit`).

Behavior highlights:
- Select a page to enable editing; otherwise the editor is read‑only.
- Autosave on typing (debounced), focus‑out, and Ctrl+S.
- Add/Rename/Delete Page/Section updates the tree immediately.
- State persists: last notebook/section/page, expanded binders, splitter sizes, theme, paste mode.

Shortcuts:
- Ctrl+S — save current page.
- Ctrl+Up / Ctrl+Down — reorder binders (focus left tree) and sections/pages (focus right panel).

Notes:
- Legacy tabbed wiring has been removed from setup; `setup_tab_sync` configures only the two‑pane UI.
- If you use a legacy `.ui` without `pageEdit`, setup becomes a no‑op; switch to `main_window_2_column.ui` for full functionality.

## Features

- Resizable split panes with saved layout
- Rich text editing with:
	- Bold/Italic/Underline/Strike, font family and size
	- Bulleted and numbered lists with nested levels
	- Classic outline numbering (I, A, 1, a, i, …) or Decimal
	- Tab/Shift+Tab or Ctrl+]/Ctrl+[ to indent/outdent lists
	- Insert horizontal rule and images
	- Paste modes: Rich, Text-only, Match Style, Clean
	- Single-click links open in your browser
- Default paste mode and list schemes are persisted
- Media storage and “Clean Unused Media” tool

### Paste modes
- Rich (default): standard paste.
- Text-only (Ctrl+Shift+V): inserts plain text, no formatting.
- Match Style: keeps structure but normalizes to current font and size.
- Clean: drops most inline styles/classes; keeps links, images, lists.

You can set the default paste mode from the Edit menu.

### Lists and indenting
- Use the toolbar buttons or keyboard:
	- Indent: Tab or Ctrl+]
	- Outdent: Shift+Tab or Ctrl+[
- Switch list scheme (Classic/Decimal) from the Format menu.

### Links
- Paste a URL in Match Style or Clean to auto-link it (e.g., https://example.com).
- Click a link to open it in your default browser.

## Dev setup

- Create and activate a virtual environment (recommended)
- Install developer requirements (runtime + Qt Designer tools)
- Run the app

### Windows PowerShell

```powershell
# From repo root
python -m venv .venv
& .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements-dev.txt

# Launch
& .\.venv\Scripts\python.exe .\main.py
```

## VS Code

- Open this folder in VS Code.
- Accept the recommended extensions when prompted.
- Run the task "Setup dev env (.venv + requirements-dev)" once from Terminal > Run Task.
- To run the app:
	- Use Run and Debug panel and select "Python: Run NoteBook", or
	- Run the task "Run NoteBook".

### Notes
- Qt Designer binaries and tools come from `pyqt5-tools` and related packages.
- If you only need runtime deps, use `requirements.txt` instead.
- Settings are stored in `settings.json` alongside the app.
- The UI is loaded from `main_window_2_column.ui` (if present) or falls back to `main_window.ui` via `ui_loader.py`.

## Troubleshooting
- If the UI fails to load with an UnsupportedPropertyError for `list`, remove any `sizes` property on QSplitter in the `.ui` file; set sizes in code instead.
- If panel sizes don’t restore, ensure `settings.json` is writable and that the app exits cleanly at least once to save initial sizes.
 - If the editor stays read‑only, make sure a page is selected in the left tree.
