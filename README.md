# NoteBook

A lightweight, portable desktop notebook application for Windows built with PyQt5. Organize your notes in binders, sections, and pages with rich text editing.

![NoteBook Screenshot](screenshots/main_page.png)

## Features

- **Two-pane interface** — Browse binders/sections/pages on the left, edit on the right
- **Rich text editing** — Bold, italic, underline, strikethrough, fonts, sizes
- **Lists** — Bulleted and numbered lists with nested levels and multiple numbering schemes
- **Tables** — Insert tables with currency column formatting and auto-totals
- **Images** — Embed images with resize handles
- **Auto-save** — Changes saved automatically as you type
- **Portable** — No installer required; settings stored in `%LOCALAPPDATA%\NoteBook\`
- **Themes** — Default and high-contrast themes included

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Save | Ctrl+S |
| Paste as plain text | Ctrl+Shift+V |
| Indent list | Tab or Ctrl+] |
| Outdent list | Shift+Tab or Ctrl+[ |
| Reorder items | Ctrl+Up / Ctrl+Down |

## Installation (Windows)

### Download Release

1. Download `NoteBook_Release.zip` from the [latest release](../../releases/latest)
2. Extract to any folder (e.g., `C:\Program Files\NoteBook\`)
3. Run `NoteBook.exe` — no installation required!
4. (Optional) Run `add_to_start_menu.cmd` to add a Start Menu shortcut

### Build from Source

```powershell
# Clone and enter the repo
git clone https://github.com/don-fink/NoteBook.git
cd NoteBook

# Create virtual environment
python -m venv .venv
& .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements-dev.txt

# Run the app
python main.py

# Build executable (optional)
.\build.cmd
```

## Usage

### Getting Started

1. Launch the app — a default notebook is created automatically
2. Use the left panel to create binders (notebooks), sections, and pages
3. Click a page to start editing in the right panel
4. Your changes are saved automatically

### Paste Modes

Access from the Edit menu:

- **Rich** (default) — Standard paste with formatting
- **Text-only** (Ctrl+Shift+V) — Plain text, no formatting
- **Match Style** — Keeps structure, normalizes to current font
- **Clean** — Removes most styling, keeps links and images

### Currency Columns

Right-click in a table and select "Mark Column(s) as Currency + Total" to:
- Format numbers as currency (`$1,234.56`)
- Auto-calculate column totals
- Updates automatically when you edit cells

## Development

### VS Code Setup

1. Open this folder in VS Code
2. Accept the recommended extensions when prompted
3. Run task: **Terminal → Run Task → Setup dev env**
4. Debug with: **Run and Debug → Python: Run NoteBook**

### Project Structure

- `main.py` — Application entry point
- `ui_logic.py` — Main window logic and event handling
- `page_editor.py` — Rich text editor implementation
- `db_*.py` — Database access layer
- `services/` — Business logic services
- `themes/` — QSS stylesheets

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.

---

## Troubleshooting

### App crashes intermittently

Run in Safe Mode to disable heavy overlays:

```powershell
$env:NOTEBOOK_SAFE_MODE='1'; python main.py
```

Check these log files for details:
- `crash.log` — Python exceptions and Qt messages
- `native_crash.log` — Low-level crash traces

### Editor stays read-only

Make sure a page is selected in the left tree panel.

### Panel sizes don't restore

Ensure `settings.json` is writable and exit the app cleanly at least once.

### UI fails to load

If you see `UnsupportedPropertyError for 'list'`, the `.ui` file may have incompatible properties. Use the included `main_window_2_column.ui`.

---

## Advanced

<details>
<summary>Order Index Normalization</summary>

Over time, item ordering indexes can have gaps or duplicates. Use **Tools → Normalize Page Order** to fix:

1. Sorts items by current order
2. Reassigns sequential values (1, 2, 3...)
3. Optionally creates a backup first

This only changes ordering numbers, not content.

</details>

<details>
<summary>Release Packaging Notes</summary>

When building releases:

- Do NOT include `settings.loc` — it's machine-specific
- Settings default to `%LOCALAPPDATA%\NoteBook\settings.json`
- Run `create_release_simple.ps1` to package for distribution

</details>
