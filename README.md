# NoteBook

A PyQt5 desktop notebook app with binders, sections, and pages.

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

### Notes
- Qt Designer binaries and tools come from `pyqt5-tools` and related packages.
- If you only need runtime deps, use `requirements.txt` instead.
- Settings are stored in `settings.json` alongside the app.
- The UI is loaded from `main_window_5.ui` via `ui_loader.py`.

## Troubleshooting
- If the UI fails to load with an UnsupportedPropertyError for `list`, remove any `sizes` property on QSplitter in the `.ui` file; set sizes in code instead.
- If panel sizes don’t restore, ensure `settings.json` is writable and that the app exits cleanly at least once to save initial sizes.
