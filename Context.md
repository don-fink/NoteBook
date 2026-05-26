# NoteBook Project Context

## Session Snapshot (May 24, 2026)

- Continued PyQt6 migration debugging in `ui_richtext.py` and `ui_planning_register.py`
- Table behavior fixes completed locally:
    - `Distribute Columns Evenly` now preserves the current table width instead of expanding the table to full editor width
    - table width handling was aligned across `Fit Table to Width`, `Table Properties`, and manual column-width adjustments
    - the `Table Properties` dialog now opens at a more usable size
    - `Border width`, `Cell padding`, and `Space between cells` now apply visibly to the active table
    - nested table insertion now sizes the child table relative to the containing cell instead of the full editor width
    - nested table insertion no longer resizes the parent table because inserted-table detection now runs after the edit block completes
    - planning-register cost-list detection was tightened so ordinary nested `Description` / `Cost(s)` tables are no longer misclassified as special planning tables
- Current test result / stop point:
    - the last round of table fixes appeared to correct the reported live-editing problems
    - testing paused before checking preset save/reinsert behavior for the nested-table layout
- Best resume point next time:
    - manually test saving the nested 1x2 parent-table layout as a preset and reinserting it
    - note exactly what fails after reinsertion, if anything: nesting, widths, borders, spacing, cursor placement, editability, or planning/currency behavior leakage
    - if the preset path is still wrong, continue locally in `ui_richtext.py` preset capture/insert flow with the new repro details

## Session Snapshot (May 23, 2026)

- Release/update work completed locally:
    - fixed PyQt6 `QFileDialog.Options()` compatibility in `qt_compat.py`
    - fixed inline media cleanup so page-embedded images/videos are not removed as unused media
    - fixed two-pane navigation so page edits are saved before binder/section/page context switches
    - updated PyInstaller/release scripts to run non-interactively with `--noconfirm`
    - built a new PyInstaller release and published prerelease `pyqt6-preview-2026-05-23`
- Table preset work started in `ui_richtext.py` and `ui_planning_register.py`:
    - preset saving now tries to capture the outermost containing table so nested layouts can be reproduced exactly
    - preset insertion no longer rebuilds the old forced planning-register wrapper; it now reinserts the saved HTML directly
    - planning-register detection now has a structural fallback so HTML-saved tables can still be recognized after round-trip
- Current test result / stop point:
    - saving and repeating a single table appears to work correctly
    - nested table reproduction is still problematic and needs a better manual repro write-up before the next code pass
- Best resume point next time:
    - repeat the experiment with two tables inside a two-column outer table
    - document the exact problems seen when saving that layout as a preset and reinserting it
    - specifically note whether the failure is in capture, insertion, sizing, borders, cursor placement, or editing behavior after insertion
    - after documenting the exact symptoms, continue locally in `ui_richtext.py` preset save/insert flow

## Session Snapshot (May 22, 2026)

- Continued PyQt6 migration debugging in `main.py`, `qt_compat.py`, and `spell_check.py`
- Help/About menu work completed locally:
    - help menu actions were confirmed to be wired correctly from `main.py`
    - help document lookup was made more tolerant across script/executable/CWD launch locations
    - help dialog crash fixed by removing a Qt5-only `QDialogButtonBox.Close` lookup
    - `notebook.spec` updated so packaged builds include the `docs/` folder
    - `qt_compat.py` now aliases common `QMessageBox` icon/button/button-role enums for Qt6 compatibility
    - About dialog now works under the compatibility layer instead of failing on `QMessageBox.Information`
- Spell checking issue resolved locally:
    - spell checker was available and enabled, but the highlighter used a Qt5-only `QTextCharFormat.SpellCheckUnderline` enum
    - `spell_check.py` now resolves the underline style compatibly for Qt6
    - offscreen probe confirmed misspelled words now produce extra selections again
- Current stop point / best resume point:
    - investigate the `Insert Attachment` menu item, which currently produces an error when triggered
    - stay local to the attachment action wiring and its file-handling path in `main.py` on the next pass

## Session Snapshot (May 21, 2026)

- Continued PyQt6 migration testing focused on plain-table behavior in `ui_richtext.py`
- Bottom-row border issue was resolved for appended rows:
    - confirmed Qt6-specific behavior where `QTextTable.insertRows(rows, 1)` drops per-cell border formats when appending at the bottom
    - added a row-format copy helper so inserted rows inherit neighboring cell border formats
    - fixed a PyQt6 format-conversion mismatch where `QTextTableCellFormat(...)` constructor assumptions silently failed under the compatibility layer; switched affected paths to Qt6-safe table-cell format coercion
- Table insert dialog improvements completed:
    - dialog width widened slightly from its size hint so default numeric controls are fully visible
    - table gap control renamed from misleading `Cell spacing` to `Cell border gap`
    - free numeric gap control replaced with practical presets: `None`, `Small`, `Medium`, `Large`
    - border normalization no longer flattens nonzero gap values back to zero
- Table width behavior was investigated and partially redesigned:
    - confirmed mismatch between floating percentage-width tables on creation and fixed-width tables after column-width operations
    - new plain tables are now created as fixed-width tables based on the editor width at insertion time
    - fixed-width HTML is now generated directly for new plain tables when editor width is available, rather than relying only on later `QTextTableFormat` mutation
    - current dialog label remains percentage-based because it means `percent of current editor width at creation time`, not live floating width
    - user verified this is acceptable for now: created tables are fixed, old tables should be left alone and evolve naturally when edited
- Narrow validation completed during this session:
    - `py_compile` passed after each targeted change to `ui_richtext.py` and `ui_planning_register.py`
    - offscreen Qt probes verified:
        - bottom-appended rows now retain borders after format copy
        - selected nonzero table gap survives `_enforce_uniform_table_borders(...)`
        - generated HTML for new plain tables can now emit fixed pixel widths instead of `%` widths
- Current stop point / best resume point:
    - continue manual user testing of notebook workflows
    - specifically watch for any remaining mismatch between what the table dialogs say and what gets created or preserved after save/reload
    - one known nuance to revisit later if desired: the table width control still displays `Table width (% of current editor)` even though newly created tables are written as fixed-width tables immediately after creation

## Session Snapshot (May 20, 2026)

- Investigated unexpected app exits during PyQt6 migration when interacting with the page editor
- Reproduced two failure modes that shared the same editor activation path:
    - clicking inside a page could close the app
    - pressing Enter inside a page could close the app after a short delay
- Root-cause fixes applied:
    - removed `QApplication.processEvents()` from `two_pane_core.save_current_page()` to avoid re-entrant event processing during autosave
    - stopped passive toolbar sync in `ui_richtext.py` from mutating font/size combo boxes on cursor movement; this path was a likely PyQt6 native crash trigger when clicking into the editor
- Narrow validation completed:
    - `py_compile` passed for `main.py`, `two_pane_core.py`, `ui_richtext.py`, and `page_editor.py`
    - import smoke test passed for the same modules
- Manual user verification after the fix:
    - cursor remains active when clicking into a page
    - Enter works without terminating the app
    - edits persist after restart
    - main menu interaction that had been failing now works
- Table work continued after the editor crash fix:
    - plain table creation was switched to HTML insertion rather than live `QTextCursor.insertTable()`
    - new plain tables now honor requested width (for example 50% vs 100%)
    - new cell content wraps within the cell instead of expanding table width
    - blank-cell placeholder was corrected so new cells no longer start two rows tall
    - context-menu "Set Current Column Width" was fixed and is now working again
- Current unresolved table issue / best resume point:
    - inserting rows below from the bottom row still creates usable rows without visible borders
    - pressing Tab in the right-most cell to append a row shows the same missing-border behavior
    - inserting rows within the middle of a table appears to preserve borders correctly
    - additional clue: if a new table is created with nonzero cell spacing, that spacing carries into appended rows, but the configured cell border width does not
    - next investigation should stay local to row-append paths in `ui_richtext.py`, especially `_table_insert_rows_from_selection`, `_table_add_remove`, `_TableTabHandler`, and any bottom-row cell-format inheritance after `insertRows()`

## Session Snapshot (May 18, 2026)

- Began PyQt6 migration bootstrap on branch `pyqt6-migration`
- Added compatibility module: `qt_compat.py`
    - Imports PyQt6 first, falls back to PyQt5 if needed
    - Installs enum aliases so legacy `Qt.<name>` usage still works
    - Installs API aliases (`QAction`/`QShortcut` on `QtWidgets`, `exec_` aliases)
    - Registers a `PyQt5` module shim in `sys.modules` for legacy imports
- Updated `main.py` startup imports to use `qt_compat`
- Switched dependencies to PyQt6 in `requirements.txt` and `requirements-dev.txt`
- Next immediate task: run smoke startup, capture first runtime errors, and fix them iteratively

## Session Snapshot (May 17, 2026)

- Clean migration clone prepared at C:/Python Projects/NoteBook_pyqt6_migration
- Active branch in new clone: pyqt6-migration
- Remote in new clone: origin -> https://github.com/don-fink/NoteBook-pyqt6-lab.git
- Virtual environment created in new clone: .venv
- Dependency setup completed:
    - requirements.txt installed
    - requirements-dev.txt installed
    - pip check passed (no broken requirements)
    - import smoke test passed for PyQt6 and enchant
- Old folders were intentionally left untouched for later archival.

### Fast Resume Steps

1. Open C:/Python Projects/NoteBook_pyqt6_migration in VS Code.
2. Activate venv: & .\.venv\Scripts\Activate.ps1
3. Run app: .\.venv\Scripts\python.exe .\main.py
4. Continue migration work from this clean clone.

---

## Project Overview
A PyQt6-based notebook/notes application with hierarchical organization (Binders → Sections → Pages) and rich text editing capabilities. Data is stored in SQLite databases with HTML content for pages.

## Architecture Summary

### Core Files
| File | Purpose | Lines |
|------|---------|-------|
| [main.py](main.py) | Application entry point, UI wiring, menu connections | ~4,100 |
| [ui_richtext.py](ui_richtext.py) | Rich text toolbar, image/video handling, table operations | ~5,400 |
| [ui_planning_register.py](ui_planning_register.py) | Planning register tables with auto-totals | ~880 |
| [two_pane_core.py](two_pane_core.py) | Two-column UI logic, page load/save, auto-save | ~875 |
| [page_editor.py](page_editor.py) | Thin wrapper around two_pane_core | ~150 |
| [db_pages.py](db_pages.py) | Page CRUD operations | ~190 |
| [db_sections.py](db_sections.py) | Section CRUD operations | |
| [settings_manager.py](settings_manager.py) | Persistent settings (JSON) | |
| [media_store.py](media_store.py) | Media file storage for images/videos | |

### Technology Stack
- **UI Framework**: PyQt6 6.7.1 with `qt_compat` for migration shims
- **Rich Text Editor**: QTextEdit with custom toolbar layer
- **Document Model**: QTextDocument (HTML-based)
- **Storage**: SQLite database + HTML content
- **Media**: Stored in adjacent folders, referenced by relative paths

---

## Current Rich Text Implementation

### QTextEdit-Based System ([ui_richtext.py](ui_richtext.py))

**Toolbar Features:**
- Undo/Redo (Ctrl+Z/Ctrl+Y)
- Bold/Italic/Underline/Strikethrough
- Font family and size selection
- Text color and highlight
- Paragraph alignment (Left/Center/Right/Justify)
- Ordered and unordered lists with indent levels
- Horizontal rule insertion (1x1 table with top border)
- Image insertion with sizing modes (default/fit-width/original/custom)
- Video insertion with thumbnail extraction
- Table insertion and editing
- HTML source editor with syntax highlighting
- Paste modes: Rich, Plain text, Match style

**Key Functions:**
- `add_rich_text_toolbar()` - Main toolbar setup (~400 lines)
- `_install_table_context_menu()` - Right-click table operations
- `_insert_image_via_dialog()` / `_insert_image_from_path()` - Image handling
- `sanitize_html_for_storage()` - Clean HTML before saving
- `paste_text_only()` / `paste_match_style()` - Paste handlers

### Planning Register Tables ([ui_planning_register.py](ui_planning_register.py))

**Structure:**
- Outer container: 1×2 table (100% width, 50/50 split)
- Left cell: 3-column table (Description | Estimated Cost | Actual Cost)
  - Header row (shaded, bold)
  - Data rows
  - Totals row (shaded, auto-calculated)
- Right cell: 2-column cost list table (70/30 split)

**Key Features:**
- `_PlanningRegisterWatcher` class monitors cursor position
- Auto-formats currency on cell exit (`$1,234.56`)
- Auto-recalculates totals when values change
- Protected rows (header and totals cannot be edited)
- Tab on last data cell auto-inserts new row
- Right-aligned numeric columns

**Detection Functions:**
- `_is_planning_register_table()` - Checks header labels and "Total" in bottom-left
- `_is_cost_list_table()` - Checks for Description/Cost headers

---

## Auto-Save Mechanism ([two_pane_core.py](two_pane_core.py#L588-L630))

```
textChanged signal → _on_text_changed()
    ↓
Sets _two_col_dirty = True
Sets _autosave_ctx = (section_id, page_id)
Starts _autosave_timer (1200ms single-shot)
    ↓
Timer fires → _autosave_fire()
    ↓
Validates context matches current section/page
    ↓
Calls save_current_page()
```

**Also triggers save on:**
- Editor FocusOut event
- Page/section navigation
- Ctrl+S (manual save)

---

## Known Issues & Areas for Investigation

### 1. "Add a line from editor not persistent" Bug
**Potential causes to investigate:**
- Context mismatch in `_autosave_fire()` - if `_autosave_ctx` doesn't match current section/page, save is silently skipped
- Silent exception swallowing - most functions have `except Exception: pass` blocks
- Race condition between navigation and save timer

**Debugging approach:**
```python
# Add logging to _autosave_fire() in two_pane_core.py:
def _autosave_fire():
    ctx = getattr(window, "_autosave_ctx", None)
    sid_now = getattr(window, "_current_section_id", None)
    pid_now = ...
    print(f"Autosave context: {ctx}, current: ({sid_now}, {pid_now})")
    if isinstance(ctx, tuple) and len(ctx) == 2 and ctx[0] == sid_now and ctx[1] == pid_now:
        print("Context match - saving")
        save_current_page(window)
    else:
        print("Context MISMATCH - NOT saving!")
```

### 2. Brittle Planning Register Table Calculations
**Symptoms:** Totals work sometimes, then stop, fidgeting fixes them.

**Likely causes:**
- `_PlanningRegisterWatcher` losing track of table identity after edits
- Table detection (`_is_planning_register_table()`) failing due to whitespace/formatting changes in header text
- Cursor position tracking falling out of sync

**Investigation points:**
- Check `_prev` tracking in `_PlanningRegisterWatcher._on_cursor_changed()`
- Verify `_is_planning_register_table()` still matches after save/load cycle
- Look for cases where `_updating` flag prevents recalculation

---

## Planned Improvements

### Priority Features (User Requested)
1. **Spell Checking** - Can use `QSyntaxHighlighter` subclass for squiggly underlines
2. **Improved Undo/Redo** - Better grouping with `beginEditBlock()`/`endEditBlock()`
3. **Bug fixes** for persistence and table calculations

### Future Considerations
- Web/mobile version (suggests keeping HTML storage format portable)
- Image/video embedding (already has media_store infrastructure)

### Editor Upgrade Options Evaluated
| Option | Effort | Risk | Notes |
|--------|--------|------|-------|
| Enhance QTextEdit | Low-Med | Low | Recommended path - preserves 6,000+ lines of custom code |
| QWebEngineView + JS editor | Med-High | Med | Full HTML5/CSS3, but requires rebuilding table logic in JS |
| Markdown + Preview | Med | Med | UX paradigm shift, complex tables don't map well |
| Custom QGraphicsScene | Very High | High | Not recommended unless building commercial product |

---

## File-by-File Reference

### Database Layer
- [db_access.py](db_access.py) - Low-level database utilities
- [db_pages.py](db_pages.py) - Page table operations
- [db_sections.py](db_sections.py) - Section table operations
- [db_version.py](db_version.py) - Schema versioning/migration
- [schema.sql](schema.sql) - Database schema definition

### UI Layer
- [main_window_2_column.ui](main_window_2_column.ui) - Qt Designer UI file
- [ui_loader.py](ui_loader.py) - UI file loading
- [ui_logic.py](ui_logic.py) - UI business logic
- [ui_richtext.py](ui_richtext.py) - Rich text toolbar and features
- [ui_planning_register.py](ui_planning_register.py) - Planning register tables
- [ui_sections.py](ui_sections.py) - Section management UI
- [ui_tabs_shim.py](ui_tabs_shim.py) - Legacy tab-based UI compatibility
- [ui_toast.py](ui_toast.py) - Toast notifications
- [left_tree.py](left_tree.py) - Left navigation tree logic

### Support
- [settings_manager.py](settings_manager.py) - JSON-based settings persistence
- [media_store.py](media_store.py) - Image/video file management
- [backup.py](backup.py) - Database backup/export/import
- [maintenance_order.py](maintenance_order.py) - Order index normalization

### Configuration
- [requirements.txt](requirements.txt) - Runtime dependencies (PyQt6)
- [requirements-dev.txt](requirements-dev.txt) - Development dependencies
- [pyproject.toml](pyproject.toml) - Project metadata
- [themes/](themes/) - QSS stylesheets (default.qss, high-contrast.qss)

---

## Development Commands

```powershell
# Activate virtual environment
& .\.venv\Scripts\Activate.ps1

# Run application
python main.py

# Run with safe mode (disables risky UI hooks)
$env:NOTEBOOK_SAFE_MODE = "1"; python main.py

# Enable image resize overlay (opt-in feature)
$env:NOTEBOOK_ENABLE_IMAGE_RESIZE = "1"; python main.py
```

---

## Next Steps for AI Assistant

When continuing work on this project:

1. **For persistence bug**: Add logging to `_autosave_fire()` and `save_current_page()` to trace when saves are skipped
2. **For table bugs**: Add logging to `_is_planning_register_table()` and `_recalc_planning_totals()` 
3. **For spell check**: Research `QSyntaxHighlighter` integration with hunspell or similar
4. **For undo improvements**: Audit all table operations for proper `beginEditBlock()`/`endEditBlock()` usage

---

*Generated: February 22, 2026*
