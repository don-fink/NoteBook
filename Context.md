# NoteBook Project Context

## Project Overview
A PyQt5-based notebook/notes application with hierarchical organization (Binders → Sections → Pages) and rich text editing capabilities. Data is stored in SQLite databases with HTML content for pages.

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
- **UI Framework**: PyQt5 5.15.9
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
- [requirements.txt](requirements.txt) - Runtime dependencies (PyQt5)
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
