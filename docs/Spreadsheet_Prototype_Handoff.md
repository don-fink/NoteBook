# Spreadsheet Prototype Handoff

Last updated: 2026-05-15

## Current Objective
Prototype spreadsheet capability in the NoteBook fork while keeping the original project safe.

**Active Phase**: Phase 2 - UI Consolidation & Conversion to .ui Files (PyQt5)

## Safety Setup (Completed)
- Stable project: c:\Python Projects\NoteBook_clean
- Working fork: c:\Python Projects\NoteBook_spreadsheet_fork
- Backup tag on GitHub: backup/pre-spreadsheet
- Prototype branch on GitHub: feature/spreadsheet-prototype

## Phase 2: UI Refactor (In Progress)

### Completed (2026-05-15)
- ✅ Created `ui_forms/` folder to consolidate all Designer UI files
- ✅ Moved 4 .ui files from root to `ui_forms/`:
  - main_window_2_column.ui
  - settings_dialog.ui
  - rename_database.ui
  - settings_dialog_old.ui
- ✅ Updated `ui_loader.py` to resolve .ui files from `ui_forms/` subfolder
- ✅ Updated `notebook.spec` for PyInstaller to include `ui_forms/` folder
- ✅ Verified file resolution working correctly
- ✅ Removed legacy 2-column/3-column layout menu remnants from main window UI
- ✅ Converted help/license/changelog markdown viewer dialog to `help_markdown_viewer.ui`
- ✅ Converted HTML Source Editor dialog to `html_source_editor.ui`

### In Progress
1. Convert remaining hard-coded dialogs to .ui files:
   - Image Properties dialog (ui_richtext.py:277)
   - Image Resize Overlay widget (ui_richtext.py:3263)
2. Decide whether to archive or remove `settings_dialog_old.ui`

### Known Issues to Address in Future Versions
- **Rich Text Formatting (v2.5 or 3.0)**: Current implementation is brittle and has too many capabilities. Plan to unify format handling and narrow scope in a dedicated refactor phase.

## Phase 1 (After Phase 2): PyQt5 → PyQt6 Migration
Once all UI files are in .ui format:
- Migrate imports and API calls to PyQt6
- .ui files are largely version-agnostic (XML-based)

## Architecture Direction (Spreadsheet Implementation)
- Store spreadsheet files in the media folder, not as DB blobs.
- Keep DB as metadata manifest only (page-to-file mapping).
- Keep templates in a dedicated templates folder.

## Initial Build Plan (When Ready)
1. Add DB migration for spreadsheet metadata table.
2. Add media subfolders:
   - media/spreadsheets/
   - media/spreadsheet_templates/
3. Add menu actions for:
   - Insert Spreadsheet
   - Insert From Spreadsheet Template
4. Implement basic flow:
   - Create/open spreadsheet file
   - Link file to page in DB
   - Open externally for troubleshooting
