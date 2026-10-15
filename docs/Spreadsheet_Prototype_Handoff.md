# Spreadsheet Prototype Handoff

Last updated: 2026-06-03

## Current Objective
Ship spreadsheet support as managed linked files stored under the notebook media area, with editing handled by the user's spreadsheet application rather than an inline notebook surface.

**Active Phase**: Spreadsheet simplification and rollback on the `pyqt6-migration` branch

## Safety Setup (Completed)
- Project folder: c:\Python Projects\NoteBook (older folders were consolidated; the original PyQt5 repo's backups are archived in c:\Python Projects\Backups\NoteBook_clean_archive)
- Backup tag on GitHub: backup/pre-spreadsheet
- Prototype branch on GitHub: feature/spreadsheet-prototype

## Current Spreadsheet Status (In Progress)

### Completed
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
- ✅ Removed the dead `View -> Rich Text Tools` menu path
- ✅ Added spreadsheet insert flows for new, template, and existing files
- ✅ Store spreadsheet files under `db.media/media/spreadsheets/` with DB metadata in `spreadsheets`
- ✅ Added `Manage Page Spreadsheets...` for open/folder/path/link insertion/path copy actions
- ✅ Spreadsheet insertion now uses a simple page link instead of embedding a preview table in the note
- ✅ Removed embedded spreadsheet preview generation and preview-refresh actions from the notebook page workflow
- ✅ Removed spreadsheet-preview-specific rich-text interception and context menu behavior
- ✅ Load/save now normalize older embedded spreadsheet preview blocks back to simple links so existing pages are cleaned up over time
- ✅ Spreadsheet work is now routed to the external spreadsheet application via normal file open behavior
- ✅ Guarded rich-text row insertion against stale/deleted `QTextTable` objects in PyQt6

### Current State
1. Spreadsheets are treated as linked media assets, not inline spreadsheet widgets.
2. Notebook pages now use a simple `Spreadsheet: ...` link as the insertion surface.
3. Existing legacy preview blocks are collapsed back to links when pages are loaded and saved.
4. Spreadsheet editing should happen in Excel, LibreOffice, or another external spreadsheet application.

### Next Session Priorities
1. Add orphan spreadsheet cleanup for files left behind under `media/spreadsheets/` after links are removed from page content.
2. Investigate the table creation regression where a newly inserted table appears briefly and then disappears a couple of seconds later.
3. Fix link-edit behavior so typing after an existing link does not extend the hyperlink into the new text.
4. Retest the simplified spreadsheet workflow in the running app:
   - insert a new/template/existing spreadsheet and confirm a plain link is inserted
   - use `Manage Page Spreadsheets...` and confirm `Open` launches the external spreadsheet app
   - reopen a page that previously had an embedded spreadsheet preview and confirm it is cleaned back to a link
5. Decide whether the page manager should later gain lightweight metadata fields such as description or note text for each spreadsheet link.

### Known Issues to Address in Future Versions
- **Rich Text Formatting (v2.5 or 3.0)**: Current implementation is brittle and has too many capabilities. Plan to unify format handling and narrow scope in a dedicated refactor phase.
- **Spreadsheet Orphan Cleanup**: removing spreadsheet links from page HTML does not currently remove the underlying files from `media/spreadsheets/`; add cleanup/reconciliation later.
- **Spreadsheet Metadata UX**: if users need more context than a link title, add optional description/help text rather than rebuilding inline spreadsheet rendering.
- **Table Insert Regression**: newly inserted tables may disappear shortly after creation; inspect autosave/load or sanitization interactions.
- **Link Cursor Behavior**: typing after a link can continue the anchor formatting into subsequent text; adjust cursor/link termination behavior in the editor.

## Migration Status
The PyQt6 migration is already underway on the `pyqt6-migration` branch.
- Continue replacing remaining legacy Qt imports/usages with `qt_compat` or native PyQt6-safe equivalents
- Keep `.ui` files version-agnostic where possible

## Architecture Direction (Spreadsheet Implementation)
- Store spreadsheet files in the media folder, not as DB blobs.
- Keep DB as metadata manifest only (page-to-file mapping).
- Keep templates in a dedicated templates folder.
- Preserve the ability to open spreadsheet files outside the app for troubleshooting and direct editing.
- Treat notebook pages as the place to describe and organize spreadsheet files, not as the place to reproduce spreadsheet interaction.
- Keep any future spreadsheet preview work strictly optional, static, and low-complexity.

## Initial Build Plan (When Ready)
1. Expand the basic flow beyond starter spreadsheet creation:
   - Create/open spreadsheet file
   - Link file to page in DB
   - Open externally for troubleshooting
2. Add richer template management under `media/spreadsheet_templates/`
3. Decide whether to add a dedicated spreadsheet management surface beyond link insertion
