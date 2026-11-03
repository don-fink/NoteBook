
## Unreleased
### Added
- Dark theme: a dark frame around a dimmed paper page, so existing pages and tables stay readable. Choose it under Tools → Settings → Themes.
- Message windows have a new look that comes from the theme: a colored bar down the left edge for each kind of message (information, question, warning, error), matching icons, and a highlighted main button. Prompts, notices, the tree's drag line and the image resize handles follow the theme too. `docs/Themes.md` lists what a theme can style.
- Resize images and video thumbnails by dragging a corner handle, which appears when the mouse is over the image. The image always keeps its shape, can't be dragged wider than the page or its table cell, and changes size on release as one undo step. Alt text, title and video links are kept.

### Changed
- Messages are worded the same way throughout: a short headline, then a line explaining or saying what to do next. Technical error details are under **Show Details…** instead of in the message.
- Confirmations name the action on their buttons (**Delete**, **Replace**, **Move**) instead of Yes and No. For deleting or overwriting, Enter chooses Cancel.
- Fit to editor width leaves room for the scrollbar, so a fitted image no longer ends up slightly wider than the page.

### Fixed
- Print and Print Preview failed with "type object 'QPrinter' has no attribute 'HighResolution'", left over from the move to PyQt6. Images on printed pages also printed as blank placeholders; they now print.


## 1.7.0 - 2026-11-02
### Added
- Table and Planning Register presets can be inserted again: Insert → Insert Table from Preset… and Insert Planning Register from Preset…, also under Insert on the right-click menu. Presets come back exactly as saved, text and amounts included.
- Insert → Insert Table…, Manage Table Presets… and Manage Planning Register Presets… (rename and delete).
- Tools → Import Preset Tables and Registers…: choose presets from another database to copy into the open one. Name clashes can be replaced, renamed or skipped, one at a time or all at once.

### Changed
- Menus have a border, rounded corners and a stronger highlight in both themes.
- Themes are read from the themes folder instead of a fixed list, so adding a `.qss` file adds a theme. Themes can also be added in a `themes` folder next to `settings.json`, without rebuilding; one with the same name replaces the shipped theme.
- Presets are stored in the database, in separate lists for tables and Planning Registers, so they're included in backups. Presets saved in settings by earlier versions are copied into each database the first time its presets are used.
- Saving a preset under a name already in use asks before replacing it. Renaming does the same.
- Save Table as Preset inside a Planning Register saves just that table; Save Planning Register as Preset saves the whole register, from anywhere inside it. Both are on the right-click menu.
- On first launch, NoteBook asks you to create a new database or open an existing one, and where to keep it, instead of creating `notes.db` in the folder it was started from. It also asks if the database it used last time can't be found, instead of creating an empty one in its place.
- The Enter key diagnostic log (`richtext_key_debug.log`) is off unless `NOTEBOOK_RICHTEXT_KEY_DEBUG=1` is set.

### Fixed
- After creating or opening another database, the page from the previous database stayed in the editor and could still be edited, and new images went to the previous database's media folder. The page is now saved and closed first, and media follows the database that's open.
- Fit Table to Width did nothing visible inside a Planning Register. It now fits the whole register to the page, and its tables follow. A table nested in another table's cell now fits that cell instead of the whole page.
- A Planning Register whose cost columns were marked as currency ("Estimated Cost ($)") stopped being treated as a register after its page was reopened or it was inserted from a preset, losing its own totals, cost formatting and number alignment.


## 1.6.0 - 2026-11-01
### Added
- Spreadsheets: add one to a page as a new file, from a saved template, or by copying an existing file, and edit it in a built-in editor (CSV and XLSX).
- Add Link: select text, right-click and choose Add Link… to link to a web address, email address or file path.

### Changed
- The last page you had open reopens at startup, instead of leaving the editor empty until you click a page.
- Images are resized from the right-click menu (Image Properties, Fit to editor width, Reset to original size). The small image button that appeared near the caret is gone.
- Insert and paste commands now show a message when no editable page is open, instead of adding content that would never be saved.

### Fixed
- Saving could lose formatting or trailing text, and text such as "<" could be read as markup on the next load.
- Table widths set by hand were reset to defaults when a page loaded, and borderless tables gained borders.
- Column widths no longer drift when a column is added or removed. Planning Registers no longer change width with the window.
- Currency and Planning Register tables lost their formatting and totals after a restart until a column was re-marked.
- A spreadsheet link cleanup could delete a table placed soon after the link, along with the text before it.
- Settings and pointer files starting with a byte-order mark could not be read, and the next save could overwrite existing settings.
- The table right-click menu now acts on the clicked cell. Remove Selected Row could crash when the caret was outside the table.
- Image Properties now keeps alt text and title, and Fit to editor width stays inside a table cell.
- Text typed at the end of a link no longer extends the link.


## 1.5.0 - 2026-05-25
### Changed
- Planning Register presets now save and restore composite table layouts more faithfully, including nested companion cost tables.
- Preset table insertion now preserves saved HTML directly and improves inserted table sizing and border normalization.

### Fixed
- Removed the dead "New Planning Register" placeholder from the Insert Planning Register chooser so only real saved presets are offered.
- Removed the unused View > Rich Text Tools menu placeholder from the main window.
- Hardened Planning Register and cost-list table detection so existing saved layouts rebuild and style correctly.


## Unreleased - 2025-10-11
### Removed
- Formula (table cell) feature fully rolled back; all related menu actions, persistence hooks, and parsing dependencies removed (beautifulsoup4 pruned).

### Added / Restored
- Reintroduced database initialization helpers (ensure_database_initialized, migrate_database_if_needed, create_new_database_file) after rollback cleanup.

- 22c1c21 Changed over to 2-Column dis
- 2b7c280 General HouseKeeping
- a060ff4 Added copy and paste functionality to the rich text editor toolbar.
- e235d3f Docs: add Features, paste modes, list indenting, and links to README
- 71a186c Rich text improvements: list nesting with Tab/Shift+Tab; Classic numbering persistence across reloads; paste modes with default mode; background/size normalization; single-click links; clearer selection highlight; add .gitignore; media store scaffolding
- 9c00a5d Add VS Code tasks/launch and recommendations; README VS Code usage
- 2b0260c Add dev environment: requirements-dev.txt and README with setup; clarify runtime requirements
- 72fe282 Add requirements.txt with PyQt5 runtime deps; document optional dev tools
- 61fd538 Fixed a bunch of bugs and behaviors regarding the tree behavior in the left and right panels.
- 01888c1 Bunches of changes today.
- b5b373a Initial Commit