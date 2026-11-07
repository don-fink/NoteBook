# NoteBook Future Features

> **Note (February 2026):** Most features documented in earlier versions of this file have been implemented. The sections below reflect remaining or newly proposed work.

---

## Planned: Currency Column Presets Refactor

### Problem
The current implementation for adding currency columns and storing presets is:
- Convoluted and difficult to understand
- Hard to manage and extend
- Brittle and prone to bugs

### Goal
Redesign the currency column feature to be:
- Straightforward and easy to understand
- Simple to configure and use
- More reliable and maintainable

### Status
Under consideration. Will be revisited after additional real-world usage to identify specific pain points.

---

## Planned: Main Menu Cleanup

### Problem
The main menu contains several unused and unnecessary entries, including:
- 2-column and 3-column display options (no longer applicable)
- Other legacy menu items that clutter the interface

### Goal
- Remove unused menu entries
- Reorganize menu structure for clarity
- Group related actions logically
- Ensure all menu items have working functionality

### Status
Planned for a future cleanup pass.

---

## Planned: Theme System Expansion

### Current State
The app has basic theming with `default.qss` and `high-contrast.qss` stylesheets.

### Goal
Create a proper theme system with:
- **Light theme** — Clean, modern light appearance (refine current default)
- **Dark theme** — A polished, easy-on-the-eyes dark mode
- **Theme selector** — Add theme switching to the main menu (View or Settings menu)

### Considerations
- Ensure all UI elements are properly styled in each theme
- Test readability of text, tables, and editor content
- Persist theme selection in settings

### Status
Planned. Dark theme is the priority addition.

---

## Proposed: Search (discussion stage, 2026-11-06)

### The user's spec
- A search symbol in the upper right of the window, possibly near the Edit menu.
- Every search starts from where the cursor is in the left panel:
  - a binder is selected → search only inside that binder;
  - the whole database is selected → search the entire database.
- When pages match or nearly match what's typed in the search box, show that page
  and highlight it in the left panel, so the user can step through the results one at a time.
- Open to suggestions.

### Facts from the code (for whoever builds it)
- Left panel is the `QTreeWidget` `notebookName`: Binder (top level) → Section → Page → subpages.
  Item data: 1000 = id, 1001 = kind ("section"/"page"), 1002 = section id, 1003 = deleted flag.
- There is no "whole database" node in the tree today; binders are the top-level items.
- Pages live in `pages` (`title`, `content_html`, `section_id`, `parent_page_id`, `deleted_at`).
  No tags column. Spreadsheets are separate files listed in `spreadsheets`.
- No full-text index yet. The bundled SQLite (3.50) supports FTS5.
- Existing helpers to reuse: `two_pane_core.select_left_tree_page` (expands binder, section
  and parent pages, then selects), `ensure_left_tree_sections`, `load_page`.
- Menus: File, Insert, Edit (Undo/Redo/paste modes), Help. Ctrl+F is not used yet.

### Suggestions to discuss
- Search box in the menu bar's top-right corner (`QMenuBar.setCornerWidget`), Ctrl+F to focus it,
  plus Edit → Find… entry.
- Scope from the left-panel selection (binder / section / page and its subpages), shown in the
  box's placeholder ("Search in Recipes"); a way to pick "Whole database" since the tree has no root node.
- Search page titles and body text (HTML stripped); tables' text comes along with the body.
  Skip deleted pages unless Show Deleted Items is on.
- "Nearly match": typo-tolerant by default, with an Exact toggle (and maybe Match case / Whole word).
- Results: "N of M" counter, Enter / Shift+Enter (or F3 / Shift+F3) for next / previous; each hit
  expands the tree, selects the page, opens it and highlights the words in the page; Esc clears.
- Optional "Show matches only" filter that hides non-matching pages in the tree.
- For large databases: an FTS5 index over title + plain text, kept up to date on page save.

### Decisions (2026-11-06)
- Whole database: a scope choice on the search box (current selection / Whole database); the tree is unchanged.
- Matching: typo-tolerant by default with an Exact switch; exact matches are listed before near matches.
- Keys: Enter / Shift+Enter in the box and F3 / Shift+F3 anywhere step through results.
- "Show matches only" filter: later, after the first version.

### Status
First version built 2026-11-06 (`search_engine.py`, `ui_search.py`). Still to come: the "Show matches only" filter, and an FTS5 index if large databases get slow.

---

## Proposed: Move or Copy Sections (and maybe Pages) to Another Binder (2026-11-07)

### The user's idea
- Move a section, or a copy of it, into a different binder. Sections first; pages possibly later.
- A "Move or Copy" kind of command rather than (or as well as) drag-and-drop.

### Facts from the code (for whoever builds it)
- Today the left panel only reorders: dragging and Ctrl+Up/Down keep a section in its binder and a
  page among its siblings (`LeftTreeDnDFilter` and `move_item_by_key` in `left_tree.py`).
- `sections.notebook_id` holds a section's binder; moving one means changing it and giving it an
  `order_index` at the end of the new binder (`db_sections.set_sections_order`). Its pages follow,
  since they hang off `section_id`.
- For pages, `db_pages.set_pages_parent_and_order` can already move pages to another section or
  parent page, but nothing calls it yet. Subpages would need to come along.
- A copy needs new rows for the section, every page and subpage (keeping parent links), and the
  page media: images and attachments live under the database's `.media` folder and are tracked by
  `media_store`, so copied pages must not share media that a later delete would remove.
- Deleted (`deleted_at`) sections and pages: decide whether they move/copy along or stay behind.

### Suggestions to discuss
- Right-click a section → **Move or Copy…**: a window listing the binders, a "Make a copy" checkbox,
  and where to put it (end of the binder, or before a chosen section).
- Optionally allow dragging a section onto another binder to move it, with Ctrl held to copy.

### Status
Proposed; not started.

---

## Future: Full Versioning System

### Overview
Track all changes with timestamps for true undo/redo across the entire application.

### Approach Options

| Approach | Complexity | Description |
|----------|------------|-------------|
| **Page content versioning** | Medium | Store page HTML snapshots on each save with timestamp. Allow viewing/restoring previous versions. |
| **Command pattern** | High | Application-level undo stack recording each operation and its inverse. Complex to implement correctly. |
| **Event sourcing** | Very High | Store all changes as events; reconstruct state by replaying. Overkill for this app. |

### Recommended: Page Content Versioning

**Database Changes:**
```sql
CREATE TABLE page_versions (
    id INTEGER PRIMARY KEY,
    page_id INTEGER NOT NULL,
    content TEXT,
    saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (page_id) REFERENCES pages(id) ON DELETE CASCADE
);
```

**Features:**
- "Version History" panel showing save timestamps
- Click to preview any version
- "Restore" button to revert to selected version
- Auto-cleanup: keep last N versions or versions from last M days

**UI:**
- Right-click page → "View History"
- Or toolbar button when editing a page
- Side panel showing version list with timestamps
- Diff view (optional, more complex)

---

## Session Notes

### 2026-02-22: Font Persistence Fix
- **Problem**: Font changes weren't persisting after save/reload
- **Root cause**: `mergeCharFormat()` was appending fonts to existing font stacks instead of replacing them
- **Solution**: Changed `_apply_font_family()` to iterate through each character and use `setCharFormat()` which fully replaces the format
- **Files changed**: `ui_richtext.py`

### 2026-02-22: Undo/Redo Added
- Added Edit menu items: Undo (Ctrl+Z), Redo (Ctrl+Y)
- Wired to QTextEdit's built-in undo/redo stack
- Menu items auto-enable/disable based on availability
- **Files changed**: `main_window_2_column.ui`, `main.py`

### 2026-02-22: HR Persistence Fix (previous session)
- Fixed horizontal rule CSS properties being stripped by sanitizer
- Sanitizer now preserves expanded border properties on `<hr>` tags
