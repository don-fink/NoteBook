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
