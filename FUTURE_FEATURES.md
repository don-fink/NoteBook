# NoteBook Future Features

## Next Up: Soft Delete / Trash System

### Problem
Currently, deleting pages, sections, or binders is permanent and immediate. There's no undo for structural operations (only text editing within pages has undo/redo via Qt's built-in QTextEdit stack).

### Proposed Solution: Soft Delete with Trash View

**Database Changes:**
- Add `deleted_at TIMESTAMP` column to `pages`, `sections`, and `notebooks` tables
- "Delete" sets `deleted_at = datetime.now()` instead of removing rows
- All queries filter by `WHERE deleted_at IS NULL` by default

**UI Changes:**
- Add "Trash" view accessible from Tools menu or left panel
- Trash view shows deleted items grouped by type (Pages / Sections / Binders)
- Right-click context menu: "Restore" or "Delete Permanently"
- Optional: "Empty Trash" action to permanently remove all deleted items

**Benefits:**
- Simple to implement (~100-150 lines)
- Users can recover accidentally deleted content
- No complex undo stack required
- Deleted items can be auto-purged after N days (configurable)

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
