Here's a discussion on what we need to do next with undo/Restore:

We've covered undo/redo with the text inside a page, but we haven't provided for the eventuality that a page, section, or binder will be incorrectly deleted. This will require some database changes.
Here's what I want to happen:

1 - When a Page, Section, or Binder is deleted, I want the item to simply be marked as deleted in a new field in the database.
2 - In normal viewing, the item will not be displayed
3 - When we right click in the left column, the context menu will have an option to "Show Deleted Items".
4 - Likewise, the main menu should have an item under File to "Show Deleted Items".
5 - When selected, all deleted items will be displayed in the left panel, except they will be greyed out, but still selectable.
6 - We should be able to click on and examine the deleted item just like normal, but not necessaroly edit.
7 - A right click on the deleted item reveals an option to restore it, or delete it permanently.
8 - There'll also be an item in the Files part of the main menu to permanently delete all items marked for deletion.
9 - Items marked for deletion will not time out, but remain in the deleted status until restored or permanently deleted.
10 - We'll also need to impliment a database upgrade when the program encounters an older database. This would give the user the option to upgrade "Yes/No", or load another database. That gives someone who's a bit nervous an opportunity to exit the program and make manual backups of the database before proceeding.

## Implementation Notes

**Cascade Behavior:**
- When a Section is deleted, all its Pages are also marked deleted
- When a Binder is deleted, all its Sections (and their Pages) are also marked deleted
- On restore, children are auto-restored as well
- No orphans allowed - permanent deletion of a parent permanently deletes all children

**Schema Change:**
- Add `deleted_at TIMESTAMP NULL` column to pages, sections, and binders tables
- NULL = active, non-NULL = soft-deleted (timestamp records when deletion occurred)

**"Show Deleted" Toggle:**
- Persist this setting between sessions (store in settings)

**Read-Only Deleted Items:**
- Deleted items are viewable but not editable
- Prevents accidental work on items pending deletion

**Confirmation Dialogs:**
- Confirm before permanent deletion of individual items
- Confirm before "Empty All Deleted Items" operation

**Empty State:**
- Display notice when "Show Deleted" is enabled but no deleted items exist