# Troubleshooting

## App Crashes Intermittently

If NoteBook closes unexpectedly, run it in Safe Mode to disable some heavier overlay processes:

```powershell
$env:NOTEBOOK_SAFE_MODE='1'; python main.py
```

Check these log files for clues (look in the same folder as the app):

- `crash.log` — Python exceptions and Qt error messages
- `native_crash.log` — Low-level crash traces

## Editor Stays Read-Only

The editor is read-only when no page is selected in the left panel. Click a page name to enable editing.

If you have a page selected and the editor is still read-only, try clicking a different page and then clicking back. If the problem persists, check that the database file itself is not marked read-only in Windows (right-click the `.db` file → Properties → uncheck Read-only).

## Panel Sizes Do Not Restore on Startup

NoteBook saves the left/right panel split position when you exit. If the sizes reset every time you launch:

1. Make sure `settings.json` is writable. It lives at `%LOCALAPPDATA%\NoteBook\settings.json`.
2. Exit the app cleanly at least once — do not force-close or kill the process via Task Manager.

## UI Fails to Load

If you see an error mentioning `UnsupportedPropertyError for 'list'` or a similar Qt message, the `.ui` layout file may have an incompatible property. Make sure you are using the `main_window_2_column.ui` file that came with this version of NoteBook.

## Fonts or Formatting Look Wrong After Paste

If pasted text has unexpected fonts, colours, or spacing, the source content carried heavy formatting. Try one of these:

- Press **Ctrl+Shift+V** to paste as plain text (removes all formatting).
- Use **Edit → Paste Clean Formatting** to keep links and images but strip styling.
- Use **Edit → Default Paste Mode → Text Only** to make plain-text paste the default going forward.

## Database Won't Open

If a `.db` file fails to open:

- Make sure no other program (including another NoteBook instance) has the file open.
- If the file is in a cloud-synced folder (OneDrive, Dropbox, etc.), wait for the sync to finish and try again. Better yet, move the database to a local folder — see **Help → Backups and Recovery** for why this matters.
- If the file was received from another computer, make sure it is fully downloaded and not just a cloud placeholder.

## Data Appears Missing After Restore

When restoring from a `.bundle` backup, images and attachments will be missing if the `media/` folder was not restored alongside the database:

- Open the `.bundle` file as a ZIP archive.
- If there is a `media/` folder inside, extract it next to the restored `.db` file.
- Rename the media folder to match the database filename. For example, if your database is `notes.db`, the media folder must be named `notes.db.media`.

See **Help → Restore Backup Bundles** for the full step-by-step process.

## Order Numbers Look Wrong in the Tree

If items in the navigation tree appear out of order after many edits, deletions, or imports, use **Tools → Normalize Page Order**. This reassigns clean sequential order numbers to all items. It only affects the sort order — no content is changed.
