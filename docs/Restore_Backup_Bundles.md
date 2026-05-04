# Restore From Backup Bundles

NoteBook exit backups are saved as `.bundle` files. A `.bundle` is a ZIP archive with a different file extension.

Each backup bundle normally contains:

- The SQLite database file at the root of the archive, such as `notes.db`
- A `media/` folder if that database has stored images or attachments

## Important: Backup Rotation Is Per Database Name

Backup retention is tracked by database filename.

If your backup setting says to keep the last 6 backups, that means:

- `notes.db` keeps its own last 6 `.bundle` files
- `project_a.db` keeps its own last 6 `.bundle` files
- `archive_copy.db` keeps its own last 6 `.bundle` files

So if you switch between different databases, you can have more than 6 total backup files in the backup folder. The limit applies separately to each database name.

This can be useful during recovery. If you restore a backup under a new database name, future exit backups for that restored file will rotate under the new name instead of pruning backups for the original database name.

## How To Restore A Backup Bundle

1. Close NoteBook before replacing or testing database files.
2. Go to the folder where your `.bundle` backups are stored.
3. Pick the backup you want to inspect or restore.
4. Copy that `.bundle` file to a local folder first if it is stored in iCloud and may not be fully downloaded.
5. Open the `.bundle` with 7-Zip, WinRAR, or by renaming the file from `.bundle` to `.zip`.
6. Extract the contents to a working folder.

After extraction, you should usually see:

- A database file such as `notes.db`
- Possibly a `media` folder

## Putting The Restored Files Back

If the extracted backup contains only the `.db` file, you can open that database directly in NoteBook.

If the extracted backup also contains a `media` folder, place it beside the restored database and rename the folder to match the database filename plus `.media`.

Example:

- Database file: `notes.db`
- Media folder expected by NoteBook: `notes.db.media`

If you rename the restored database, rename the media folder to match.

Example:

- Restored database renamed to `notes_recovered.db`
- Matching media folder must be `notes_recovered.db.media`

## Open The Restored Database In NoteBook

1. Start NoteBook.
2. Use the main menu option to open a database.
3. Browse to the restored `.db` file and open it.
4. Check that the pages, sections, and any images or attachments look correct.

## Recommended Recovery Workflow

When you are not sure which backup is good:

1. Extract one backup at a time.
2. Rename each restored `.db` to something descriptive, such as `notes_recovered_2026_05_01.db`.
3. If there is a media folder, rename it to match the restored database name plus `.media`.
4. Open that database in NoteBook and inspect it.
5. Keep the restored file under its own name until you are sure it is the version you want.

This avoids accidentally mixing a recovery copy into the backup rotation of your original working database.

## Notes

- There is currently no built-in menu command to restore a `.bundle` automatically.
- Restoring from a `.bundle` is manual by design.
- A `.bundle` is safe to inspect with normal ZIP tools.