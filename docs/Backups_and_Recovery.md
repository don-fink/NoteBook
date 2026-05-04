# Backups and Recovery

NoteBook automatically protects your notes by creating backup copies every time you close the application. This page explains how backups work and how to manage them.

## Automatic Backups on Exit

Every time you close NoteBook cleanly, it creates a timestamped backup of the current database. Backups are saved as `.bundle` files.

A `.bundle` file is a ZIP archive (you can open it with 7-Zip or WinRAR) that contains:

- The database file (e.g. `notes.db`)
- A `media/` folder if the database has embedded images or attachments

Backups are named like: `notes-20260504-143022.bundle`

## Where Backups Are Stored

The backup folder is set in **Tools → Settings**. The default is a folder named `NoteBook_Backups` in your Documents folder, but you can change it to any local folder you prefer.

## Manual Backup

You can create a backup at any time without closing the app:

- Go to **File → Backup Database**

This creates a timestamped `.bundle` file in your backup folder immediately.

## Backup Rotation

NoteBook keeps a limited number of backup files per database to avoid filling your disk. The default limit is 6 backups.

The count is tracked **per database name**. If you have multiple databases (e.g. `work.db` and `personal.db`), each one keeps its own independent set of backups.

You can change the backup limit in **Tools → Settings**.

## Recovering from a Backup

For step-by-step instructions on opening and restoring a `.bundle` file, see:

**Help → Restore Backup Bundles**

## Keeping Your Data Safe — Recommendations

**Avoid cloud-synced folders for the live database.**
Services like OneDrive, Dropbox, and iCloud Drive can lock a file at the exact moment NoteBook needs to write to it. This can corrupt a write transaction and cause data loss. Keep the active `.db` file on a local drive.

**Use the backup folder for cloud sync instead.**
Set your NoteBook backup folder to a location that *is* synced to the cloud. That way, finished, closed `.bundle` files are safely uploaded to the cloud — without risking the live database being locked mid-write.

**Periodically check your backup folder.**
Glance at the backup folder occasionally to confirm recent `.bundle` files are being created. If the folder is empty or the files are old, check your backup path in **Tools → Settings**.

**Before major changes, make a manual backup.**
Before doing a large reorganisation — moving many pages, deleting sections, or restructuring binders — use **File → Backup Database** first.
