# Choosing Storage Locations

NoteBook uses three separate file locations: one for the **main database**, one for **backups**, and one for the **settings file**. Choosing the right location for each is important for keeping your data safe and reliable.

---

## Main Database

The main database (`.db` file) is a live SQLite file that NoteBook reads from and writes to constantly while the application is running.

**Do NOT store the main database in a cloud-synced folder.**

Cloud sync services such as OneDrive, Dropbox, Google Drive, and iCloud Drive periodically lock files to upload them. If the sync service locks the database at the exact moment NoteBook needs to write to it, the write transaction can fail or become corrupted. SQLite databases are particularly vulnerable to this because a single interrupted write can leave the file in an inconsistent state.

**Store the main database on a local drive instead.**

A plain local folder — one that is not watched by any sync service — is the safest location. The folder does not need to already exist; NoteBook will use whatever path you specify in **Tools → Settings**.

**Example (Windows):**

```
C:\Notebooks\notes.db
```

The folder `C:\Notebooks` is local, not synced to any cloud service, and is therefore safe for the live database.

---

## Backup Location

Backups are saved as finished, closed `.bundle` files. Because NoteBook is not actively writing to these files after they are created, cloud sync services can safely upload them.

**The backup folder is the ideal place to use cloud sync.**

By pointing the backup folder at a cloud-synced directory, your backups are automatically uploaded offsite without any risk to the live database.

**Example (Windows):**

```
C:\NotebooksBU\
```

If `C:\NotebooksBU` is synced to OneDrive, Dropbox, or another service, finished backup bundles will be uploaded automatically after each session.

**Also set a generous backup count.**

In **Tools → Settings**, increase the number of backups to keep. The default is 6, but keeping 10–20 backups costs very little disk space and gives you a longer history to recover from if something goes wrong.

---

## Settings File

The settings file stores your preferences (database path, backup path, theme, font, etc.). It can be placed anywhere on the computer.

By default, the settings file is created in the same directory as the NoteBook program files. On a typical Windows installation this is:

```
C:\Users\{username}\AppData\Local\NoteBook\
```

You do not normally need to move the settings file. If you reinstall NoteBook or move it to a different folder, NoteBook will create a new settings file with default values, and you can re-enter your preferred paths in **Tools → Settings**.

---

## Summary

| File | Recommended location | Cloud sync? |
|---|---|---|
| Main database (`.db`) | Local drive only | **No** — risk of corruption |
| Backup bundles (`.bundle`) | Cloud-synced folder | **Yes** — safe and recommended |
| Settings file | Program folder (default) | Not required |
