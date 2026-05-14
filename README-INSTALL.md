# NoteBook — Installation Guide

## What's in the zip file

The zip file contains a self-contained folder with everything NoteBook needs to run. There is no installer — just copy the folder to wherever you want it and run `NoteBook.exe`.

```
NoteBook-v1.0.2.zip
└── NoteBook_Release\
    ├── NoteBook.exe              ← the application
    ├── add_to_start_menu.cmd     ← optional: adds a Start Menu shortcut
    ├── clear_icon_cache.cmd      ← optional: fixes missing/stale icon
    ├── README.md
    └── _internal\                ← supporting files (do not move or delete)
```

---

## Step 1 — Extract the zip

Right-click `NoteBook-v1.0.2.zip` and choose **Extract All**, then choose a destination folder.

**Recommended location:**

```
C:\Users\{username}\AppData\Local\NoteBook\
```

You can open this path quickly by typing `%LOCALAPPDATA%` in the Windows Explorer address bar.

However, you can place the folder **anywhere on your computer** — a desktop folder, a `C:\Apps\` directory, a USB drive, etc. NoteBook is fully portable and does not write anything inside its own folder after installation.

---

## Step 2 — Run NoteBook

Double-click `NoteBook.exe` inside the extracted folder. No further setup is required.

On first launch, NoteBook will ask you to create or open a database file. See the **Choosing Storage Locations** section in the in-app Help menu for advice on where to keep your database and backups.

---

## Step 3 (Optional) — Add to the Windows Start Menu

To make NoteBook searchable from the Start Menu and easier to pin to the taskbar:

1. Open the extracted `NoteBook_Release` folder
2. Double-click **`add_to_start_menu.cmd`**
3. Follow the on-screen prompts

This script creates a shortcut in your personal Start Menu (`%APPDATA%\Microsoft\Windows\Start Menu\Programs\`). The program itself stays wherever you put it — running this script does **not** move any files.

You only need to run this script once. If you later move the NoteBook folder to a different location, run it again to update the shortcut.

---

## Settings file

NoteBook stores your preferences (database path, backup path, theme, etc.) in a settings file at:

```
C:\Users\{username}\AppData\Local\NoteBook\settings.json
```

This file is created automatically the first time you run the application. You do not need to create or edit it manually — all settings are managed through **Tools → Settings** inside the app.

> **Note:** The `settings.json` file is machine-specific. Do not copy it between computers; settings will be re-created automatically on any new machine.

---

## Uninstalling

NoteBook does not use the Windows registry and does not have an uninstaller entry in "Add or Remove Programs".

To uninstall:

1. Delete the NoteBook folder
2. If you added a Start Menu shortcut, delete it from **Start Menu → All Apps** (right-click → Unpin / Delete)
3. Optionally delete the settings file at `%LOCALAPPDATA%\NoteBook\settings.json`

Your database and backup files are **not** deleted — those are wherever you chose to save them and are yours to keep or remove as you wish.

---

## System requirements

- Windows 10 or Windows 11 (64-bit)
- No Python installation required
- No additional software required
