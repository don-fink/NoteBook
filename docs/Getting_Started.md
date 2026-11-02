# Getting Started with NoteBook

NoteBook is a lightweight desktop note-taking application that organizes your notes into **Binders**, **Sections**, and **Pages** — similar to a physical ring binder.

## The Interface

The window is divided into two panels:

- **Left panel** — the navigation tree, where you browse and manage your Binders, Sections, and Pages
- **Right panel** — the page editor, where you read and write your notes

## Opening a Database

Your notes are stored in a single database file (`.db`). When NoteBook starts for the first time, it asks whether to create a new database or open an existing one, and where to keep it. It asks again if it can't find the database it used last time.

To open a different database:

- Go to **File → Open** and browse to a `.db` file

To create a new, empty database:

- Go to **File → New Database**

## Your First Notes

1. In the left panel you will see a default Binder. Right-click it to rename it to something meaningful.
2. Right-click the Binder to add a **Section** — think of these as tabbed dividers inside the binder.
3. Right-click a Section to add a **Page**.
4. Click a Page name to select it and begin typing in the right panel.

## Auto-Save

NoteBook saves your work automatically as you type. You do not need to press Save manually. **Ctrl+S** is available if you prefer to save explicitly, but nothing will be lost if you just close the window.

## Settings

Access application settings from **Tools → Settings**. From there you can configure:

- Default database location
- Backup folder and number of backups to keep
- Theme
- Font defaults

## Switching Themes

Go to **Tools → Settings → Themes** and choose a theme. NoteBook comes with **Default** and **High Contrast**. The change takes effect immediately without restarting.

A theme is a Qt stylesheet (`.qss`) file. To add your own, put a `.qss` file in a `themes` folder next to your `settings.json` (the Themes tab shows the exact folder), then reopen Settings. The file name becomes the theme's name: `dark-slate.qss` appears as **Dark Slate**. A file named like a theme that comes with NoteBook, such as `default.qss`, replaces that theme. Copying one of the files from NoteBook's own `themes` folder is the easiest way to start.

Menus take their look from the `QMenu` rules in the theme. A `border-radius` there gives them rounded corners.

## Moving On

Once you are comfortable with the basics, explore these other Help topics:

- **Organizing Your Notes** — more detail on binders, sections, and pages
- **Rich Text Editing** — formatting, lists, paste modes, and images
- **Tables and Currency Columns** — structured data with auto-totals
- **Backups and Recovery** — how NoteBook protects your data
