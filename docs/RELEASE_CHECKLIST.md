# NoteBook Release Checklist

This checklist covers committing the current source, building a Windows
distribution, and preparing it for local or external testing.

## 1. Confirm the working state

Open PowerShell in the repository root and inspect the branch and worktree:

```powershell
git status
git branch --show-current
git log -5 --oneline --decorate
```

Review all source changes before staging them:

```powershell
git diff
git diff --check
```

Do not commit machine-specific files, databases, settings, logs, virtual
environments, or generated build output. The release folder and ZIP are
ignored by Git.

## 2. Validate the source

Use the project virtual environment:

```powershell
& .\.venv\Scripts\python.exe -m compileall -q -f .
& .\.venv\Scripts\python.exe -m pytest -q
```

There may not be automated tests for every application feature. In that case,
manually verify that the source application starts and that the main workflows
still work:

- Open or create a notebook database.
- Create, rename, reorder, and delete notebooks, sections, and pages.
- Edit and save rich text.
- Insert and reopen images, attachments, and spreadsheets.
- Test spreadsheet CSV and XLSX operations if those features are included in
  the release.
- Confirm backups and database migrations work on a copy of a database.

## 3. Commit the source changes

Stage only the intended source, documentation, and configuration files:

```powershell
git add <intended-files>
git diff --cached --check
git diff --cached --stat
git status
```

Commit after reviewing the staged diff:

```powershell
git commit -m "Describe the release changes"
```

Push the branch to GitHub:

```powershell
git push origin <branch-name>
```

Do not commit `NoteBook_Release`, `dist`, `build`, `.venv`, local settings,
databases, logs, or ZIP files unless a deliberate repository policy later
requires one of them.

## 4. Build the Windows distribution

Run the repository packaging script from PowerShell:

```powershell
& .\create_release_simple.ps1
```

The script uses `.venv\Scripts\python.exe` when available and runs PyInstaller
with the configuration in [notebook.spec](../notebook.spec). It creates the
one-folder distribution in `NoteBook_Release`.

Confirm that the release contains at least:

```text
NoteBook_Release\
├── NoteBook.exe
├── README-INSTALL.md
├── README.txt
├── add_to_start_menu.cmd
├── clear_icon_cache.cmd
└── _internal\
```

Do not move or delete `_internal` from the distribution folder.

## 5. Test the packaged application

Test the executable from the completed distribution folder, not from the ZIP
file:

```text
NoteBook_Release\NoteBook.exe
```

Perform the same core workflow checks used for the source application. Also
confirm that:

- No Python installation is required.
- No Microsoft Office installation is required for the supported spreadsheet
  features.
- The application creates per-user settings under
  `%LOCALAPPDATA%\NoteBook\`.
- A database can be opened from a user-selected location.
- The release folder does not contain personal `settings.json`, `settings.loc`,
  database files, or private logs.

If the application fails during startup, inspect `crash.log` and
`native_crash.log` in the development environment. Do not distribute those
files unless they are intentionally included for diagnostics.

## 6. Create the test ZIP

After the folder has been tested, create a ZIP for transfer to another
computer:

```powershell
if (Test-Path .\NoteBook_Release.zip) {
    Remove-Item .\NoteBook_Release.zip -Force
}
Compress-Archive -Path .\NoteBook_Release\* `
    -DestinationPath .\NoteBook_Release.zip `
    -CompressionLevel Optimal
```

Extract the ZIP completely before testing it on another computer. Keep the
database and backups separate from the application ZIP.

See [README-INSTALL.md](../README-INSTALL.md) for end-user installation and
uninstallation instructions.

## 7. Future project rename

The migration folder and repository should be renamed only after a known-good
release has been committed and tested.

Planned sequence:

1. Commit and push the final migration-era source.
2. Optionally create a GitHub tag or release.
3. Rename the GitHub repository.
4. Update the local `origin` URL:

   ```powershell
   git remote -v
   git remote set-url origin https://github.com/<owner>/<new-repository>.git
   ```

5. Close VS Code and rename the local folder from
   `NoteBook_pyqt6_migration` to `NoteBook`.
6. Reopen the renamed folder.
7. Recreate `.venv`, because its configuration may contain the old absolute
   path:

   ```powershell
   Remove-Item .venv -Recurse
   python -m venv .venv
   & .\.venv\Scripts\python.exe -m pip install --upgrade pip
   & .\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
   ```

8. Search documentation, VS Code configuration, scripts, and packaging files
   for the old repository and folder names.
9. Re-run the source checks, build a new distribution, and test the renamed
   project.

The application product name is already `NoteBook`; the rename primarily
removes the temporary migration name from the repository and local workspace.
