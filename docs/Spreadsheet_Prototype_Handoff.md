# Spreadsheet Prototype Handoff

Last updated: 2026-05-14

## Current Objective
Prototype spreadsheet capability in the NoteBook fork while keeping the original project safe.

## Safety Setup (Completed)
- Stable project: c:\Python Projects\NoteBook_clean
- Working fork: c:\Python Projects\NoteBook_spreadsheet_fork
- Backup tag on GitHub: backup/pre-spreadsheet
- Prototype branch on GitHub: feature/spreadsheet-prototype

## Architecture Direction
- Store spreadsheet files in the media folder, not as DB blobs.
- Keep DB as metadata manifest only (page-to-file mapping).
- Keep templates in a dedicated templates folder.

## Initial Build Plan
1. Add DB migration for spreadsheet metadata table.
2. Add media subfolders:
   - media/spreadsheets/
   - media/spreadsheet_templates/
3. Add menu actions for:
   - Insert Spreadsheet
   - Insert From Spreadsheet Template
4. Implement basic flow:
   - Create/open spreadsheet file
   - Link file to page in DB
   - Open externally for troubleshooting

## Next Session Start Checklist
- Open folder: c:\Python Projects\NoteBook_spreadsheet_fork
- Confirm branch: feature/spreadsheet-prototype
- Start with migration + menu scaffolding (no formula engine yet)
