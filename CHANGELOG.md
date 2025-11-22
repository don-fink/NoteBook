

## Unreleased - 2025-10-11
### Removed
- Formula (table cell) feature fully rolled back; all related menu actions, persistence hooks, and parsing dependencies removed (beautifulsoup4 pruned).

### Added / Restored
- Reintroduced database initialization helpers (ensure_database_initialized, migrate_database_if_needed, create_new_database_file) after rollback cleanup.

- 22c1c21 Changed over to 2-Column dis
- 2b7c280 General HouseKeeping
- a060ff4 Added copy and paste functionality to the rich text editor toolbar.
- e235d3f Docs: add Features, paste modes, list indenting, and links to README
- 71a186c Rich text improvements: list nesting with Tab/Shift+Tab; Classic numbering persistence across reloads; paste modes with default mode; background/size normalization; single-click links; clearer selection highlight; add .gitignore; media store scaffolding
- 9c00a5d Add VS Code tasks/launch and recommendations; README VS Code usage
- 2b0260c Add dev environment: requirements-dev.txt and README with setup; clarify runtime requirements
- 72fe282 Add requirements.txt with PyQt5 runtime deps; document optional dev tools
- 61fd538 Fixed a bunch of bugs and behaviors regarding the tree behavior in the left and right panels.
- 01888c1 Bunches of changes today.
- b5b373a Initial Commit