# NoteBook Maintainability Refactor Plan

Date: 2026-03-07
Status: Planning document only (no runtime behavior changes yet)

## Why This Document
The app is currently functional and valuable in day-to-day use. This plan is intended to preserve that stability while improving maintainability before implementing additional formatting/paste fixes.

## Current Findings (Architecture Snapshot)

### What is already good
- Feature separation exists in several places:
  - Rich text behavior is mostly in `ui_richtext.py`.
  - Two-pane/page loading is in `two_pane_core.py` and `page_editor.py`.
  - Settings persistence is centralized in `settings_manager.py`.
  - DB responsibilities are split across `db_access.py`, `db_pages.py`, and `db_sections.py`.
- The application is operational and user-validated in real usage.

### Main maintainability risks
- Two oversized modules are acting as control hubs:
  - `ui_richtext.py` is very large (about 5300+ lines, many responsibilities).
  - `main.py` is very large (about 4300+ lines, extensive wiring + behavior).
- Orchestration logic is spread between places where policy and implementation are mixed:
  - Paste mode logic is implemented in `ui_richtext.py`.
  - Paste mode menu wiring, shortcuts, and settings integration are heavily wired in `main.py`.
- `main.py` has high coupling to settings and multiple feature areas, making regressions more likely when touching unrelated behavior.

## Formatting/Paste-Specific Observations

From user testing and code review, the likely issue pattern is:
- Incoming browser HTML/CSS contains nested inline style complexity.
- Sanitized modes reduce some formatting but may still produce content that does not respond to toolbar font/size commands consistently.
- Toolbar state and effective text formatting can become desynchronized in mixed-format fragments.

### Plausible root causes (to verify when implementation time is available)
- Paste modes are not all passing through one canonical normalization pipeline.
- Mode behavior is not defined tightly enough (what to keep/drop in each mode).
- Font/size application may not fully override style precedence in all pasted fragments.
- Selection format reporting and command application may behave differently across mixed inline runs.

## Recommended Refactor Direction (Low-Risk First)

### Goal
Make feature ownership explicit, reduce coupling, and create one reliable pipeline for paste + formatting behavior.

### Phase 0: Safety and Baseline (no behavior change)
1. Add this architecture plan (done).
2. Add lightweight logging hooks around paste path entry points (mode selected, source hasHtml/hasText, inserted char count).
3. Capture 3-5 reproducible paste samples in a test note/database for regression checks.

### Phase 1: Move Policy Out of `main.py` (minimal behavior change)
1. Introduce a dedicated module, for example: `services/paste_pipeline.py`.
2. Move mode routing into one public function, e.g. `run_paste_mode(text_edit, mode)`.
3. Update `main.py` so it only wires actions/shortcuts to that service.

Expected win: one place to reason about paste behavior.

### Phase 2: Define an Explicit Paste Mode Contract
Document and enforce exact behavior for each mode:
- `rich`: preserve rich content, allow safe structure and expected styles.
- `text-only`: insert plain text only, no inherited hidden rich style.
- `match-style`: preserve structure (paragraphs/lists/links/images where allowed), normalize font family/size/background to current context.
- `clean`: preserve readable structure and links, drop most styling/classes and nonessential inline formatting.

Expected win: no ambiguity in expected outcomes; easier debugging and testing.

### Phase 3: Split `ui_richtext.py` by responsibility
Suggested decomposition:
- `richtext_toolbar.py` (toolbar creation + sync)
- `richtext_paste.py` (paste modes + normalization)
- `richtext_sanitize.py` (HTML cleanup + storage sanitization)
- `richtext_tables.py` (table/currency/formula behavior)
- `richtext_images.py` (image insertion/resizing/context actions)

Keep `ui_richtext.py` as a compatibility facade during migration to avoid breaking imports.

Expected win: smaller files, clearer ownership, safer future edits.

### Phase 4: Regression Guardrails
1. Add focused tests for paste mode outputs (or deterministic fixture checks if full UI tests are too heavy).
2. Add a manual verification checklist for key editor workflows:
   - Paste in all 4 modes
   - Change font/size on pasted selection
   - Move caret and verify toolbar reflects effective style
   - Save/reload page and verify formatting persistence

Expected win: confidence to iterate without breaking core editing behavior.

## Suggested First Work Session (2-4 Hours)
1. Implement Phase 1 only:
   - Create `services/paste_pipeline.py`.
   - Move mode routing into service.
   - Leave internals of paste functions untouched for now.
2. Add a tiny architecture map doc (`docs/ARCHITECTURE.md`) listing feature owners.
3. Confirm app behavior is unchanged with a quick smoke test.

This first session gives immediate maintainability benefits with low functional risk.

## Non-Goals for Early Refactor
- No broad UI redesign.
- No database schema changes.
- No large behavior changes while ownership boundaries are still unclear.
- No simultaneous deep changes in paste and table/image systems in the same pass.

## Definition of Success
- `main.py` is mostly orchestration and startup wiring.
- Paste behavior can be traced end-to-end in one module path.
- Rich-text features have clear file ownership boundaries.
- Formatting/paste bug fixes become local and predictable.

## Final Note
The current app is already useful and stable enough to protect. The right strategy is incremental refactoring with clear boundaries, not a rewrite.
