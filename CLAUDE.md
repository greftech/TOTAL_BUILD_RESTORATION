# TBR — Total Build Restoration Apps Scripts

Google Apps Script projects supporting Total Build Restoration's internal operations (job tracking, budget ingestion, folder provisioning). Each subfolder mirrors a single Apps Script project; this repo is the source of truth for that code.

## Layout

- `Budget PDF Ingestion/` — Sheets-bound script for the TBR Budget Template. Sidebar UI sends a scope-of-work PDF to Gemini, parses line items, and writes them into the budget sheet. Logic lives in the `TBRBudgetTools` library project (also tracked here under `Budget PDF Ingestion/TBRBudgetTools/`); the bound script is a thin wrapper exposing menu items, sidebar callbacks, and an `onEdit` gridline guard.
  - `Code.gs` — bound script (menu, sidebar, `google.script.run` entry points, `onEdit`).
  - `Sidebar.html` — sidebar UI.
  - `TBRBudgetTools/TBRBudgetTools.gs` — library: Gemini call, budget-sheet writes, formatting.
- `Job Number Creation Form/` — Form-bound script. On submit, allocates the next `TBR-<year>-<seq>` job number, creates the project folder from a category template in Drive, and appends a row to the central tracking sheet.
- `Job Number Creation Sheet Script/` — Sheet-bound script for the TBR Job Numbers tracking sheet. Edit trigger fires when "Create Job #?" flips to YES, retroactively provisioning folders / job numbers for rows that bypassed the form.

## Conventions

- Apps Script projects: one `.gs` file per logical unit, kept flat. The Apps Script editor doesn't preserve subfolders, but we mirror the project name as a folder here for clarity.
- Configuration constants (sheet IDs, template folder IDs, prefixes) sit at the top of each `Code.gs`. The Sheet Script project intentionally points at a **test copy** of the tracking sheet — production IDs live in `Job Number Creation Form/Code.gs`.
- Secrets (e.g. `GEMINI_API_KEY`) live in Apps Script **Script Properties** on the owning library project, never in source. Don't paste keys into `.gs` files.

## Working with this repo

- Edits made in the Apps Script web editor must be copied back into this repo to stay versioned. Consider [`clasp`](https://github.com/google/clasp) if/when round-tripping becomes painful.
- When adding a new Apps Script project, create a new top-level folder named after the project and drop its `.gs` / `.html` files in. Update this file's Layout section.
