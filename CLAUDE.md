# TBR — Total Build Restoration Apps Scripts

Google Apps Script projects supporting Total Build Restoration's internal operations (job tracking, budget ingestion, folder provisioning). Each subfolder mirrors a single Apps Script project; this repo is the source of truth for that code.

## Layout

- `Budget PDF Ingestion/` — Sheets-bound script for the TBR Budget Template. Sidebar UI sends a scope-of-work PDF to Gemini, parses line items, and writes them into the budget sheet. Logic lives in the `TBRBudgetTools` library project (also tracked here under `Budget PDF Ingestion/TBRBudgetTools/`); the bound script is a thin wrapper exposing menu items, sidebar callbacks, and an `onEdit` gridline guard.
  - `Code.gs` — bound script (menu, sidebar, `google.script.run` entry points, `onEdit`).
  - `Sidebar.html` — sidebar UI.
  - `TBRBudgetTools/TBRBudgetTools.gs` — library: Gemini call, budget-sheet writes, formatting.
- `Job Number Creation Form/` — Form-bound script. On submit, allocates the next `TBR-<year>-<seq>` job number, creates the project folder from a category template in Drive, and appends a row to the central tracking sheet.
- `Job Number Creation Sheet Script/` — Sheet-bound script for the TBR Job Numbers tracking sheet. Edit trigger fires when "Create Job #?" flips to YES, retroactively provisioning folders / job numbers for rows that bypassed the form.
- `Doc Gen/` — Second sheet-bound script on the TBR Job Numbers spreadsheet. Adds a "TBR Docs" menu + sidebar that merges project-row fields into Doc templates (Contract, Work Authorization, Mold Waiver, Certificate of Completion), saves PDF + editable Doc into the project's `Documents` subfolder, and dual-logs every generation (admin Activity tab + per-project Activity Log Doc). Config (doc types, placeholder map, required fields, activity log) lives in a separate admin spreadsheet, referenced by `CONFIG_SHEET_ID` at the top of `Code.gs`. `initializeConfigSheet` is a one-off scaffold function run from the Apps Script editor, not menu-wired. Uses the Advanced Drive Service (`supportsAllDrives: true`) so project folders inside shared drives work.
  - `Code.gs` — menu, sidebar entry, `google.script.run` callbacks, merge engine, dual-log, one-time `initializeConfigSheet` bootstrap.
  - `Sidebar.html` — vanilla HTML+JS sidebar (300px), project dropdown, validation block, doc-type checkboxes, generate + results.
  - `appsscript.json` — manifest with Advanced Drive Service enabled and required OAuth scopes.

## Conventions

- Apps Script projects: one `.gs` file per logical unit, kept flat. The Apps Script editor doesn't preserve subfolders, but we mirror the project name as a folder here for clarity.
- Configuration constants (sheet IDs, template folder IDs, prefixes) sit at the top of each `Code.gs`. The Sheet Script project intentionally points at a **test copy** of the tracking sheet — production IDs live in `Job Number Creation Form/Code.gs`.
- Secrets (e.g. `GEMINI_API_KEY`) live in Apps Script **Script Properties** on the owning library project, never in source. Don't paste keys into `.gs` files.

## Working with this repo

- Edits made in the Apps Script web editor must be copied back into this repo to stay versioned. Consider [`clasp`](https://github.com/google/clasp) if/when round-tripping becomes painful.
- When adding a new Apps Script project, create a new top-level folder named after the project and drop its `.gs` / `.html` files in. Update this file's Layout section.
