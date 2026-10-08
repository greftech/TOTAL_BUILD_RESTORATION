# Doc Gen — CLAUDE.md

Sheet-bound Apps Script on the **TBR Job Numbers** spreadsheet. Adds a **TBR Docs** menu + sidebar that merges a project row into Doc templates (Contract, Work Authorization, Mold Waiver, Certificate of Completion), drops the PDF + editable Doc into the project's `Documents` subfolder, and dual-logs every run. See [README.md](README.md) for end-user docs and one-time setup.

## Two source-of-truth surfaces

1. **`Code.gs` constants** — `CONFIG_SHEET_ID` (must be set), `JOB_NUMBERS_TAB`, tab name constants, `TIMEZONE`, `LOCK_TIMEOUT_MS`.
2. **Admin Config spreadsheet** (referenced by `CONFIG_SHEET_ID`) — 4 tabs: `Doc Types`, `Placeholder Map`, `Required Fields`, `TBR Document Activity`. Doc-type definitions, template IDs, placeholder/format mappings, and required fields all live here, **not in code**. Adding a 5th doc type is a config change; no code edits.

If a behavior depends on a doc type or a placeholder, it belongs in the Config sheet. Resist the pull to special-case in `Code.gs`.

## Invariants — do not break these

- **Additive only.** Every Generate run creates a new timestamped Doc + PDF. No overwrite, no delete, no rollback. Re-running for the same doc type is allowed and produces a second pair.
- **Atomic per doc type.** A failure on one selected doc type logs an error row and continues; the rest of the selection still runs.
- **Whole run is serialized.** `LockService.getScriptLock()` with `LOCK_TIMEOUT_MS` wraps the full generate sequence — don't move it inside the per-doc loop.
- **Dual-log asymmetry:** the admin `TBR Document Activity` tab gets **every** event (success + error). The per-project `Activity Log` Doc gets **successes only** — it's a completions narrative, not a debug log.
- **Doc Gen never creates the project folder.** The `Project Folder` column is a precondition; if it's blank/invalid, Generate is blocked. Folder provisioning is upstream (`Job Number Creation Form/`).
- **Shared drives require the Advanced Drive Service.** All Drive calls use `Drive.*` with `supportsAllDrives: true` (and `includeItemsFromAllDrives: true` on lists). Don't replace these with `DriveApp` — the project folders frequently live in a shared drive and `DriveApp` won't see them.
- **`initializeConfigSheet` is editor-only.** It is intentionally not wired to the menu — it's a one-off bootstrap. Keep it idempotent if you touch it.

## Coexistence on the same spreadsheet

There's a second bound script on TBR Job Numbers (`../Job Number Creation Sheet Script/`). Both define `onOpen`; Apps Script runs each project's `onOpen` separately, so both menus appear. Don't rename this project's menu to `TBR` or anything that would collide — keep it `TBR Docs`.

## Editing the code

- Per repo convention (`../CLAUDE.md`): the Apps Script web editor is **not** the source of truth. Any change made in the web editor must be copied back here to stay versioned.
- One `.gs` file, kept flat. The web editor doesn't preserve subfolders.
- Secrets (none today, but if any are added) live in Script Properties, never in source.
- `FOLDER_URL_COLUMN_NAME` is overridable via Script Property; default is in code as `FOLDER_URL_COLUMN_DEFAULT`. Read through `getFolderUrlColumnName_()`, don't hardcode.

## Common pitfalls

- Adding a new field format: extend the `Field Format` switch in the merge engine **and** document the value in the Config-sheet schema section of the README. The sheet is what users edit.
- Touching the activity-log row shape: the `TBR Document Activity` header is written by `initializeConfigSheet`. If you reorder columns, existing admin sheets will silently keep the old header — bump and re-scaffold deliberately, or write a migration note.
- Date/currency formatting uses `TIMEZONE` (`America/New_York`). Don't pass `Session.getScriptTimeZone()` — installs may be on a different TZ.
