# TBR Job Numbers — row striping and a totals row

Date: 2026-09-21
Project: `Job Number Creation Sheet Script` (plus one helper in `Job Number Creation Form`)
Status: approved by Joe, ready to implement

## Problem

Two things, both triggered by the Job-Mover deleting and adding rows.

1. The white/gray striping on all three tabs is hand-painted onto cells. When the
   mover deletes a row, the rows below slide up and the stripe breaks: two grays
   or two whites end up adjacent. `TBR - Completed` is already broken this way,
   with rows 2 and 3 both gray.
2. `TBR - Completed` has no totals row. Joe wants one carrying summed Contract
   Value, Estimated Cost and Actual Cost, and moves must insert *above* it rather
   than landing on top of it.

## Established facts

From a diagnostic run against production on 2026-09-21:

| Fact | Value |
| --- | --- |
| Header fill (row 1, all tabs) | `#a4c2f4` |
| Stripe colour on row 2 | `#d9d9d9` (gray) |
| Stripe colour on row 3 | `#ffffff` (white) |
| Native banding in use | none, on any tab; every stripe is painted |
| Column A fill | `#b7e1cd` green / `#ea4335` red, a YES/NO flag |
| Named columns | A–W on all three tabs |
| Totals row on Completed | does not exist |
| Frozen rows | 1 on Job Numbers, 0 on both archive tabs |

Row 2 is gray. Joe's "white, gray, white, gray" counts the header as the first
white, which matches his worked example of row 20 being gray.

## Decisions

| Decision | Choice | Why |
| --- | --- | --- |
| Striping mechanism | Sheets' native alternating colours | Re-flows itself on every insert and delete. Painted fills cannot. |
| Striping columns | B through the last named column | Column A keeps its green/red flag. A painted fill always covers banding, so column A could not show a stripe without losing its colour. |
| Striping extent | row 2 to the last row with data | Joe's choice. No stripes running down the empty sheet. |
| Totals row location | `TBR - Completed` only | Joe's choice. JNS keeps appending at the end. |
| Totals columns | Contract Value, Estimated Cost, Actual Cost | Joe's choice. Job Profit % excluded: summing percentages is meaningless. |
| Totals row identity | the text `TOTAL` in column A | A stray note below it cannot then break the insert position. |
| Setup and one-time work | **manual, by Joe. Not in the code.** | Joe's instruction. The code assumes setup is done and only maintains it. |

## Scope boundary: setup versus maintenance

This is the governing constraint. **The script never creates banding and never
creates the totals row.** Joe does both by hand, once. The script only keeps them
correct as rows come and go.

Consequences:

- `refreshBandingExtent_` resizes an existing banding range. If a tab has no
  banding, it logs and returns. It does not create one.
- `findTotalsRow_` returns 0 when a tab has no totals row, and the mover then
  falls back to appending at the end, exactly as it does today.

## Runtime behaviour

### Striping

`refreshBandingExtent_(sheet)` stretches or shrinks the tab's single existing
banding range to `B2:<lastNamedCol><lastDataRow>`. It is idempotent and does
nothing when the range is already right.

It runs:

- after every job move, on both the source tab and the destination tab;
- from `onChangeInstallable`, an installable On change trigger, on `INSERT_ROW`
  and `REMOVE_ROW`, which covers rows a person inserts or deletes by hand;
- from the lead form script, after it appends a row.

The third one exists because the form runs in a separate Apps Script project, and
an Apps Script change does not reliably fire another script's triggers. The
helper is duplicated into that project rather than shared, matching the existing
convention there: `generateProjectNumber` is already duplicated the same way.

### Totals row

On a move into a tab that has a totals row:

1. `insertRowBefore(totalsRow)` opens a gap; the totals row shifts down by one.
2. The job row is written into the gap, formatting first then frozen values, as
   the mover already does.
3. `updateTotals_` rewrites each configured SUM bounded to `row 2` through the
   row above totals, so the formulas cannot drift as rows come and go.
4. The banding range is refreshed on both tabs.

Rewriting the formulas is necessary rather than cosmetic. A `=SUM(J2:J9)` written
by hand does **not** expand when a row is inserted at row 10, so the newest job
would be left out of the total.

The source row is still deleted last, so no failure path can lose a row.

## What does not change

The header-match guard, the confirm dialog, the document lock, the row-identity
re-check after the dialog, and the frozen-value copy all stay exactly as they
are. The guard compares row 1 only, so a totals row does not affect it.

## Manual steps Joe performs

Per tab (`TBR Job Numbers`, `TBR - Completed`, `TBR - JNS`):

1. Select `B2` to the last named column, down to the last data row. Clear the
   fill so the painted white and gray are gone and banding can show through.
2. Format > Alternating colours on that same range. Header off. Colour 1
   `#d9d9d9`, colour 2 `#ffffff`.

On `TBR - Completed` only:

3. Add a row under the last job. Put `TOTAL` in column A, bold it, give it a top
   border. Leave the sum cells empty; the script fills them on the first move.

Once:

4. Triggers page: add an installable **On change** trigger pointing at
   `onChangeInstallable`.

## Risks

- Clearing fills across `B2:W` also clears any other highlight colour a person
  applied in that area. Column A is outside the range and is safe. A throwaway
  snippet that clears only white and gray is available if wanted.
- If Joe skips the banding setup on a tab, striping silently does nothing there.
  The script logs it rather than creating banding, by instruction.

## Testing

Fake Sheets objects driven from the real production layout, as with the existing
mover tests. Cases: insert lands above the totals row and pushes it down; totals
formulas span exactly the data rows; a tab with no totals row still appends at
the end; the banding range is resized to the data extent; a tab with no banding
is left alone without throwing; the header guard still refuses a mismatched tab.

## Addendum: what an appended row inherits

Added after Joe reported that a form-appended row picked up the stripe but no
gridlines.

`appendRow` adds a completely bare row. Native banding still looks right on it,
because banding is a range rule rather than a cell property, but everything that
*is* a cell property arrives empty. Two helpers in the form script fill that gap,
both running after the row is safely written and both wrapped in try/catch:

| Helper | Copies from the row above | Why it is needed |
| --- | --- | --- |
| `inheritRowFormat` | borders and gridlines, fonts, alignment, currency and date number formats | `PASTE_FORMAT`. Runs first. |
| `inheritRowValidation` | the YES/NO, Project Manager and Job Status dropdowns | Runs second, so it always has the last word. |

A row inserted by hand inherits all of this natively. An append does not.

The Job-Mover needed the same treatment on its side, and takes it from the row
above **on the archive tab** rather than from the source row. Copying from the
source drags that row's own quirks across, and a tracking-sheet row that was
itself appended bare would land in the archive just as bare. When the archive has
no data row to copy from, which is the first job ever moved into a tab, it falls
back to the source row. Row 1 is never a formatting source.

## Open item: column A's flag colour

Column A's green/red fill is painted by hand, not driven by the YES/NO value.
`inheritRowFormat` copies it down with the rest of the formatting, so when the
row above says YES and the new lead says NO, the new row inherits the wrong
colour until somebody notices.

The fix is to replace the painted fill with a conditional-formatting rule keyed
on the cell value. That would make a wrong colour impossible, survive every
insert and delete for free, and remove the only caveat on inheriting formatting.
Joe chose to keep the painted fill for now; this is here so the trade-off is not
forgotten.
