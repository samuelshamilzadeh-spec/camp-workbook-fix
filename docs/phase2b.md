# Phase 2b — appending the patients who have no queue row

Built and first run 2026-07-31. **`United Refuah` is done: 85 patients across 13
camp blocks, written to the live workbook.** The four other queues are not.

## The first live run

    to append:     85 patients
    rows added:    98 (patients plus one divider per new camp)
    row inserts:   0 — writes only past the end of the sheet
    operations:    3

    --- after ---
    patient rows:  85   (expected 85)
    camp blocks:   13
    TOTAL row:     100 declaring 85
    rows carrying a SyncID: 85

A second run plans nothing, which is the property that matters: the source rows
were stamped, so the reconciler no longer thinks they need a queue row.

Three range writes for 85 patients — the block, the SyncID column, the TOTAL —
plus 43 range writes stamping the source rows across the daily sheets. No row
inserts, no shading, nothing written back to a daily sheet.

## What the first run found

Two live defects, neither of them in the appender, both of which had been
silently costing patients:

**The queue header row is row 1 on all five tabs, not row 9.** The build brief
described an instruction block in rows 1-7 with headers on row 9, and
`LAYOUT.queue` was set to that on the brief's authority. There is no instruction
block. Reading from row 10 meant **rows 2-9 of every mirror tab were never
read**: 18 real patient rows invisible to Phase 1 and skipped by Phase 2a's
adoption, which is where the "679 of 681" figure came from. `detectQueueShape`
now reads the header off the sheet, and those rows are visible.

**`Missing Info (New)` had been renamed to `Missing Info`.** `resolveSheetName`
reported the tab as absent, the cycle logged `queue.sheet_missing`, and **194
patients were being computed into a queue with nowhere to put them** — no error,
no failure, just a warning nobody was reading. With the name corrected the
figure drops to 8, because the other 186 were already sitting on the tab all
along.

A tab rename is a silent outage in this design. Worth knowing before the office
plans one.

**The reconciler would have duplicated a patient on every run.** `queueById` was
keyed by SyncID alone, one row per ID. A patient whose column B changed sits on
two queue tabs for a while — the new one, and the old one waiting for Phase 4 to
remove it — so the second tab parsed overwrote the first, and which survived
depended on the order of `QUEUE_SHEET_NAMES`. When the survivor was the stale
row, the reconciler concluded "wrong queue" and planned an append to the queue
the patient was **already on**. Every run would have added another copy.

Caught by re-running the appender against Verify Insurance immediately after
writing it: 243 rows written, and it still wanted to append one of them. The
index now holds every row an ID appears on; a patient already on the right queue
is never appended, stale rows on other queues are removed, and a genuine
duplicate on one queue is reported rather than guessed at.

Nothing was duplicated on the workbook — the second run was a dry run, which is
the entire reason for checking idempotency before trusting a write path.

## What it does

`npm run append -- "<queue>"` appends every patient the reconciler says belongs
on that queue and has no row there yet. One queue per run.

    npm run append -- "United Refuah"                    # dry run, prints the plan
    npm run append -- "United Refuah" --apply            # writes
    npm run append -- "United Refuah" --apply --allow-inserts

Per queue, it writes:

- patient rows in camp order, 17 columns in one range write per contiguous block
- a `Camp - N patients` divider above each camp block
- a `TOTAL - N patients` line at the foot
- the `SyncID` in column `BA`, blank on divider rows
- dark red (`#C00000`) shading on required fields that are blank
- whole-row inserts when rows have to go into the middle of a live sheet

`src/domain/append.ts` decides all of it and touches nothing;
`scripts/append-queue.ts` executes what it produces. So the plan is printable and
reviewable before a single cell changes.

## Where the remaining queues stand

Measured after the United Refuah run, with both defects above corrected:

| Queue | Rows to append | Tab state | Migrated |
|---|---|---|---|
| `United Refuah` | 0 | **done** — 85 rows, 13 camps | yes |
| `Verify Insurance` | 0 | **done** — 243 rows, 17 camps | yes |
| `Not Accepted ` | 11 | 400 live rows | yes |
| `Ineligible & Inactive` | 5 | 1,221 live rows | yes |
| `Missing Info` | 1 | 229 live rows | yes |

All five now carry the house style, real dates and — on the four work queues —
the `Resolved` column. The 17 remaining appends need `--allow-inserts`.

The two done tabs were empty, so every row landed past the end and no gap was
opened. The remaining 17 rows all need `--allow-inserts`, which will be the
first time this project shifts a live row — and the first live exercise of the
dark red shading, since 11 of them have a blank required field.

Adoption has been re-run: **19 of the 20 unlinked rows are now stamped**, 38 IDs
across 17 range writes. Two orphans remain and both need a human — one
`identity-mismatch`, one `unknown-sync-id`.

Also outstanding, computed but not applied: **9 `remove-queue-row` intents**,
patients whose column B was resolved at source and who should come off their
queue. That is Phase 4, which does not exist.

## Run United Refuah first

It is the safest write available and it exercises the whole path:

- the tab is empty, so every row is written past the end of the sheet and **no
  row insert happens at all** — the one operation that could misalign columns is
  not used
- `REQUIRED_FIELDS['United Refuah']` is empty, so nothing is shaded
- it is append-only: the office confirmed a row lands there and is never
  changed, so no write-back to a daily sheet can follow

85 rows. Then `Verify Insurance` (also new, also empty), then the three tabs that
already hold live rows — those need `--allow-inserts`.

## Three decisions the implementation makes

**Stamps go in before appends.** The reconciler decides what to append by asking
which source rows have no queue row, which it can only ask of a row carrying an
ID. Stamp first and an interrupted run leaves a stamped row with no queue row,
which the next run appends once under the same ID. Append first and an
interrupted run leaves an unstamped row, which the next run mints a *new* ID for
and appends *again*. Same crash, opposite outcome — one recovers, the other
duplicates patients silently.

**The plan is emitted bottom-up.** Inserting rows shifts everything below and
invalidates any row number already computed for the rows underneath. Rather than
recomputing offsets, every address stays in the coordinates of the sheet as it
was read and the operations run from the bottom of the sheet upwards, so an
insert higher up carries already-written content down without invalidating
anything still pending. `tests/append.test.ts` applies a plan to a simulated
sheet and checks that every existing row keeps its own SyncID — the failure a
column-scoped insert would cause.

**Inserts are refused unless asked for.** A plan that only writes past the end of
a sheet cannot misalign anything. `--allow-inserts` turns the other case into a
decision instead of a surprise, and the dry run says which kind of plan it is.

## What changed outside the appender

**`Source Row` is now a plain number.** It was being written as
`July 30, 2026!B45`, which reads well but is not a number — and `planAdoption`
parses that cell with `Number()`. Left alone, every row this project appended
would have been unadoptable, and the column would hold two different kinds of
thing. `Date of Visit` still carries the daily sheet's own name
(`July 30, 2026`); `resolveSourceSheet` accepts that or an Excel serial, and the
name needs no locale-dependent parsing on the way in.

**The queue header row is read off the sheet, not taken from config.**
`LAYOUT.queue` says headers on row 9 and data from row 10, which is right for the
four mirror tabs and wrong for `Verify Insurance` and `United Refuah` — this
project created those with `create-queue-sheet.ts` and their header is on row 1.
Trusting the config there would have skipped rows 2-9 when reading and left eight
blank rows above the first patient when writing. `detectQueueShape` finds the
header by its labels and falls back to the config; `parseQueueSheet` reports
which happened, and **the applier refuses to write when the shape was guessed.**

## Open, and worth a human's eye on the first dry run

- **The row-insert address form is untested against Graph.** Two spellings reach
  the same operation and the documentation does not settle which one the workbook
  API takes. `Workbook.insertRows` tries `range(...)/entireRow/insert` and falls
  back to `range('10:12')/insert` on a 400, which is safe because a 400 means
  Graph rejected the URL before doing anything. United Refuah does not exercise
  this at all.
- **New camp blocks go at the end of the body, alphabetically among themselves.**
  Camps already on a sheet keep their place and are never reordered. Nobody has
  said what order the live tabs are in, so nothing tries to guess it.
- **No blank spacer row between camp blocks.** Adding one later is harmless —
  `parseQueueSheet` ignores blank rows.
- **`Date of Visit` is written as text** (`July 30, 2026`) while the existing
  rows on the mirror tabs hold Excel serials. Both read correctly; the column
  will look mixed on the three populated tabs.
- **A camp with no name gets a block labelled `(no camp)`.** Ugly on purpose:
  the label has to equal what `campKey` produces for a missing camp, or the next
  run would not recognize its own block and would write another one below it.

---

# The Resolved column, and what it costs

Added 2026-07-31, and applied to all five tabs.

## How a row leaves a queue

A staff member fills in whatever was missing on the queue row, marks `Resolved`,
and `npm run resolve-rows -- --apply` does three things in an order that is not
negotiable:

1. **Writes the fix back to the daily sheet.** Until this runs, the queue row is
   the only place that phone number exists. Deleting first would destroy the work
   the marker is reporting.
2. **Clears column B at the source.** The reconciler decides everything from what
   column B says right now, so a row still reading `missing info` is re-appended
   within seconds. Blank means "not processed yet", which returns the patient to
   the office's normal flow to be entered into an EMR.
3. **Deletes the queue row**, whole-row, bottom-up per tab — deleting row 40
   moves row 41 up into it, so a top-down pass would delete the wrong patient
   from the second row onwards.

The marker is an allow-list: `done`, `yes`, `y`, `x`, `fixed`, `complete`,
`completed`, `resolved`, `true`, `✓`. Anything else in that cell is **reported
and never acted on**. Treating any non-blank cell as a signal would mean a stray
keystroke silently pulls a patient off the queue and wipes their status at
source, and the column sits right next to Last Name.

`United Refuah` has no `Resolved` column. It is an append-only record — a row
lands there and never changes — so there is nothing to resolve, and the marker
would only be a cell nobody should touch. That makes the column layout
per-destination: 18 columns on the work queues, 17 there.

## Two things Graph will not do

Both were built, both were rejected by the API against this workbook, and both
are now one-off manual settings that live in the file once set:

| Wanted | What happens | Do this instead, once per tab |
|---|---|---|
| Freeze the header row | 400 on every spelling of `freezePanes` | View > Freeze Panes > Freeze Top Row |
| `Resolved` dropdown | 400 on `dataValidation`, even on a GET | Data > Data Validation > List > `Done` |

Neither is load-bearing. The allow-list accepts the words staff type, so the
marker works with or without a dropdown to click.

## Writing null does not clear a cell

Graph accepts a values PATCH containing nulls, returns 200, and leaves every one
of those cells exactly as it was. Verified directly: three cells holding SyncIDs,
PATCHed with `[[null],[null],[null]]`, still held them; the `/clear` action
emptied them.

This was found cleaning up after the column insert and it mattered a great deal
more than it looked. Step 2 above — clearing column B — was written as a null.
It would have returned 200, done nothing, left the keyword in place, and the
patient would have come back on every cycle forever. Anything that needs a cell
to become empty goes through `Workbook.clearRange`.

## The column insert moves more than you think

Inserting `Resolved` at C shifts **every** column to its right, which is the
correct behaviour and the reason a whole-column insert was used rather than
rewriting a fixed window of A..Q: `Ineligible & Inactive` carries
`Updated Insurance Carrier`, `Updated Insurance ID #` and `Updated Medicaid #`
out at R, S and T, and a windowed rewrite would have destroyed them. They moved
to S, T and U, intact.

It also shifts `BA` — the SyncID column — to `BB`, which would leave every link
between a queue row and its patient's daily row in a column nothing reads. The
migration reads the IDs first, puts them back in `BA` afterwards, and clears
`BB`. Confirmed after all five tabs: orphans unchanged at 2, appends unchanged
at 17.

## The house style

In `STYLE`, drawn by `src/domain/style.ts`, applied by `npm run migrate`:

- one dark header bar (`#051C2C`), white and bold
- horizontal hairlines only — vertical lines are what make a sheet look like a
  form rather than a table
- camp dividers as quiet bands (`#E8EBEE`), bold, the width of the table
- the grand total set apart by weight and a rule, not by colour
- `Source Row`, `Resolved`, `Gender` and `State` centred; everything else left
- `Date of Visit` and `Date of Birth` as real dates, `mm/dd/yyyy`
- every column sized to what it holds

Subtractive on purpose. The one loud thing stays loud: the office's red on a
blank required field is the only colour that catches the eye.

`migrate` doubles as a redraw, because the divider and TOTAL counts are snapshots
and go stale the moment a row is removed. It also converts the office's bare camp
labels (`Achim`) into counted ones (`Achim - 6 patients`) and adds the TOTAL line
none of their tabs had.
