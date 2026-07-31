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

| Queue | Rows to append | Tab state | Inserts needed |
|---|---|---|---|
| `Verify Insurance` | 243 | empty, header only | no |
| `Not Accepted ` | 17 | 400 live rows | yes |
| `Ineligible & Inactive` | 11 | 1,221 live rows | yes |
| `Missing Info` | 8 | 229 live rows | yes |
| `United Refuah` | 0 | **done** | — |

`Verify Insurance` has the same safety profile as United Refuah — empty tab, so
every row lands past the end and no gap is opened. The other three need
`--allow-inserts`, which is the first time this project shifts a live row.

**21 queue rows carry no SyncID**, 18 of them the rows 3-9 that were invisible
until today. `npm run adopt:apply` matches 19 of the 20 candidates and would
stamp 38 IDs across 17 range writes; one is an `identity-mismatch` that needs a
human. Until they are linked, a staff edit on those rows cannot reach the daily
sheet — so adoption should run before those three tabs are appended to.

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
