# Phase 2b — appending the patients who have no queue row

Built 2026-07-31. **Not yet run against the live workbook.** Everything below is
verified by tests against a simulated sheet; nothing here has touched the real
file.

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
