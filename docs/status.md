# Where this project stands

Last updated 2026-07-31.

## Done and verified against the live workbook

- **Auth and access.** App registration, `Files.ReadWrite.All`, admin consent.
  Workbook addressed by `driveId` so it is found wherever it lives (it is in a
  personal OneDrive, not a site).
- **Workbook layout confirmed.** Daily sheets A–R, status in B (`EMR`), camp in
  C, header row 1, data from row 2. Sheet naming `July 30, 2026`.
- **Column B vocabulary.** All 111 distinct values staff have typed are routed,
  deliberately ignored, or terminal. **Zero unrecognized.** Includes the
  1,268-row `needs ohi` / `needs lasante` trap and the 366 `dont accept` rows
  the brief never mentioned.
- **Phase 1, read-only reconciliation.** Runs end to end in ~1.0s per cycle,
  producing byte-identical output across repeated runs.
- **Performance.** Full 46-sheet scan in 0.7s, down from 66s.
- **Adoption of existing queue rows.** 680 of 681 link to their source row.
- **Two tabs created**: `Verify Insurance` and `United Refuah`.

## Not built yet

### Phase 0 / 2a — stamping IDs — DONE
679 of 681 existing queue rows now share a SyncID with their source row on the
daily sheet. 1,358 IDs written to column BA across 39 range writes; re-running
finds nothing to do. One row remains unlinked pending a date of birth.

The standalone `phase0-backfill.ts` still needs updating for the two-identifier
model (`SyncID` per visit, `PatientID` per person) before it is used.

### Phase 2b — appending the missing rows — UNITED REFUAH DONE
`src/domain/append.ts` plans the appends and `npm run append -- "<queue>"`
applies them: camp order, `Camp - N patients` dividers, a `TOTAL` line, dark red
shading on blank required fields, and whole-row inserts. Dry run by default, one
queue per run, and it refuses to open gaps in a live sheet without
`--allow-inserts`.

**`United Refuah` (85 patients, 13 camps) and `Verify Insurance` (243 patients,
17 camps) are written.** Both were empty tabs, so neither needed a row insert,
and both now plan nothing on a re-run. Remaining: `Not Accepted ` (11),
`Ineligible & Inactive` (5) and `Missing Info` (1) — 17 rows that all need
`--allow-inserts`, the first time this project shifts a live row.
See [`docs/phase2b.md`](phase2b.md).

### Phase 2c — the Resolved column and the house style — DONE on all five tabs
`Resolved` sits between `Source Row` and `Last Name` on the four work queues and
is absent from `United Refuah`. Marking a row `Done` writes the fix back to the
daily sheet, clears column B at source and deletes the queue row —
`npm run resolve-rows`, dry run by default. Nothing is marked yet, so it has
not been run for real.

All five tabs carry the house style: dark header bar, camp bands, TOTAL line,
`mm/dd/yyyy` dates, centred `Source Row`, sized columns. `npm run migrate` applies
it and doubles as a redraw.

**Two things Graph refuses** against this workbook and are manual one-offs per
tab: freezing the header row, and the `Resolved` dropdown. Neither is
load-bearing.

### Three live defects the first runs exposed — all fixed
- **The queue header row is row 1, not row 9.** The brief's instruction-block
  layout does not exist. Reading from row 10 hid rows 2-9 of every mirror tab —
  18 patient rows invisible to Phase 1 and missed by Phase 2a. That is where
  "679 of 681" came from.
- **`Missing Info (New)` was renamed `Missing Info`.** The tab stopped
  resolving, and 194 patients were being routed to a queue with no sheet. Only a
  `queue.sheet_missing` warning marked it. A tab rename is a silent outage here.
- **The reconciler would have duplicated a patient on every run.** `queueById`
  held one row per SyncID, so a patient sitting on two queue tabs — the right one
  and a stale one awaiting Phase 4 — collapsed to whichever tab parsed last. When
  that was the stale row, it planned an append to the queue the patient was
  already on. Caught by re-running the appender against a tab just written; the
  index now keeps every row an ID appears on.
- **Writing `null` does not clear a cell.** Graph returns 200 and leaves it
  untouched; only the `/clear` action empties it. This would have made the
  resolve flow's "clear column B at source" step silently do nothing, so every
  resolved patient came straight back. Found before it ran.

**Adoption re-run: 19 of 20 unlinked rows stamped.** Two orphans remain, both
needing a human — one `identity-mismatch`, one `unknown-sync-id`.

### Phase 3 — write-back to the daily sheets
Intents are computed; no applier. The patient fan-out logic
(`resolveFieldAcrossVisits`) is built and tested but not yet wired into the
reconciler.

### Phase 4 — clear and remove
Intents are computed; no applier.

### Phase 5 — the notes sheet
Not started. Note that adding `Notes` after `Source Row` shifts every column
right on all five tabs — a migration, not an edit.

### Azure infrastructure — none of it exists
No resource group, storage account, Function App, Application Insights, or
managed identity. **Nothing runs on a timer.** Every result so far came from
running the code by hand. The GitHub Actions deploy workflow will fail until
the Function App and OIDC credentials exist.

## Before writes are enabled

- **Concurrency test.** Not run. The one thing the brief insists precedes
  Phase 2.
- **`SYNC_LAYOUT_VERIFIED` is now `true`.** That gate is spent; future write
  scripts will not stop to ask.
- **Rotate the client secrets.** Several were created during setup and at least
  two appeared in a chat transcript. `az ad app credential list` shows them.
- **Reconciliation gate.** The brief wants old-vs-new diffed for several days
  before cutover. The old formulas are already gone, so the baseline is the
  static snapshot of the mirror sheets rather than live formula output.

## Open questions

- `lasante-e` vs `lasante-o` — assumed both simply mean done, and the initial
  (Esti / Osnat) does not affect routing. 2,043 rows depend on it.
- **72 patients whose repeat visits disagree** on an already-filled field — 36
  on phone number, 16 on billing address. Not created by this project; it is
  pre-existing data that nothing previously surfaced. A list can be produced.
- **1 queue row** whose Date of Visit resolves to no daily sheet.
- The workbook lives in one person's **personal OneDrive**. If they leave, it is
  deleted with their account. Unrelated to this build, but worth moving.

## Reference

- `docs/phase2b.md` — how the appender works and how to run it
- `docs/inspection-2026-07-30.md` — what the live workbook actually contains
- `docs/column-b-values-2026-07-30.md` — all 111 status values and their routing
- `docs/phase2-decisions.md` — the rules Phase 2 must follow
- `docs/migration.md` — cutover sequence
- `docs/scope-narrowing.md` — path back to `Sites.Selected`
