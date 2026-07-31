# Where this project stands

Last updated 2026-07-30.

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

### Phase 0 — stamping IDs into the workbook
The backfill script exists and is dry-run-only. It needs updating for the
two-identifier model (`SyncID` per visit, `PatientID` per person) and to apply
the adoption matches. **Nothing has been stamped.**

### Phase 2 — populating the queue sheets
The reconciler decides *what* to write. Nothing *does* the writing. Still
needed: appending rows in camp order, the `Camp - N patients` dividers and
`TOTAL` row, dark red shading on blank required fields, and whole-row inserts so
columns never fall out of alignment.

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
- **`SYNC_LAYOUT_VERIFIED` is still `false`**, which blocks every write path by
  design. Flipping it is deliberate, not incidental.
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

- `docs/inspection-2026-07-30.md` — what the live workbook actually contains
- `docs/column-b-values-2026-07-30.md` — all 111 status values and their routing
- `docs/phase2-decisions.md` — the rules Phase 2 must follow
- `docs/migration.md` — cutover sequence
- `docs/scope-narrowing.md` — path back to `Sites.Selected`
