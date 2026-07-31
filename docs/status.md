# Where this project stands

Last updated 2026-07-31.

## The workbook is fully reconciled

A full cycle now plans **nothing**: zero stamps, zero appends, zero write-backs,
zero removals. Every patient a keyword routes to a queue has a row on that queue,
and every one of those rows is linked to the daily row it came from.

    Missing Info           195 rows, 26 camps
    Not Accepted           375 rows, 35 camps
    Ineligible & Inactive  137 rows, 18 camps
    Verify Insurance       243 rows, 17 camps
    United Refuah           85 rows, 13 camps
    ----------------------------------------
                         1,035 patient rows

Three things are outstanding and all three need a human, not code:

- **2 orphan queue rows.** `Not Accepted` row 269 carries no SyncID (no date of
  birth to match on). `Missing Info` row 70 carries one no daily sheet knows.
- **1 unmarked edit.** `Not Accepted` row 377's Date of Birth disagrees with its
  source row. Nobody marked it Resolved, so nothing is written — see below.

## Done and verified against the live workbook

- **Phase 1, read-only reconciliation.** ~1s per cycle, byte-identical output
  across runs, zero unrecognized column B values.
- **Phase 2a, adoption.** Every linkable queue row shares a SyncID with its
  source. 19 more were linked once the header-row bug stopped hiding them.
- **Phase 2b, appending.** All 1,035 rows, including the first live row inserts.
- **Phase 2c, the `Resolved` column and the house style**, on all five tabs.
- **Phases 3 and 4**, driven by the marker: `npm run resolve-rows`.
- **The cycle applies.** `runCycle` no longer throws above phase 1; the phase
  gate decides how much of a plan runs.
- **The concurrency test finally has an answer.** In-session: p50 125 ms, p95
  155 ms, 40/40 writes, zero locks, zero throttling, zero read-back mismatches.

## The rule that matters most

**Nothing reaches a daily sheet without a human saying so.**

The reconciler cannot tell a staff edit from a stale mirror value — both are
"the queue differs from the source", and roughly 700 rows were adopted from
mirror sheets that stopped updating on 2026-07-30, so on those rows the daily
sheet is the *newer* copy. A write-back therefore requires the `Resolved`
marker. A difference on an unmarked row is reported (`queue.unmarked_edit`) and
never written, including when the row is about to be deleted.

## What the live runs exposed

Each of these was silently costing patients or was one run away from it.

| | |
|---|---|
| Queue header is row 1, not row 9 | rows 2-9 of every mirror tab were never read: 18 patients invisible |
| `Missing Info (New)` renamed | 194 patients routed to a queue with no tab; only a warning marked it |
| One `QueueRow` per SyncID | a patient on two tabs collapsed to one; every run appended another copy |
| String comparison of values | 1,179 "staff edits" that were case, punctuation, and `8701` vs `08701` |
| Writing `null` | does not clear a cell; the clear-column-B step would have done nothing |
| Blank on the queue | was written back as `""`, which *does* clear — over billing data |
| Deleting on stale row numbers | no identity check before a delete, on a file staff are always in |
| `counts` in every log line | redacted to `[redacted:number]`; the job's observability said nothing |
| The concurrency test | had no workbook session, so it measured its own staleness |
| `Retry-After: 0` | honoured literally; five retries fired inside a tenth of a second |
| Write-backs needed no signal | a stale mirror value could overwrite a corrected one on the billing copy |

Full detail in [`docs/phase2b.md`](phase2b.md).

## Not built

### Phase 5 — the notes sheet
Not started. Adding `Notes` shifts every column right on all five tabs again;
`scripts/migrate-queue.ts` is now the tool for that and its guard is structural,
so it can be pointed at the job.

### Azure infrastructure — DEPLOYED 2026-07-31, code not yet published
Resource group `camp-workbook-sync` in eastus holds all eight resources:
Consumption plan, storage `ste24s24d74i4egcampworkb` with the state container,
Log Analytics capped at 1 GB/day, Application Insights, the Function App, and a
user-assigned managed identity.

**The code has not been published to it and nothing runs on a timer yet.**

The identity cannot reach the workbook until a Global Administrator assigns it
`Files.ReadWrite.All` on Microsoft Graph — Bicep cannot do that. Until then the
Function App uses the app registration's client secret, which already carries
that permission. Both routes and the swap between them are in
[`infra/README.md`](../infra/README.md).

## Before the timer is enabled

- **The half of the concurrency test that needs people.** Whether a human's
  in-flight edit survives a concurrent write. That needs two staff in the file
  during a run and cannot be answered from here. What can be said: this session
  made several hundred live writes and every reconcile afterwards matched.
- **Two settings in Excel, per tab.** Graph refuses both: freeze the header row,
  and the `Resolved` dropdown. Neither is load-bearing.
- **Rotate the client secrets.** Several were created during setup and at least
  three have appeared in chat transcripts. `az ad app credential list`.

## Open questions

- `lasante-e` vs `lasante-o` — assumed both mean done. 2,043 rows depend on it.
- **72 patients whose repeat visits disagree** on a filled field. Pre-existing.
- The workbook lives in one person's **personal OneDrive**. If they leave, it
  goes with their account.

## Reference

- `docs/phase2b.md` — the appender, the Resolved column, and every live finding
- `docs/inspection-2026-07-30.md` — what the live workbook actually contains
- `docs/column-b-values-2026-07-30.md` — all 111 status values and their routing
- `docs/phase2-decisions.md` — the rules Phase 2 must follow
- `infra/README.md` — provisioning, and the levers that cost money
- `docs/migration.md` — cutover sequence
