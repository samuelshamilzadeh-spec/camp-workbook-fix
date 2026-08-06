# Where this project stands

Last updated 2026-08-06.

## The two big queues are split by month

`Missing Info` and `Ineligible & Inactive` are now four tabs each — June, July,
August, and `Other` for a visit outside those months or a date that will not
parse. The other three queues are unchanged.

The routing change is small on purpose. A column B keyword picks a **family**, as
it always did; the visit's month picks which of that family's **tabs** the row
sits on; `src/domain/segments.ts` is the only place the two meet. No keyword rule
was touched, and "wrong month" reuses the existing wrong-queue path rather than
introducing a second kind of move.

**Done on the live workbook, 2026-08-06.**

    Missing Info           212 rows ->  June 21   July 175   August 16   Other 0
    Ineligible & Inactive  162 rows ->  June 17   July 145   August  0   Other 0

Every row confirmed by SyncID on arrival, and a full cycle afterwards plans
nothing — 0 stamps, 0 appends, 0 write-backs, 0 removals across 52 daily sheets.
The originals are now `Missing Info (old)` and `Ineligible & Inactive (old)`,
hidden and still holding every row exactly as it stood.

**The colours are copied, not re-derived, and that was the whole lesson.** The
first version of the script regenerated fills from `REQUIRED_FIELDS` and
`planQueueStyle`, on the reasoning that every fill on these tabs came from those
rules. The live tab disproved it: 17 distinct fill colours on `Missing Info`,
only four of which this codebase knows about — amber and yellow in `Date of
Visit`, blues in `Source Row`, red note-cells reading `Staff kid`, greens for
confirmed insurance, orange down `Medications`, purple down `Allergies`. Staff
built that by hand over a season. Regenerating would have erased all of it, and
the run that would have done so reported it as an improvement.

Font colour is copied for the same reason — a `Medicaid #` reading `inactive` in
red is somebody saying something.

Verified afterwards by an independent colour census of source versus destination.
Every patient-cell colour matched exactly. Three counts differed and all three
are explained: the header bar and TOTAL row now exist once per tab, a camp with
patients in two months gets its divider band (and any colour on it) in both, and
one dark-red cell sat on the `YBH` divider — one of **7 camp headings on the old
tab that had no patients under them at all**, so there was no block for it to
move to.

Excel's own copy would have been better than any of this. Graph refuses it: both
`worksheets/copy` and `range/copyFrom` answer "Resource not found for the
segment" against this workbook, the same way `dataValidation` and `freezePanes`
do. Hence one request per cell, ~45s to capture a tab.

Two things still need a person, per new tab: the `Resolved` dropdown and the
frozen header row. Graph refuses both against this workbook.

### Run the migration before the timer

**The code change and the migration go together.** `Missing Info` and
`Ineligible & Inactive` are no longer tab names as far as this code is concerned,
so between deploying and running `npm run split` the two queues are frozen: the
old tabs are never read, the monthly tabs do not exist yet, and every row bound
for them is skipped with `apply.destination_unavailable`.

Nothing is lost — the rows sit untouched on the old tabs and the reconciler is
state-based, so it picks up exactly where it left off once the tabs exist. But it
is a window in which the two queues stop updating, and the timer is not running
yet, so the sequence is simply: run the split, check the tabs, then enable it.

## The workbook is fully reconciled

A full cycle now plans **nothing**: zero stamps, zero appends, zero write-backs,
zero removals. Every patient a keyword routes to a queue has a row on that queue,
and every one of those rows is linked to the daily row it came from.

    Missing Info           195 rows, 26 camps   -> split across June/July/August/Other
    Not Accepted           375 rows, 35 camps
    Ineligible & Inactive  137 rows, 18 camps   -> split across June/July/August/Other
    Verify Insurance       243 rows, 17 camps
    United Refuah           85 rows, 13 camps
    ----------------------------------------
                         1,035 patient rows

Those counts are from 2026-07-31, before the split. `npm run split` moves the
rows without changing any of them, so the totals per family should be unchanged
afterwards — the dry run prints the per-month breakdown to check against.

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

### Azure infrastructure — DEPLOYED 2026-07-31
Resource group `camp-workbook-sync` in eastus holds all eight resources:
Consumption plan, storage `ste24s24d74i4egcampworkb` with the state container,
Log Analytics capped at 1 GB/day, Application Insights, the Function App, and a
user-assigned managed identity.

**This section said "the code has not been published and nothing runs on a
timer" and that line went stale nine minutes after it was written.** It was
committed at 08:47 on 2026-07-31. The four commits that follow it — 08:56
through 09:05 — are somebody deploying: no SCM log stream on Consumption, a
deployment that registers zero functions without `EnableWorkerIndexing`,
config-zip not working on Linux Consumption, and finally the fix that makes
`func azure functionapp publish` work. Nobody came back and corrected this
paragraph.

On 2026-08-06 that cost real time. The office reported that a patient marked
`ineligible` at source was not moving to the Ineligible queue, and this file was
quoted back at them as proof that nothing could be moving anything, twice. The
actual cause was the opposite: the timer WAS running and had gone blind on five
queue tabs that had been renamed, which is why the two split queues had a
backlog and the three unrenamed ones had none. That asymmetry was visible in the
data the whole time.

**Do not trust this paragraph. Ask Azure.**

```bash
az functionapp function list -g camp-workbook-sync -n camp-workbook-sync -o table
```

`syncTimer` in that list means it is live and running every five seconds.
`SYNC_PHASE` and `SYNC_DRY_RUN` in the app settings say how much of a plan it
applies. Those two commands are the truth; a sentence in a markdown file is a
snapshot of what somebody believed at the time.

**Publishing is manual.** The GitHub Actions workflow builds and tests on every
push to `main` and has never once deployed — the `Sign in to Azure` step fails
because the OIDC secrets (`AZURE_CLIENT_ID`, `AZURE_TENANT_ID`,
`AZURE_SUBSCRIPTION_ID`) were never set on the repository, so `Deploy to the
Function App` is skipped. **Merging to `main` does not ship anything.** Until
those secrets exist, every config change — including a queue tab rename — needs
`func azure functionapp publish` by hand or it never reaches the running job.

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
