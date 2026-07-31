# Where this project stands

Last updated 2026-07-31, after a full pre-merge sweep of the branch.

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

## What the pre-merge sweep found

Four defects, none of which any test caught, all fixed on this branch.

| | |
|---|---|
| **A retried row delete removes a second patient** | `insertRows` and `deleteRows` are POSTs, and 500/502/503/504 and a client-side timeout were all on the retry list. None of those means the operation did not happen — a delete that committed and then timed out, retried, takes whoever moved up into that row. The identity check in `applyRemovals` runs above the retry and has already passed by then. A write against this workbook was observed hanging for 31 minutes, so this was reachable. Those two calls now retry only on 409/423/429, which are refusals that never touched the file. |
| **New camp blocks went undressed at the phase we go live in** | Recounting the dividers and styling a camp block the cycle just created were written below the phase 3 and phase 4 early returns, so they only ran at phase 4. Appending starts at **phase 2**, which is the setting the cutover uses. The first camp to arrive would have got a bare label between banded ones — exactly what de71c6b was written to prevent. Phases 3 and 4 are conditional blocks now and the function has one exit. |
| **A typo'd number setting degraded silently** | `Number('ninety')` is NaN and NaN does not throw. `SYNC_MAX_SHEETS_PER_CYCLE` made `daily.length <= NaN` false, switching the scan to tiering; `SYNC_COLD_BATCH_SIZE` then took NaN sheets per rotation, so the cold pool was never read and any sheet outside the hot window stopped reconciling. Nothing logged. All five numeric settings are now range-checked at startup, as `SYNC_PHASE` always was. |
| **The checkpoint could silently stop persisting in Azure** | With no storage connection string the state store fell back to a local file. A Consumption app's filesystem is read-only, so every write throws, `lastSelfWriteModified` never persists, and the loop guard stops working — the job then treats its own writes as staff changes every five seconds. Easy to reach by switching `AzureWebJobsStorage` to the identity-based `__accountName` form. It now fails loudly at startup instead. |

Smaller: `cleared` and `styled` were counted and then dropped before the log line;
a configuration error thrown while building the sync context escaped before any
logger existed, so an app failing every invocation said nothing.

### Code that was written and then overtaken

Removed, all recoverable from git history:

- **`src/domain/patient.ts`** and its 23 tests — cross-visit propagation, "a fix
  typed once reaches every visit". Superseded by the rule that nothing reaches a
  daily sheet without a human marking `Resolved`. Imported by nothing but its own
  test. The open question about 72 patients whose repeat visits disagree is still
  open; this module was not answering it.
- `groupAppends` — append batching, which ended up in `planQueueAppend`.
- `planBlankRequiredShading` and the `shadeBlankRequired` option that was
  declared on `PlanQueueStyleInput` and never read by the function.
- `RESOLVED_DROPDOWN_VALUE` imported into `style.ts` and never used — left over
  from setting the dropdown through Graph, which 400s.
- `LAYOUT.controlSheetName` — the loop guard was going to be a marker on a hidden
  `_SyncControl` sheet before it became `lastSelfWriteModified` in the blob
  checkpoint. `notesSheetName` went the same way; Phase 5 can declare its own.
- `IGNORED_TABS`, duplicating `knownNonDailySheets`, which is the list actually
  consulted.
- `queueColumnLetter`, `Workbook.clearFill`, `resetSyncContext`,
  `SyncState.lastSelfWriteBy` — no callers.

`lastSeenETag`, `lastScannedSheets` and `lastFullCycleAt` are written to the
checkpoint and never read back. Kept deliberately: they cost nothing and they are
what you look at when a cycle misbehaves.

## Not built

### Phase 5 — the notes sheet
Not started. Adding `Notes` shifts every column right on all five tabs again;
`scripts/migrate-queue.ts` is now the tool for that and its guard is structural,
so it can be pointed at the job.

### Azure infrastructure — DEPLOYED 2026-07-31, publishing not yet confirmed
Resource group `camp-workbook-sync` in eastus holds all eight resources:
Consumption plan, storage `ste24s24d74i4egcampworkb` with the state container,
Log Analytics capped at 1 GB/day, Application Insights, the Function App, and a
user-assigned managed identity.

**Nothing runs on a timer yet.** Publishing the code by hand ran into four
separate Linux Consumption traps, each of which presents as a healthy deployment
that does nothing — `config-zip` setting `WEBSITE_RUN_FROM_PACKAGE=1` (which that
plan does not accept), a missing `AzureWebJobsFeatureFlags=EnableWorkerIndexing`,
`func` unable to identify the project without a `local.settings.json`, and
`az webapp log tail` 404ing because there is no SCM log stream. All four are
written up in [`infra/README.md`](../infra/README.md).

Whether the publish finally landed is not recorded anywhere and cannot be read
off this repository. Check it:

```bash
az functionapp function list -g camp-workbook-sync -n "$APP" -o table
```

`syncTimer` should be listed. An empty table is the failure to expect.

Once this branch is on `main`, `.github/workflows/deploy.yml` deploys on every
push to `main` over OIDC, which is the route to prefer over publishing by hand.

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
  four have now appeared in chat transcripts — treat every one of them as
  compromised. `az ad app credential list --id <app-id> -o table`, then
  `az ad app credential delete --id <app-id> --key-id <keyId>` for each old key
  after the replacement is in place. The Function App does not need one at all
  once the identity has its Graph grant.

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
