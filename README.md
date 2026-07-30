# camp-workbook-fix

An Azure Function that keeps four work-queue worksheets in sync with the daily
patient visit sheets inside one Excel workbook on SharePoint, replacing the
formula-based mirror sheets that made the workbook unusably slow.

**It writes plain static values. It never writes a formula.** That is enforced at
the single chokepoint every write passes through
(`assertNoFormulas` in `src/graph/workbook.ts`), not left to reviewer discipline.

---

## Status: Phase 1, dry run

| Phase | What it does | State |
|---|---|---|
| 0 | Add the `SyncID` column, stamp only rows currently carrying a queued keyword | Built — `npm run backfill`, dry run by default |
| 1 | Read-only. Authenticate, read, log what it *would* change. Zero writes | Built — this is what the timer runs today |
| 2 | Populate the queue sheets | Not started |
| 3 | Write-back of staff edits to the daily sheets | Not started |
| 4 | Clearing and row removal | Not started |
| 5 | Notes sheet | Not started |

The reconciler already computes the *intents* for phases 2 through 4 — appends,
write-backs, removals — because computing them is read-only and it is exactly
what Phase 1 is supposed to log. Nothing applies them. `runCycle` throws if it is
asked to write above phase 1, and there is a second, independent gate
(`SYNC_LAYOUT_VERIFIED`) described below.

**Dry-run mode is permanent, not a Phase 1 scaffold.** `SYNC_DRY_RUN=true` keeps
working at every phase.

---

## Before this can do anything: the workbook has not been inspected

Everything in the `LAYOUT` block of `src/config.ts` is a guess taken from the
build brief and marked `UNVERIFIED`. The guesses that matter:

| Assumption | Current guess | How to confirm |
|---|---|---|
| Daily sheet naming | a date, `2026-07-30` or `7/30/2026` | `npm run inspect` prints every sheet name |
| Season size | ~60 sheets, so every sheet is read every cycle | `inspect` prints the total and the resulting scan mode |
| How far ahead the office pre-creates sheets | 14 days (only matters above the full-scan threshold) | ask the office; `inspect` shows the furthest future sheet |
| Daily header row / first data row | row 1 / row 2 | `inspect` guesses the header row per sheet |
| Camp name column | C ("believed to be", per the brief) | `inspect` prints the distinct values in C |
| Whether camp names need normalizing | casing and whitespace only | `inspect` prints them; look for aliases |
| `SyncID` column | BA on every sheet | `inspect` prints the widest used column and every named range |
| Daily-sheet field columns (D–M) | see `fieldColumns` | read the header row `inspect` prints |
| Required fields per status | modelled per status, currently near-identical | ask the office |

Until someone checks these and sets `SYNC_LAYOUT_VERIFIED=true`, **every write
path refuses to run** — the Phase 0 backfill `--apply`, and any future applier.
Read paths work fine, so `inspect` and the Phase 1 dry run are how you do the
checking.

```bash
npm run inspect                 # answers the "before you write code" questions
npm run reconcile               # one Phase 1 cycle, prints the plan
npm run reconcile -- --force    # ignore the checkpoint and do the work anyway
```

### Open questions this codebase could not resolve on its own

1. **Precedence within a cell.** The brief says "the first queued keyword found
   wins" without saying whether "first" means first in the keyword table or first
   by position in the cell. `src/domain/status.ts` implements *table order*,
   because it is deterministic and independent of phrasing. Every cell where the
   two readings could disagree is logged as `status.multi_match`, so the choice
   is reviewable. If the office wants position order, it is a four-line change.
2. **"Clear column B on the queue sheet."** Requirement 4 says staff signal
   removal by clearing column B of the *queue* sheet — but the documented queue
   layout has `Source Row` in column B and no status column at all. Until that is
   settled, `reconcile.ts` uses the conservative reading: a queue row counts as
   cleared only when every patient field is blank. That cannot misfire, whereas
   keying deletion off one column would delete a row on a stray backspace.
3. **Values that start with `+`.** Write-back sends `+1 555 …` to Excel as a
   string, and Excel treats a leading `+` as formula-ish. A leading `=` is
   rejected outright; `+` is allowed, because refusing to sync a phone number
   would be worse. Confirm what Excel actually stores during the concurrency
   test, before Phase 3.
4. **Existing staff notes.** The brief's own open question — whether manual
   annotations exist anywhere today, most likely on the daily sheets. If they do
   they need mapping into the notes sheet keyed by `SyncID`, and that mapping is
   the one part of this migration that is not regenerable. Resolve before Step 1
   of the cutover.

---

## Local development

Requires Node 20+ and, for running the Function locally, [Azure Functions Core
Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local).

```bash
npm ci
cp .env.example .env                                  # for the scripts
cp local.settings.json.example local.settings.json    # for `func start`
# fill in tenant/client/secret and the site + item ids in both
npm test
npm run inspect
```

Neither `.env` nor `local.settings.json` is committed — both hold a client
secret. `.gitignore` also blocks `*.xlsx` and `*.csv`, so a workbook copy cannot
land in the repo by accident.

To run the Function locally:

```bash
npm start   # cleans, builds, then `func start`
```

### Handling workbook copies

Every copy, backup and test file carries the same obligations as the original.
Keep copies inside the same SharePoint site or another location covered by the
organisation's BAA. Do not email the workbook, do not leave a copy in a personal
Downloads folder or consumer cloud storage, and do not upload it to any
third-party tool for inspection. When a test copy is created, set its deletion
date at the same time.

`npm run inspect` exists partly so nobody has to take a copy off SharePoint to
answer a structural question.

### Never log patient data

`src/logging.ts` redacts by allow-list: a payload key not on `SAFE_KEYS` is
replaced with `[redacted:type]`, recursively, before anything reaches the sink.
The only patient-linked identifiers that survive are the `SyncID`, the sheet name
and the row number. Graph error bodies are dropped entirely, because a failed
range write echoes the payload back. There are tests for this.

If you add a log line, add its key to `SAFE_KEYS` deliberately or it will be
redacted — the default is safe.

---

## What to ask the Azure admin for

Two things.

### 1. An Entra app registration with `Files.ReadWrite.All`

- A new app registration (suggested name: `camp-workbook-sync`).
- Microsoft Graph **application** permission `Files.ReadWrite.All`, with admin
  consent:

  ```bash
  az ad sp show --id 00000003-0000-0000-c000-000000000000 \
    --query "appRoles[?value=='Files.ReadWrite.All'].{id:id,value:value}" -o table

  az ad app permission add --id {appClientId} \
    --api 00000003-0000-0000-c000-000000000000 \
    --api-permissions {roleId}=Role

  az ad app permission admin-consent --id {appClientId}
  ```

- Either a client secret (note the expiry and diarise the rotation) **or**,
  preferred, no secret at all: assign a user-assigned managed identity to the
  Function App and grant the permission to that identity instead. The code
  supports both — set `AZURE_USE_MANAGED_IDENTITY=true`.

#### Scope decision, recorded

The build brief specified `Sites.Selected` scoped to the single SharePoint site,
and explicitly ruled out `Files.ReadWrite.All`. **That was overridden
deliberately** in favour of the broader permission, to avoid the separate
site-level grant step. This note exists so the next person reads it as a decision
rather than an oversight.

What that means in practice:

- The credential can read and write **every file in every OneDrive and SharePoint
  site in the tenant**, not just this workbook. The code only ever touches the one
  `driveItem` in `GRAPH_ITEM_ID`, but nothing at the permission layer enforces
  that.
- A leaked or mis-scoped client secret is therefore a tenant-wide incident, not a
  single-site one. Treat secret rotation and the Function App's access
  restrictions as load-bearing controls, since the permission scope no longer is.
- Prefer the managed identity over a client secret here. With this permission
  breadth, removing the secret entirely is worth more than it would have been
  under `Sites.Selected`.
- Narrowing later is a permission change and a one-line grant, with **no code
  change** — the Graph calls are identical under either permission. If the
  security posture is revisited, see `docs/scope-narrowing.md`.

`GraphClient` deliberately does not retry a 403, so an auth misconfiguration
still fails loudly on the first cycle instead of looking like a transient error.

### 2. A resource group with a Consumption-plan Function App

- Resource group in the region nearest the SharePoint tenant.
- **Function App on the Consumption (Y1) plan. Not Flex Consumption.** Flex bills
  a minimum of 1,000 ms per execution at a larger default instance size, which
  turns this 5-second polling pattern into a real monthly bill. Consumption bills
  a minimum of 128 MB × 100 ms and this workload fits inside the free grant. At
  ~17,000 invocations a day the difference is the whole cost story.
- Runtime: Node 20, Functions v4.
- A storage account (required by the Functions runtime; also holds this job's
  checkpoint blob).
- Application Insights, **with sampling enabled from the start** — already
  configured in `host.json`, but set the daily cap on the AI resource too.
  Default verbose logging at this invocation rate is the one thing here that can
  generate a surprising bill.
- App settings matching `.env.example`, minus the secret if managed identity is
  used.
- For the GitHub Actions deploy: a federated credential (OIDC) on the app
  registration for this repo, plus repo secrets `AZURE_CLIENT_ID`,
  `AZURE_TENANT_ID`, `AZURE_SUBSCRIPTION_ID` and the variable
  `AZURE_FUNCTIONAPP_NAME`. OIDC means there is no publish profile sitting in the
  repo secrets.

**Rollback lever:** disabling the timer is a single app setting —
`AzureWebJobs.syncTimer.Disabled = true`. It takes effect without a deploy and
does not depend on the Function being healthy. Everyone involved in the cutover
should know it.

---

## How a cycle works

```
timer (5s)
  └─ GET driveItem lastModifiedDateTime        ← one call; most cycles stop here
       ├─ unchanged since last cycle              → exit
       ├─ matches our own last write (loop guard) → exit
       └─ changed
            ├─ open a workbook session (non-persisted while read-only)
            ├─ read the four queue sheets' usedRange
            ├─ pick sheets to scan: all of them, at one season's size
            ├─ read those daily sheets' usedRange, 8 at a time
            ├─ reconcile → a list of intents
            └─ log the plan (Phase 1) / apply it (Phase 2+)
```

Design points worth not undoing:

- **State-based, never transition tracking.** Each cycle reads what column B says
  right now. Nothing detects "changed from X to Y". This is what makes the job
  idempotent, makes catch-up after downtime identical to the normal path, and
  makes a cleared cell need no special case. There is a test asserting that a
  second run over the same state plans nothing.
- **Lazy ID stamping.** An ID is stamped at the moment a row first enters a
  queue, in the same operation that reads it — never a bulk walk looking for
  blanks. A patient who never gets a keyword in column B never needs one, so
  `run_daily.py` needs no changes, and group headers and instruction blocks are
  excluded automatically. The rejected alternative, "stamp any row whose
  identifying fields are populated", fails on exactly this population: a Missing
  Info row is missing a field by definition.
- **Every daily sheet, every cycle.** This workbook is one camp season — a week
  of June, all of July, most of August, about 60 sheets — and 60 concurrent reads
  fit comfortably inside a five-second cycle. So a keyword typed onto any sheet,
  however old, is picked up on the next cycle. No rotation, no staleness, nothing
  to reason about.

  What makes that fit is `SYNC_READ_CONCURRENCY` (default 8). Sixty *sequential*
  `usedRange` calls at a few hundred milliseconds each is 12-24 seconds, which
  overruns the timer; eight at a time is a couple of seconds.

- **Tiering, once the workbook outgrows one season.** Add the 2027 sheets and
  it is 120; a few seasons and full scanning stops fitting. Above
  `SYNC_MAX_SHEETS_PER_CYCLE` (default 90) the scan degrades automatically
  instead of overrunning:

  | Tier | What | When |
  |---|---|---|
  | hot | a date window around today (`SYNC_HOT_DAYS_BACK` / `SYNC_HOT_DAYS_FORWARD`), plus every sheet an existing queue row points at | every cycle |
  | cold | everything else, `SYNC_COLD_BATCH_SIZE` at a time, rotating on a checkpointed cursor | over successive cycles |

  Even then nothing is excluded, only deferred, and the cycle logs
  `scan.tiering_engaged` when it crosses over — the freshness guarantee changes
  from "every sheet every cycle" to "every sheet within `sweepCycles`", which is
  worth noticing rather than discovering.

- **The office creates daily sheets in advance.** On 3 August there may already
  be sheets out to the 8th, empty until their day arrives. This is why the hot
  window reaches *forward* as well as back, and why nothing anywhere picks sheets
  by "latest date" — the latest-dated sheets are the empty future ones.

- **Reads are cheap, writes are expensive.** One `usedRange` call per sheet, then
  filter in memory. Never a targeted read per row, never a write per cell.
- **423 and 409 are expected**, not errors — a staff member has the file open in
  Excel desktop with autosave on. They are retried with exponential backoff and
  full jitter, and logged at debug. 403 is not retried.
- **Orphans are reported, never acted on.** A queue row whose `SyncID` is not
  found in the scanned daily sheets could mean a deleted source row, or simply a
  sheet outside this cycle's scan. Those are indistinguishable from inside a
  cycle, so it logs `queue.orphan_row` and does nothing.

---

## Before Phase 2: run the concurrency test

```bash
node --env-file=.env node_modules/.bin/tsx scripts/concurrency-test.ts \
  --sheet="2026-07-30" --cell=BZ1 --iterations=60 --interval=5000
```

Against a **copy**, with one person in Excel desktop and one in Excel Online,
both editing during the run. The script records the 423/409/429 rate and the
latency distribution, but the result that decides the write strategy is the one
it cannot measure: **did every edit those two people made survive?** Check the
workbook by hand afterwards. A clean script run with a lost human edit is a
failed test.

---

## Migration and cutover

See [`docs/migration.md`](docs/migration.md) for the full sequence — preserve,
dependency audit, parallel run, reconciliation gate, cutover, rollback — and
[`docs/dependency-audit.md`](docs/dependency-audit.md) for the audit worksheet.
Nothing gets deleted until the diff is clean.

---

## Layout

```
src/
  config.ts             every workbook assumption, in one place, with the verify gate
  logging.ts            allow-list redaction; no patient data leaves the process
  domain/
    status.ts           column B keyword matching and precedence
    syncId.ts           row identity
    cells.ts            A1 address arithmetic
    dailySheets.ts      daily sheet parsing, camp normalizing, bounded scan
    queueSheets.ts      queue sheet parsing and camp groups
    reconcile.ts        state-based reconciliation -> intents
  graph/
    auth.ts             app-only credential, cached token
    client.ts           retries: 423/409 locks, 429 throttling, jittered backoff
    workbook.ts         Workbook API wrapper; the no-formulas chokepoint
  state/store.ts        checkpoint + loop-guard marker (blob, or a file locally)
  sync/
    cycle.ts            one reconciliation cycle
    context.ts          per-process wiring
  functions/syncTimer.ts  the 5-second timer trigger
scripts/
  inspect-workbook.ts   answers the pre-build questions, read-only
  phase0-backfill.ts    ID stamping, dry run by default
  run-reconcile.ts      one Phase 1 cycle from the CLI
  concurrency-test.ts   the pre-Phase-2 write experiment
```
