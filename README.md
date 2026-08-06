# camp-workbook-fix

An Azure Function that keeps four work-queue worksheets in sync with the daily
patient visit sheets inside one Excel workbook on SharePoint, replacing the
formula-based mirror sheets that made the workbook unusably slow.

**It writes plain static values. It never writes a formula.** That is enforced at
the single chokepoint every write passes through
(`assertNoFormulas` in `src/graph/workbook.ts`), not left to reviewer discipline.

---

## Status: reconciled, not yet automated

| Phase | What it does | State |
|---|---|---|
| 0 | Bulk-stamp `SyncID`s up front | **Removed.** Superseded by adoption plus lazy stamping |
| 1 | Read-only. Authenticate, read, log what it *would* change. Zero writes | Built — this is what the timer runs today |
| 2a | Give the existing queue rows the same `SyncID` as their source row | Done — every linkable row |
| 2b | Append the patients who have no queue row yet | Done — 1,035 rows across five tabs |
| 2c | The `Resolved` column, real dates, and the house style | Done — `npm run migrate` |
| 3 | Write-back of staff edits to the daily sheets | Done — requires the `Resolved` marker |
| 4 | Clearing and row removal | Done — `npm run resolve-rows` |
| 5 | Notes sheet | Not started |
| — | Split `Missing Info` and `Ineligible & Inactive` into one tab per month | Built — `npm run split` |

A full cycle against the live workbook now plans **nothing** — every patient a
keyword routes to a queue has a row there, linked to the daily row it came from.
`runCycle` applies a plan according to `SYNC_PHASE`, behind a second independent
gate (`SYNC_LAYOUT_VERIFIED`).

**Nothing runs on a timer.** The Azure resources are written
([`infra/main.bicep`](infra/main.bicep)) and have never been deployed.

**Nothing reaches a daily sheet without a human saying so.** The reconciler
cannot tell a staff edit from a stale mirror value — both are "the queue differs
from the source", and ~700 rows were adopted from mirrors that stopped updating
on 2026-07-30, so on those the daily sheet is the newer copy. A write-back
requires the `Resolved` marker; an unmarked difference is reported and left
alone.

**Dry-run mode is permanent, not a Phase 1 scaffold.** `SYNC_DRY_RUN=true` keeps
working at every phase.

---

## The layout, verified

`LAYOUT` in `src/config.ts` began as guesses from the build brief. All of them
have now been checked against the live workbook, and several were wrong:

| Assumption | The brief said | The workbook says |
|---|---|---|
| Daily sheet naming | a date, format unknown | `July 30, 2026` — long month, comma |
| Daily header / first data row | 1 / 2 | 1 / 2 ✓ |
| Camp name column | C, "believed to be" | C, headed `CAMP NAME` ✓ |
| Daily field columns | D–M | D–R, all of them carried ✓ |
| **Queue header row** | **9, under an instruction block** | **1. There is no instruction block** |
| `SyncID` column | BA | BA on every sheet ✓ |
| Queue tab names | four canonical names | none of them matched; see the tab map |

The queue header row is the one that cost patients: reading from row 10 meant
rows 2-9 of every mirror tab were never read at all, hiding 18 real patients from
Phase 1 and from adoption. `detectQueueShape` now reads the header off the sheet
and the configured values are only a fallback.

`SYNC_LAYOUT_VERIFIED` gates every write path and is now `true`. Read paths never
needed it, so `inspect` and a dry run remain the way to check anything.

```bash
npm run inspect                       # what the workbook actually contains
npm run reconcile                     # one cycle, prints the plan
npm run reconcile -- --force          # ignore the checkpoint, do the work anyway
npm run adopt                         # link existing queue rows        (dry run)
npm run append -- "United Refuah"     # append missing rows             (dry run)
npm run migrate -- "Missing Info"     # schema + house style + redraw   (dry run)
npm run split -- "Missing Info"       # cut a queue into monthly tabs   (dry run)
npm run resolve-rows                  # write back, clear, remove       (dry run)
npm run concurrency-test -- --sheet=... --cell=BZ1
```

Every write script is dry run by default and takes `--apply`. `npm run append`
additionally refuses to open gaps in a live sheet without `--allow-inserts`.

### One queue, several tabs

The office asked to work `Missing Info` and `Ineligible & Inactive` a month at a
time, so each is now four tabs:

    Missing Info - June           Ineligible & Inactive - June
    Missing Info - July           Ineligible & Inactive - July
    Missing Info - August         Ineligible & Inactive - August
    Missing Info - Other          Ineligible & Inactive - Other

`Verify Insurance`, `Not Accepted` and `United Refuah` stay as single tabs.

This is why a **queue family** and a **queue tab** are different types in
`src/config.ts`. A column B keyword can only tell you *what kind of problem* a
row is — that is the family, and `status.ts` is unchanged. Which tab it lands on
takes a second input, the month of the visit, and `domain/segments.ts` is the one
place the two are combined. Keeping them apart is what let the split happen
without touching a single keyword rule.

A few consequences worth knowing:

- **The month comes from the visit, so a row never moves between monthly tabs.**
  It is read from the daily sheet's own name (`July 30, 2026`), which is the
  thing the row is a record of. A patient's visit date does not change.
- **`Other` is a destination, not a failure.** A visit outside those months, or a
  `Date of Visit` that will not parse — the live workbook has one — lands there.
  A row that cannot be placed must still end up somewhere a person can see it.
- **Wrong month is just wrong queue.** A row on the June tab for a July visit
  takes the same append-then-remove path as a row whose keyword changed. There is
  no second notion of "wrong tab".
- **The year is not in the tab name.** Right for a workbook holding one season,
  wrong the moment a second is added — July 2027 would land beside July 2026. See
  `MONTH_SEGMENTS`, which also carries the 31-character limit Excel puts on a
  sheet name (`Ineligible & Inactive - September` does not fit).

`npm run split -- "Missing Info"` performs the migration, and **it copies the
colours rather than re-deriving them**. That distinction was not obvious and cost
a rewrite. The first version regenerated fills from `REQUIRED_FIELDS` and
`planQueueStyle`, reasoning that every fill on these tabs came from those rules.
Auditing the live tab disproved it: `Missing Info` carries **17 distinct fill
colours**, only four of which this codebase knows about. Amber and yellow in
`Date of Visit`, blues in `Source Row`, red note-cells reading `Staff kid` and
`Not in campminder`, greens for confirmed insurance, solid orange down
`Medications` and purple down `Allergies` — a colour-coding system staff built by
hand. Regenerating would have erased all of it. Font colour is copied too, for
the same reason: a `Medicaid #` reading `inactive` in red is somebody saying
something, and the house style paints the body black.

So the script reads every cell's fill and font colour off the source, moves the
rows, and replays those colours onto the row each patient landed on. Nothing is
added — a blank required field that nobody shaded stays unshaded. `planQueueStyle`
still supplies fonts, borders, column widths and number formats, which it
reproduces exactly; only its fill operations are dropped.

Excel's own copy would be better than any of this, and Graph refuses it. Both
`worksheets/copy` and `range/copyFrom` answer *"Resource not found for the
segment"* against this workbook — the same limitation as `dataValidation` and
`freezePanes`. Hence one request per cell, which is why the capture takes ~45s
per tab.

It is idempotent, refuses rather than silently drops if a column past the schema
holds data, proves by SyncID that each row landed where it thinks before copying
any colour onto it, and never deletes: `--archive` renames the original to
`Missing Info (old)` and hides it, still holding every row exactly as it stood.

Each new tab needs the same two settings Graph refuses to make — the `Resolved`
dropdown and a frozen header row. The script prints the reminder.

**Run it before enabling the timer.** `Missing Info` and `Ineligible & Inactive`
are no longer tab names to this code, so between deploying and running the split
those two queues are frozen: the old tabs are never read, the monthly ones do not
exist yet, and rows bound for them are skipped with
`apply.destination_unavailable`. Nothing is lost — the rows sit untouched and the
reconciler is state-based, so it resumes the moment the tabs exist — but the gap
is real, and closing it is one command.

### Open questions this codebase could not resolve on its own

1. **Precedence within a cell.** The brief says "the first queued keyword found
   wins" without saying whether "first" means first in the keyword table or first
   by position in the cell. `src/domain/status.ts` implements *table order*,
   because it is deterministic and independent of phrasing. Every cell where the
   two readings could disagree is logged as `status.multi_match`, so the choice
   is reviewable. If the office wants position order, it is a four-line change.
2. ~~**"Clear column B on the queue sheet."**~~ **Settled.** The office asked for
   an explicit `Resolved` column, which is what requirement 4 was reaching for.
   The old reading — a row counts as cleared when every field is blank — is still
   computed but is **never applied**: a row blank because somebody is midway
   through retyping it looks identical to one they meant to delete.
3. **Values that start with `+`.** Write-back sends `+1 555 …` to Excel as a
   string, and Excel treats a leading `+` as formula-ish. A leading `=` is
   rejected outright; `+` is allowed, because refusing to sync a phone number
   would be worse. Still unconfirmed — no write-back has yet carried one.
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
npm test

# Open the workbook in a browser, copy the address bar URL, then:
npm run resolve -- "https://premierassist.sharepoint.com/sites/.../Workbook.xlsx"
# prints GRAPH_DRIVE_ID / GRAPH_ITEM_ID to paste into .env, alongside the
# tenant, client id and secret

npm run inspect
```

`resolve` exists because the file's location is not obvious from a site name.
`GRAPH_DRIVE_ID` addresses the workbook wherever it actually lives — a site's
default document library, a second library on the same site, a Teams channel, or
someone's OneDrive. `GRAPH_SITE_ID` alone only reaches the site's *default*
library, which fails silently if the workbook is anywhere else.

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
            ├─ read the eleven queue tabs' usedRange
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
    queueSheets.ts      queue sheet parsing, shape detection, camp groups
    segments.ts         which monthly tab a row belongs on, from its visit date
    reconcile.ts        state-based reconciliation -> intents
    compare.ts          does a value MEAN something different? the write-back gate
    dates.ts            Excel serials, both directions
    adopt.ts            linking the pre-SyncID queue rows to their source rows
    append.ts           Phase 2b: appends -> ordered workbook operations
    style.ts            the house style, as operations
  graph/
    auth.ts             app-only credential, cached token
    client.ts           retries: 423/409 locks, 429 throttling, jittered backoff
    workbook.ts         Workbook API wrapper; the no-formulas chokepoint
  state/store.ts        checkpoint + loop-guard marker (blob, or a file locally)
  sync/
    cycle.ts            one reconciliation cycle
    apply.ts            applying a plan, gated by phase
    removals.ts         the only code here that destroys data
    context.ts          per-process wiring
  functions/syncTimer.ts  the 5-second timer trigger
scripts/
  inspect-workbook.ts   answers the pre-build questions, read-only
  run-reconcile.ts      one cycle from the CLI, prints the plan
  adopt-apply.ts        Phase 2a: link pre-SyncID queue rows
  append-queue.ts       Phase 2b: append missing rows, one queue per run
  migrate-queue.ts      Phase 2c: schema, dates, house style, redraw
  resolve-queue.ts      Phases 3 and 4: write back, clear, remove
  create-queue-sheet.ts adds a missing queue tab and its header row
  concurrency-test.ts   the write experiment
infra/
  main.bicep            everything this has never run on
```
