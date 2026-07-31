# Handoff prompt

Paste the block below into a new Claude Code session.

---

I'm continuing work on an Azure Function that syncs patient queue sheets in an
Excel workbook on SharePoint. The repo is
`samuelshamilzadeh-spec/camp-workbook-fix`, branch `main`, everything committed
and pushed.

**Read these first, in order — they contain everything learned from inspecting
the live workbook, and most of it contradicts the original build brief:**

- `docs/status.md` — what is built and verified vs what is not
- `docs/inspection-2026-07-30.md` — the real workbook layout
- `docs/column-b-values-2026-07-30.md` — all 111 status values and their routing
- `docs/phase2-decisions.md` — the rules Phase 2 must follow
- `README.md` — architecture and the phase gates

## Where things are

Phase 1 (read-only reconciliation) is complete and verified against the live
workbook: ~1s per cycle, identical output across repeated runs, zero
unrecognized status values, 1,058 patients routed across 5 queues.

Phase 2a (adoption) is **done and written to the live workbook**: 679 of 681
existing queue rows now share a SyncID with their source row on the daily
sheet. Column BA on every sheet. Re-running finds nothing to do.

**Next task: Phase 2b — append the ~380 patients who have no queue row yet.**

The reconciler already produces `append-queue-row` intents. Nothing applies
them. Still to build: appending rows in camp order, the `Camp - N patients`
dividers and a `TOTAL - N patients` row, dark red shading on blank required
fields, and whole-row inserts so columns never fall out of alignment.

Suggested approach: run it against `United Refuah` first — empty, only 85 rows,
append-only so nothing writes back — then the other four.

## Setup

`.env` is gitignored, so recreate it:

```
AZURE_TENANT_ID=dc8b1ae8-7cee-4334-9e4d-834c69fc74b5
AZURE_CLIENT_ID=27e61a22-9b04-4124-a5e0-5b691f0435c4
AZURE_CLIENT_SECRET=<generate a fresh one, see below>
GRAPH_DRIVE_ID=b!ilsgZfRGDUiGfnymDKkNsICgfZAbKLBOqB0osAi0QTHRP5x43wTNRqnEcm-FuUlo
GRAPH_ITEM_ID=01OFOCIMVCXG6SOACJIFEIWGQIMX7OP3EQ
SYNC_PHASE=1
SYNC_DRY_RUN=true
SYNC_LAYOUT_VERIFIED=true
```

Fresh secret (earlier ones leaked into a chat transcript and should be deleted):

```bash
az ad app credential reset --id 27e61a22-9b04-4124-a5e0-5b691f0435c4 \
  --years 1 --append --display-name claude --query password -o tsv
```

Useful commands: `npm test`, `npm run reconcile` (Phase 1 plan),
`npm run adopt` (adoption dry run), `npm run inspect`.

## Things that cost hours to find — do not rediscover them

**Always use a workbook session.** Session-less reads are 30-40x slower (8s vs
200ms per call) and return stale data after a write. A full 46-sheet scan is
0.7s in-session and 66s without. The session is held on the `Workbook` instance
and must outlive the cycle.

**Build worksheet URLs with `worksheetPath()`.** Writing
`/worksheets/${encodeSheet(name)}` produces `/worksheets/('Name')`, which Graph
rejects with a 400 "Empty segment" that reads like a workbook problem. Every
read and write went through this bug for hours.

**Dates are stored two ways.** Queue tabs hold Excel serials (`40147`), daily
sheets hold text (`11/30/2009`). Compare with `toDateKey`, never as strings.

**Source Row pointers drift.** During one 90-minute window, staff inserted a row
on one daily sheet and deleted one on another, and 43 queue pointers silently
started aiming at the neighbouring patient. Adoption verifies identity and falls
back to searching by name + DOB. Never trust a stored row number.

**`needs ohi` / `needs lasante` are NOT done.** 1,268 rows. The brief says `ohi`
anywhere means terminal, which marks every one of them as complete when they are
the rows still waiting to be entered.

**Graph requests need a timeout.** One hung for 31 minutes with no response.
30s abort is in place; do not remove it.

## Open questions for the office

- `lasante-e` vs `lasante-o` — assumed both mean done, initial is irrelevant.
  2,043 rows depend on this.
- 72 patients whose repeat visits disagree on an already-filled field (36 on
  phone number). Pre-existing data, not caused by the sync.
- `Not Accepted ` row 262 has no date of birth, so it cannot be matched
  confidently. Adding the DOB links it automatically.

## Not done at all

No Azure infrastructure exists — no Function App, storage, or Application
Insights. **Nothing runs on a timer.** Every result so far came from running the
code by hand.

The concurrency test was attempted twice and never produced a real answer: the
first run hung, the second ran clean but no staff edits happened during the
window, so there was no contention to observe. Phase 2b is the write where it
matters.
