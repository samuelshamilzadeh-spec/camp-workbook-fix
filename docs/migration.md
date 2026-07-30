# Migration and cutover

The existing mirror sheets are formula output, not original data — the sheet
header says so itself. The current patient list is therefore fully regenerable
from column B on the daily sheets, and no patient row exists only on a mirror
sheet.

That reframes the risk. The danger is not losing rows. It is:

- **(a)** something elsewhere in the workbook quietly depending on those sheets
  and breaking when the formulas are removed, and
- **(b)** cutting over before the new system has been proven to produce the same
  list.

**Nothing gets deleted until a diff is clean.**

Work through these in order. Each step has a checkbox because each one is a
gate, not a suggestion.

---

## Step 0 — Resolve the open question first

- [ ] Do any staff notes or manual annotations exist anywhere today, most likely
      on the daily sheets rather than the locked mirrors?
- [ ] If yes, map them into the new notes sheet keyed by `SyncID`.

This is the one part of the migration that is **not regenerable**. Everything
else can be rebuilt from column B; a note somebody typed cannot. Resolve it
before Step 1, not during.

---

## Step 1 — Preserve, before touching anything

- [ ] Take a full copy of the workbook. Keep it inside the same SharePoint site
      or another BAA-covered location.
- [ ] Confirm SharePoint version history is enabled. **Record the restore point
      by timestamp here:** `________________`
- [ ] Convert each mirror sheet to static values in place (copy → paste values).
- [ ] Rename each as a dated archive, e.g. `ARCHIVE Missing Info 2026-07-30`.
      These become frozen historical records, never edited or synced again.
- [ ] Confirm the archived sheets carry no live formulas afterward. A single
      surviving formula defeats the purpose. (Ctrl+` toggles formula view; or
      Find → `=` with "Look in: Formulas".)
- [ ] Set the deletion date for any test copy created, now rather than later.

---

## Step 2 — Dependency audit

**This is the step most likely to produce an ugly surprise.**

Before removing any formula, find everything in the workbook that references the
mirror sheet names. Produce a written list — use
[`dependency-audit.md`](dependency-audit.md), which enumerates what to search and
where.

Anything on that list needs a decision **before** cutover, or it becomes a
`#REF!` the morning after.

- [ ] Formulas on other sheets
- [ ] Named ranges
- [ ] Conditional formatting rules
- [ ] Data validation lists
- [ ] Charts
- [ ] Pivot tables
- [ ] The `United Refresh` tab
- [ ] The `Claude Log` tab
- [ ] Any other existing tab
- [ ] Every item above has a written decision

---

## Step 3 — Parallel run

Build the new queue sheets under **temporary names** so both systems run at once.

- [ ] Create the four new sheets under temporary names.
- [ ] Point the Function at them (`queueSheetNames` on `runCycle`, or rename
      after the gate — see `src/sync/cycle.ts`).
- [ ] Leave the old formula sheets working.

Do not remove the old ones yet, even though they are the source of the slowness.
A few more days of lag is cheaper than a bad cutover.

---

## Step 4 — Reconciliation gate

Write an **automated diff**, not an eyeball check. For each of the four statuses
it compares old versus new on:

- [ ] the exact patient set
- [ ] the camp each patient is grouped under
- [ ] the per-camp counts
- [ ] which fields are flagged red

Cut over only after the diff is clean **several consecutive days**, including at
least one day where the pipeline added new rows overnight.

**Any discrepancy is a finding, not noise.** A patient the old formulas caught
and the new job missed almost certainly means a keyword variant the substring
matching does not handle — check the `status.unrecognized` counts and the
`UNRECOGNIZED:` lines from `npm run inspect` first.

---

## Step 5 — Cutover, in this order

1. [ ] Clear the formulas from the old mirror sheets, but leave the sheets in
       place. Watch for breakage for a day.
2. [ ] Rename the new sheets to the canonical names.
3. [ ] Hide or archive the old sheets.
4. [ ] Delete them only later, once nothing has complained.

---

## Step 6 — Rollback

Know the exact path back **before** starting. Rollback must not depend on the
Function being healthy.

1. Disable the timer trigger with a single app setting:
   `AzureWebJobs.syncTimer.Disabled = true`. No deploy required.
2. Restore the workbook from SharePoint version history, using the timestamp
   recorded in Step 1.
3. The dated archive sheets remain as an independent record regardless.

- [ ] Everyone involved in the cutover knows all three, before Step 1 begins.

---

## Security, throughout

Every copy, backup and test file contains patient data and carries the same
obligations as the original.

- Keep copies inside the same SharePoint site or another BAA-covered location.
- Do not email the workbook.
- Do not park a copy in a personal Downloads folder or consumer cloud storage.
- Do not upload it to any third-party tool for inspection.
- If a test copy is created, plan its deletion date at the same time it is
  created.
