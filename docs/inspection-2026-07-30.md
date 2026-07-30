# Live workbook inspection, 2026-07-30

First read of `CAMP PATIENT SIGN UP SHEET.xlsx` (3.7 MB, 58 tabs). This records
what is actually in the file, against what the build brief assumed.

Scope of the scan: 28 of 46 daily sheets read successfully, **4,102 patient
rows**. The other 18 failed with `501 OpenWorkbookBlockedWorkbook` while staff
had the workbook open — see "Reliability" below.

## Structure — the brief was right

| Thing | Brief | Actual |
|---|---|---|
| Status column | B | **B, headed `EMR`** ✓ |
| Camp column | C ("believed to be") | **C, headed `CAMP NAME`** ✓ |
| Header row | 1 | **1** ✓ |
| First data row | 2 | **2** ✓ |
| Patient fields | D–M | **D–M exactly** ✓ |

Daily sheet columns: `LABS | EMR | CAMP NAME | LAST Nm | FIRST Nm | DOB | GENDER
| BILLING ADDRESS | CITY | St | ZIP | PHONE NUMBER | INS CARRIER | INS ID # |
Medicaid # | MEDICAL HISTORY | MEDS | ALLERGIES` (A–R, 18 columns).

`INS ID #`, `Medicaid #`, `MEDICAL HISTORY`, `MEDS` and `ALLERGIES` exist on the
daily sheets but are not in the queue layout. Confirm none are needed on a queue
sheet before Phase 2.

**Sheet naming is `July 30, 2026`** — long month, no leading zero, comma before
the year. Not ISO. 46 daily sheets, 16 June through 31 July 2026.

**SyncID column:** daily sheets use A–R for values, with formatting out to AY.
`BA` is clear. Only one named range points at data — `aEditHousehold` ->
`'July 4, 2026'!$H$62`. Everything else named is an `_xlfn.*` function shim.

## Column B vocabulary — the brief was wrong

Counts across 4,102 rows:

| Value staff type | Count | Routed to |
|---|---|---|
| `lasante` | 1882 | terminal |
| `ohi` | 756 | terminal |
| `verify insurance` | 221 | Verify Insurance |
| **`dont accept`** | **214** | **Not Accepted — was missed entirely** |
| `missing info` | 105 | Missing Info |
| `ineligible` | 77 | Ineligible & Inactive |
| **`united refuah`** | **47** | **terminal — was missed entirely** |
| `inactive` | 16 | Ineligible & Inactive |
| `ineligible + inactive` | 2 | Ineligible & Inactive |

**`not accepted` — the brief's keyword — appears zero times.** Staff type
`dont accept`. Left as-is, 214 patients would never have been queued.

**`united refuah` is a third EMR**, not in the brief at all, and there is a
`United Refuah` tab in the workbook. Now terminal alongside `ohi` and `lasante`.
If that reading is wrong, 47 patients are affected — worth one confirmation.

### Still unrouted — each is a real patient in no queue

Insurance-ish, 19 rows total. Verify Insurance or Missing Info? Both defensible,
so neither was assumed:

`no insurance on file` (5), `need insurance` (3), `doesnt have insurance` (2),
`no insurance` (2), `need ins` (1), `doesnt have ins` (1),
`need insurance verification` (1), `incorrect insurance` (1), `invalid ins` (1),
`wrote paper has no ins` (1), `pt doesnt have insurance` (1)

Unclear, 13 rows: `skip` (7) — possibly deliberate exclusion —
`not on campium` (3), `not on campflow` (1), `need to confirm dob` (1),
`same w/ line NN` (1).

Handled: `inegilible` (1), a misspelling, now routes to Ineligible & Inactive.

## Camp names need case normalization

32 distinct camps. Several appear under two casings and would otherwise become
two groups:

- `elky` (133) / `Elky` (52)
- `bnos naale` (290) / `Bnos Naale` (40)
- `chayeinu` (135) / `Chayeinu` (25)

Grouping is already case-insensitive, so these merge correctly. What is not yet
decided is which spelling to *display* on the group header.

No abbreviations were detected that obviously alias to a longer name, except
`BRC` (195) — confirm whether that is its own camp.

## Queue sheets do not match the brief

The four canonical names do not exist. Present instead:

| Tab | Note |
|---|---|
| `Ineligible & Inactive` | matches; 1,221 rows |
| `Not Accepted ` | **trailing space in the name** |
| `Missing Info (New)` | not `Missing Info` |
| `Missing Ins info` | a second missing-info tab? |
| `Dont Take Ins (old)` | superseded |
| `Missing info 25` | 2025? |
| *(none)* | **no `Verify Insurance` tab at all**, despite 221 rows |

`United Refuah`, `Claude Log`, `Cheat Sheet`, `2025 Archive`, `2024 Archive` and
`_Feed` (very hidden) are the remaining tabs.

**This blocks Phase 2** and needs the office to say which tab is the live queue
for each status, and where Verify Insurance rows go today.

## Performance — resolved

Superseded by the section below. Kept for the record because the diagnosis
matters: the first numbers here were measured session-less, which is what made
the workbook look unusably slow.

## Reliability — the significant risk

Per-sheet `usedRange` read latency, 28 successful reads at concurrency 4:

- **p50 6.9 s, p95 12.4 s, max 16.0 s**

And **18 of 46 sheets returned `501 OpenWorkbookBlockedWorkbook`** while staff
had the file open. A later probe returned 501 for every sheet.

Two consequences:

1. **A 5-second timer is not viable against this workbook.** 46 sheets at ~7 s
   each is well over a minute per full cycle even at concurrency 8. The
   requirement of a 5-second poll and the requirement to scan every sheet cannot
   both hold at this file\'s current speed.
2. **`501` is not in the retry set** and is not a lock error in the 423/409
   sense. It needs handling before any write phase.

Both may improve once the mirror formulas are removed — the formula load is what
makes the file slow, and that is the project\'s whole purpose. That is a reason
to sequence the cutover carefully, not a reason to assume the problem away.


---

# Update, 2026-07-30 evening: performance resolved

Two changes landed between the measurements above and these: the office stripped
the mirror formulas and removed unused content, and the code stopped opening a
workbook session per cycle.

| | first measurement | now |
|---|---|---|
| file size | 3.70 MB | **1.87 MB** |
| session-less single sheet read | 8300 ms | 1504 ms |
| per-sheet read, p50 | 6900 ms | **94 ms** |
| full 46-sheet scan | 66 s | **0.7 s** |
| warm scan (the real per-cycle cost) | — | **0.6 s** |
| sheets failing with 501 | 18 of 46 | **0** |

**0.6 s per cycle against a 5 s budget.** The earlier conclusion — that a
five-second timer could not work against this workbook — no longer holds, and
the hot/cold tiering stays an unused fallback for when the workbook accumulates
more seasons.

Which of the two changes did what: the session accounts for most of it (66 s to
9.3 s on its own, before any cleanup), the cleanup for the rest (9.3 s to 0.7 s).
The formula load was never the cause of the read latency — the cost was the
Excel service opening the workbook once per request.

Also resolved:

- **The `400 BadRequest` on `Ineligible & Inactive` is gone.** It now reads in
  91 ms. That was the removed content, not a range-size limit.
- **`_Feed` has been deleted** — 13,801 rows, 421 live formulas, very hidden.
  Nothing broke, which answers what it drove: nothing.
- `Dont Take Ins (old)` and `Missing Ins info` are now hidden rather than
  deleted, and remain in IGNORED_TABS.

Queue tabs as they now stand:

    Ineligible & Inactive    A1:T1221
    Not Accepted (trailing space)  A1:Q400
    Missing Info (New)       A1:Q230
    United Refuah            A1:S23
    Verify Insurance         STILL ABSENT — 270 rows need it

Those four carry the old mirror layout, not the queue layout in the brief, so
Phase 2 has to decide whether to rebuild them in place or create new ones
alongside.
