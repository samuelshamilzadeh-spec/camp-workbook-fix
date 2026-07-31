# Phase 2 design decisions

Settled with the office 2026-07-30. Phase 2 populates the queue sheets; these
are the rules it must follow.

## 1. Camp dividers carry a running count

Each camp block is preceded by a single cell reading `Achim - 12 patients`, and
each queue sheet ends with `TOTAL - 374 patients`.

The count is a **snapshot at write time**, not a live figure. It says what the
queue held when the row was written. That is what the office asked for and what
the old mirror sheets did.

It is a static string in one cell — not a formula, not a second column. This
workbook is being rescued from a formula load; a `COUNTIF` per camp block would
start rebuilding the thing the project exists to remove.

`isGrandTotalRow` keeps the total from being parsed as a camp named "TOTAL",
which would otherwise nest a group inside it and drift the counts on every pass.

## 2. Existing queue rows are adopted, not discarded

There are ~754 rows on the live queue tabs today, from before SyncID existed.
Staff are working in them right now, so they stay. For an edit made there to
reach the daily sheet, both ends need the same ID.

`src/domain/adopt.ts` links them:

1. Read the old pointer — `Date of Visit` (an Excel serial, e.g. 46225 =
   2026-07-22) plus `Source Row` — to find a candidate daily row.
2. **Verify** by comparing Last Name, First Name and Date of Birth.
3. Adopt only on agreement: stamp the same SyncID on both ends.
4. Anything else is reported, never guessed.

Step 2 is the point of the whole thing. The build brief is blunt that the old
pointer "breaks silently the moment anyone inserts, deletes, or sorts a row",
and calls a broken pointer the primary accuracy risk in the project. A drifted
pointer still resolves to a row — just the wrong patient's — so adopting on the
pointer alone would wire one patient's queue edits into another patient's
record. Every rejection reason is reported for a human:

| Reason | Meaning |
|---|---|
| `identity-mismatch` | pointer resolved, but to a different patient |
| `too-little-identity` | nothing to compare; agreement on nothing is not agreement |
| `source-row-out-of-range` | pointer past the end of the sheet |
| `unknown-source-sheet` | Date of Visit matches no daily sheet |
| `source-already-adopted` | two queue rows claiming one source row |

A source row that already has an ID lends it rather than getting a second one.

Adoption runs once. Afterwards every row has an ID and none of this applies.

## 3. No hyperlink in Source Row

The office's reasoning: once an edit on the queue sheet flows back to the daily
sheet, there is little need to jump to the source.

`Source Row` stays as a plain number — it is still needed for adoption and for
tracing a row by hand — but nothing builds a hyperlink. One less thing to write
into a live sheet.

## 4. Rows move as whole rows

Hard constraint from the office: **one row is one patient, and nothing may break
that alignment.** Inserting a patient must shift every column of the rows below
together.

Two consequences for the implementation:

- Inserts use Excel's row-insert (shift down) on the entire row, never a write
  into a cell range that would push some columns and not others.
- A new row's 17 columns are written in a single range write, so a row is never
  momentarily half-populated where a human could see it.

## 5. Repeat visits: every visit is a row, but a fix is typed once

A patient seen twice gets two rows on the daily sheets and two on the queue —
two visits, two insurance entries, and the office wants both visible. But
resolving the same missing information twice is wasted work.

Measured on the live workbook:

    queued visits                              1,061
    distinct patients                            836
    patients appearing more than once             148   (one appears 7 times)
    visits belonging to a repeat patient          373
    duplicate effort                              21% of the queue

So two identifiers, doing different jobs:

| | Scope | Purpose |
|---|---|---|
| `SyncID` | one visit row | links a queue row to the exact daily row it came from |
| `PatientID` | one person | lets a fix fan out to that person's other visits |

**Patient key: Last Name + First Name + Date of Birth + Camp.**

Camp was the office's addition and it is free — on the live data it yields the
same 836 people as name+DOB alone, splitting nobody, while ruling out two
children who genuinely share a name and a birthday. Name alone was never safe:
14 name-pairs in this workbook already cover more than one date of birth.

A row with no surname, first name or DOB gets no key and matches nothing. Only
1 of 1,061 rows is in that state. The failure directions are deliberately
asymmetric: a typo splits one patient into two and the office does the work
twice, exactly as today; a bad merge would put one child's insurance on
another's record.

### What propagates, and what does not

A queue value that differs from its own source row is a staff edit — someone
typed it just now. That value fans out to the patient's other visits. Anything
else is left alone.

**Two visits edited to different values is not resolved.** Nothing is written
and the conflict is reported. This is not a corner case: 72 of the 148 repeat
patients already disagree on some filled field.

    36  Phone Number        12  Insurance ID #      10  State
    16  Billing Address     11  City                 5  Zip Code
    12  Insurance Carrier                            1  Medicaid #

A pre-existing disagreement where nobody has edited anything stays untouched —
it is history, not an instruction.

There are also **67 blank cells** that a sibling visit could fill. Whether to
fill those automatically is still open; it is a separate, more aggressive
behaviour than propagating an edit.

## Still open

- **Rebuild vs append — settled as append**, 2026-07-31. Nothing is rewritten.
  New rows go into their camp's existing block, or into a new block at the foot
  of the body; the divider count and the `TOTAL` line are rewritten, and no
  existing row is moved except by the row-insert that opens a gap beneath it.
  Rebuilding a tab that staff are editing was never worth the risk when adoption
  had already made appending sufficient. See [`phase2b.md`](phase2b.md).
- **Auto-filling blanks from a sibling visit** (67 cells today). Propagating an
  edit is unambiguous; back-filling from history is a judgement call.
- **The concurrency test** has not been run. It should be, before writes are
  enabled — staff are in the file continuously, which makes the test realistic
  rather than theoretical.
