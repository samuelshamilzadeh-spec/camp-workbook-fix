# Dependency audit worksheet

**Step 2 of [`migration.md`](migration.md). Complete this before removing a
single formula.**

Goal: a written list of everything in the workbook that references the four
mirror sheet names. Anything found here needs a decision before cutover, or it
becomes a `#REF!` the morning after.

Sheet names to search for (include the archived variants once Step 1 has
renamed them):

- `Verify Insurance`
- `Missing Info`
- `Not Accepted`
- `Ineligible & Inactive`

---

## Where to look

| # | Surface | How to find references | Notes |
|---|---|---|---|
| 1 | Formulas on other sheets | Find & Replace → Options → Within: **Workbook**, Look in: **Formulas**, search each sheet name (and the `'…'!` quoted form) | An `&` in a sheet name means Excel quotes it: `'Ineligible & Inactive'!A1` |
| 2 | Named ranges | Formulas → Name Manager. Also `GET /workbook/names` — `npm run inspect` prints these | Check both workbook and worksheet scope |
| 3 | Conditional formatting | Home → Conditional Formatting → Manage Rules → Show rules for: **This Worksheet**, one sheet at a time | Rules using a formula can reference another sheet indirectly |
| 4 | Data validation lists | Data → Data Validation; or Home → Find & Select → Go To Special → Data Validation, per sheet | List sources often point at a mirror sheet range |
| 5 | Charts | Click each chart series and read the `SERIES()` formula | Easy to miss on a sheet nobody opens |
| 6 | Pivot tables | PivotTable Analyze → Change Data Source, per pivot | Also check the pivot cache source |
| 7 | `United Refresh` tab | Inspect fully — formulas, queries, buttons | Named in the brief as a likely dependant |
| 8 | `Claude Log` tab | Inspect fully | Named in the brief as a likely dependant |
| 9 | Every other tab | Repeat 1–6 | Including hidden and very-hidden sheets |
| 10 | Power Query / connections | Data → Queries & Connections | A query over a mirror sheet breaks silently |
| 11 | VBA / macros | Alt+F11, search the project for the sheet names | Only if the file is `.xlsm` |
| 12 | External links | Data → Edit Links; other workbooks may reference this one | Cannot be found from inside this file alone — ask the office |

Rows 10–12 are not in the brief's list. They are added because each one breaks
the same way and none of them shows up in a formula search of the visible
sheets.

---

## Findings

Copy this table and fill it in. One row per reference found. "None found" is a
valid and useful result — record it explicitly, with who checked and when.

| # | Where (sheet / object) | References | What it does | Decision | Owner | Done |
|---|---|---|---|---|---|---|
| 1 |  |  |  |  |  | ☐ |
| 2 |  |  |  |  |  | ☐ |
| 3 |  |  |  |  |  | ☐ |

Decision options, roughly in order of preference:

- **Repoint** at the new queue sheet (usually right, once the new sheets carry
  the canonical names).
- **Repoint** at the dated `ARCHIVE …` sheet (right when the thing is historical).
- **Retire** the dependant along with the mirror sheets.
- **Rebuild** it against the notes sheet or the daily sheets.

---

## Sign-off

- [ ] All twelve surfaces checked
- [ ] Every finding has a decision and an owner
- [ ] Findings reviewed with someone from the office who uses the workbook daily

Checked by: `________________`  Date: `____________`
