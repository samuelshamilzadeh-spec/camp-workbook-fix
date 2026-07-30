# Column B (EMR) — every value in the live workbook, 2026-07-30

Extracted from all 46 daily sheets: **111 distinct values across 5,025 rows**.
Lowercased, whitespace collapsed, runs of 4+ digits masked.

This exists because the brief's vocabulary and the office's vocabulary are not
the same, and the gaps are not visible without counting.

## The finding that matters: `needs ohi` / `needs lasante`

| | rows |
|---|---|
| `needs ohi` | 709 |
| `needs lasante` | 473 |
| `needs lasante- under esti` | 66 |
| `need ohi` | 10 |
| 6 further variants | 10 |
| **total** | **1,268** |

The brief's precedence rule is "`ohi` or `lasante` present, anywhere in the
value, means terminal — no queue row, regardless of what else the cell
contains." `needs ohi` contains `ohi`, so that rule marks **1,268 patients who
still need entering into an EMR as already done**.

Substring matching is right for suffixes (`ohi - esti` is still an OHI) and
wrong for prefixes that negate. A negation check has to run before the terminal
check.

This was found by Phase 1 read-only. With writes enabled it would have produced
no error and no failure — just 1,268 patients quietly never queued.

## Groups

| Group | Rows | Variants | Destination |
|---|---|---|---|
| `lasante-*` | 2308 | 60 | terminal (done) |
| **`needs ohi` / `needs lasante`** | **1268** | **10** | **NOT done — open question** |
| `ohi*` | 380 | 7 | terminal (done) |
| `dont accept`, `dont bill` | 366 | 3 | Not Accepted |
| `verify insurance*` | 269 | 3 | Verify Insurance |
| `missing info` | 179 | 1 | Missing Info |
| `ineligible*`, `inactive` | 133 | 9 | Ineligible & Inactive |
| `united refuah*` | 85 | 3 | United Refuah — its own sheet |
| insurance wording | 21 | 11 | Not Accepted (confirmed by the office) |
| `skip` | 11 | 1 | not billable — open question |
| unclassified | 5 | 3 | open question |

## Confirmed by the office, 2026-07-30

- Both `not accepted` and `dont accept` are used. Keep both.
- **`united refuah` is its own category** with its own sheet. Rows are copied
  there and never change — no write-back, nothing sent to the source.
- **No-insurance wording goes to Not Accepted** — `no insurance on file`,
  `need insurance`, `invalid ins`, and the rest of that group.
- `Missing Info (New)` is the live tab. `Not Accepted ` (trailing space) is the
  live tab.
- **Ignore** `Missing Ins info`, `Missing info 25`, `Dont Take Ins (old)`.
- **The Verify Insurance tab does not exist yet** and needs creating, despite
  269 rows needing it.
- A blank column B means the patient has not been processed yet. Staff work
  forward from the start of the month, so recent days are mostly blank. Not a
  data quality problem.
- `skip` means a successful follow-up visit that cannot be billed.

## Resolved by the office, 2026-07-30 (second pass)

| Value | Rows | Decision |
|---|---|---|
| `needs ohi` / `needs lasante` | 1268 | **Leave alone.** No queue row. Classified `pending`, never `terminal`. |
| `skip` | 11 | **Nothing at all.** No queue row, nothing moved. Classified `ignored`. |
| `not on campium` / `not on campflow` | 4 | **Missing Info** |
| no-insurance wording | 21 | **Not Accepted** |
| `-pa` suffix (~120) | — | No effect on routing; the base keyword decides. |

`pending` and `ignored` both produce no queue row, so they behave like
`terminal` for queue purposes. They are separate outcome kinds anyway, because
collapsing them is what hid the 1,268-row bug, and because a count of
outstanding work is worth having.

## Rows flagged for manual review

`same w/ line NN` — one row, meaning unknown:

- **`July 5, 2026` row 49**, camp `bnos naale`

While locating it, the campium/campflow rows are:

- `July 12, 2026` row 187, camp `tal` — `not on campflow`
- `July 15, 2026` rows 67 and 68, camp `bnos naale` — `not on campium`
- `July 16, 2026` row 33, camp `bnos naale` — `not on campium`

## Still open

1. `lasante-e` (1216) vs `lasante-o` (827) — reads as Esti / Osnat, given
   `lasante under esti` and `lasante under osnat` also appear. Assumed for now
   that any `lasante` simply means done and the initial does not affect routing.
2. The `Verify Insurance` tab still does not exist. 269 rows need it.
