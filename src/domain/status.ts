/**
 * Column B status matching.
 *
 * Rules, straight from the brief:
 *   - case-insensitive, substring-based, always. Never an exact match.
 *   - `ohi` and `lasante` are terminal successes and suppress everything else.
 *   - otherwise the first queued keyword found wins.
 *   - anything unrecognized is not queued. Never guess.
 *   - multi-matches are logged, never resolved silently.
 */

import type { QueueSheetName } from '../config';

export interface QueueKeyword {
  keyword: string;
  destination: QueueSheetName;
}

/**
 * Phrases meaning "still needs to be entered", which must be tested BEFORE the
 * terminal keywords.
 *
 * VERIFIED 2026-07-30: 1,268 rows — `needs ohi` 709, `needs lasante` 473,
 * `needs lasante- under esti` 66, `need ohi` 10, plus 6 smaller variants.
 *
 * This is the trap in the brief's precedence rule. "`ohi` or `lasante` present,
 * anywhere in the value, means terminal" is correct for suffixes — `ohi - esti`
 * is still an OHI — and catastrophically wrong for a prefix that negates it.
 * `needs ohi` contains `ohi`, so the documented rule marks every one of those
 * 1,268 patients as already done.
 *
 * The office's decision: these get no queue row. They are the office's own
 * pending state, handled by staff working forward through the month. But they
 * are emphatically not finished, and the distinction is preserved so that a
 * count of outstanding work is available and can never be mistaken for done.
 */
const PENDING_PATTERNS: readonly RegExp[] = [
  /\bneeds?\b[^a-z]{0,4}(ohi|lasante)/,
] as const;

/**
 * `skip` — a successful follow-up visit that cannot be billed for. Confirmed by
 * the office: no queue row, nothing moved, no action of any kind.
 */
const IGNORED_KEYWORDS: readonly string[] = ['skip'] as const;

/**
 * Terminal success states. Presence anywhere in the value suppresses queuing.
 *
 * VERIFIED 2026-07-30 against 4,102 rows of the live workbook:
 *   lasante 1882, ohi 756, united refuah 47.
 *
 * `united refuah` was not in the brief. It is a third EMR — it sits alongside
 * `ohi` and `lasante` in column B, and the workbook has a `United Refuah` tab.
 * Treating it as terminal is the only reading consistent with that; if it turns
 * out to mean something else, 47 patients are affected.
 */
export const TERMINAL_KEYWORDS: readonly string[] = [
  'ohi',
  'lasante',
  // Observed misspelling of lasante, 1 row.
  'lasanante',
] as const;

/**
 * `united refuah` is NOT terminal. The office confirmed it is its own category
 * with its own sheet: rows are copied across and then never change. It is
 * therefore a queue destination whose rows are append-only — see
 * QUEUE_ONLY_DESTINATIONS in config.ts.
 */

/**
 * Queued keywords in precedence order.
 *
 * OPEN QUESTION (see README): the brief says "first queued keyword found wins"
 * without saying whether "first" means first in this table or first by position
 * within the cell. This implements table order, because it is deterministic and
 * independent of how staff happen to phrase a compound value. Every cell where
 * the two readings could disagree is logged as a multi-match, so the decision is
 * reviewable rather than invisible.
 *
 * `ineligible` and `inactive` are independent keywords that happen to share a
 * destination. Either alone is sufficient. There is deliberately no combined
 * `ineligible/inactive` entry — that string is the sheet's name, not a value
 * staff type.
 */
export const QUEUE_KEYWORDS: readonly QueueKeyword[] = [
  // VERIFIED 2026-07-30 against 4,102 rows of the live workbook. Counts are
  // occurrences found in column B (EMR) across the 2026 season.

  // `dont accept` is what staff actually type — 214 rows. The brief's
  // `not accepted` appears ZERO times and matches nothing. Keeping both costs
  // nothing; dropping `dont accept` would silently unqueue 214 patients, which
  // is the single largest gap the inspection found.
  { keyword: 'dont accept', destination: 'Not Accepted' },
  { keyword: "don't accept", destination: 'Not Accepted' },
  { keyword: 'dont bill', destination: 'Not Accepted' },
  { keyword: 'not accepted', destination: 'Not Accepted' },

  { keyword: 'missing info', destination: 'Missing Info' }, // 105

  { keyword: 'ineligible', destination: 'Ineligible & Inactive' }, // 77
  // Observed misspelling, 1 row. Cheap to carry, and the alternative is a
  // patient sitting in no queue at all.
  { keyword: 'inegilible', destination: 'Ineligible & Inactive' },
  { keyword: 'inactive', destination: 'Ineligible & Inactive' }, // 16

  // --- Insurance wording, assigned by the office 2026-07-30 ----------------
  //
  // ORDER IS LOAD-BEARING HERE. These phrases contain one another, and the
  // first match in table order wins, so every rule must sit above the more
  // general rule it would otherwise be swallowed by:
  //
  //   `no insurance on file`        contains `no insurance`
  //   `need insurance verification` contains `need insurance`
  //   `need insurance`              contains `need ins`
  //   `pt doesnt have insurance`    contains `doesnt have ins`
  //
  // Get the order wrong and the distinctions the office just drew collapse
  // into whichever rule happens to be listed first — silently, since every
  // one of them is a valid destination.

  // Verification wanted -> Verify Insurance. Above `need insurance`.
  { keyword: 'need insurance verification', destination: 'Verify Insurance' },
  { keyword: 'verify insurance', destination: 'Verify Insurance' }, // 269

  // "we need the information" -> Missing Info. `no insurance on file` sits
  // above the bare `no insurance` below, which goes somewhere else entirely.
  { keyword: 'no insurance on file', destination: 'Missing Info' }, // 5
  { keyword: 'need insurance', destination: 'Missing Info' }, // 4
  { keyword: 'need ins', destination: 'Missing Info' }, // 1

  // Insurance exists but is wrong, so somebody must call and fix it.
  { keyword: 'incorrect insurance', destination: 'Ineligible & Inactive' }, // 1
  { keyword: 'invalid ins', destination: 'Ineligible & Inactive' }, // 1

  // Genuinely no insurance -> cannot bill -> Not Accepted.
  { keyword: 'no insurance', destination: 'Not Accepted' }, // 3
  { keyword: 'has no ins', destination: 'Not Accepted' }, // 1, `wrote paper has no ins`
  // Catches `doesnt have ins`, `doesnt have insurance` and
  // `pt doesnt have insurance` in one rule.
  { keyword: 'doesnt have ins', destination: 'Not Accepted' }, // 4
  { keyword: "doesn't have ins", destination: 'Not Accepted' },

  // Confirmed by the office: the campium/campflow wording is Missing Info.
  { keyword: 'not on campium', destination: 'Missing Info' },
  { keyword: 'not on campflow', destination: 'Missing Info' },

  // Its own sheet, append-only. Listed last so a cell naming both an EMR and a
  // real queue keyword still routes to the queue.
  { keyword: 'united refuah', destination: 'United Refuah' },
] as const;



export type StatusOutcome =
  /** Column B is empty: the patient has not been processed yet. */
  | { kind: 'blank' }
  /**
   * `needs ohi` / `needs lasante` — still waiting to be entered into an EMR.
   * Distinct from `terminal` even though neither produces a queue row, because
   * conflating the two is what marked 1,268 patients as finished.
   */
  | { kind: 'pending'; matched: string[] }
  /** `skip` — a successful follow-up visit that cannot be billed. No action. */
  | { kind: 'ignored'; matched: string[] }
  /** Entered into an EMR successfully. */
  | { kind: 'terminal'; matched: string[] }
  | { kind: 'queued'; destination: QueueSheetName; keyword: string; matched: string[] }
  | { kind: 'unrecognized' };

/**
 * Normalizes a cell value for containment testing.
 *
 * Lowercases, converts non-breaking spaces and unicode dashes to their ASCII
 * equivalents, collapses runs of whitespace, and trims. Collapsing internal
 * whitespace matters: `missing  info` typed with a double space is still Missing
 * Info, and a plain `includes` would miss it.
 */
export function normalizeStatus(raw: unknown): string {
  if (raw === null || raw === undefined) return '';
  return String(raw)
    .replace(/[\u00A0\u2007\u202F]/g, ' ') // non-breaking / figure / narrow spaces
    .replace(/[\u2010-\u2015\u2212]/g, '-') // unicode dashes and the minus sign
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/** Every keyword contained in the value, terminal and queued alike, in table order. */
export function allMatches(normalized: string): string[] {
  const found: string[] = [];
  for (const keyword of TERMINAL_KEYWORDS) {
    if (normalized.includes(keyword)) found.push(keyword);
  }
  for (const { keyword } of QUEUE_KEYWORDS) {
    if (normalized.includes(keyword)) found.push(keyword);
  }
  return found;
}

export function classifyStatus(raw: unknown): StatusOutcome {
  const normalized = normalizeStatus(raw);
  if (normalized === '') return { kind: 'blank' };

  const matched = allMatches(normalized);

  // 1. "needs ohi" before "ohi". A prefix that negates has to beat the terminal
  //    check, or 1,268 outstanding patients read as finished.
  if (PENDING_PATTERNS.some((pattern) => pattern.test(normalized))) {
    return { kind: 'pending', matched };
  }

  // 2. Deliberate no-ops.
  if (IGNORED_KEYWORDS.some((k) => normalized.includes(k))) {
    return { kind: 'ignored', matched };
  }

  // 3. Terminal wins over any queue keyword, whatever suffix follows.
  if (TERMINAL_KEYWORDS.some((k) => normalized.includes(k))) {
    return { kind: 'terminal', matched };
  }

  // 4. First queued keyword in table order.
  for (const { keyword, destination } of QUEUE_KEYWORDS) {
    if (normalized.includes(keyword)) {
      return { kind: 'queued', destination, keyword, matched };
    }
  }

  // 5. Unrecognized is never guessed at.
  return { kind: 'unrecognized' };
}

/**
 * True when a value matched more than one keyword in a way a human should look
 * at.
 *
 * Deliberately not "matched.length > 1": `ineligible/inactive` matches two
 * keywords that route to the same sheet, which is a normal thing for staff to
 * type and not worth alerting on. A multi-match is only interesting when the
 * matches disagree about the outcome — a terminal keyword sitting alongside a
 * queued one, or two queued keywords with different destinations.
 */
export function isAmbiguous(outcome: StatusOutcome): boolean {
  if (outcome.kind !== 'queued' && outcome.kind !== 'terminal') return false;
  const matched = outcome.matched;
  if (matched.length < 2) return false;

  const hasTerminal = matched.some((k) => TERMINAL_KEYWORDS.includes(k));

  // A cell matching `no insurance on file` also matches `no insurance`, and the
  // two route differently ON PURPOSE — the specific rule is deliberately listed
  // above the general one. Flagging that as a data problem would emit the same
  // warnings forever and train everyone to ignore the channel, burying the real
  // multi-matches. So: if every other matched keyword is contained within the
  // winning one, the overlap is by construction, not a human error.
  const winner = outcome.kind === 'queued' ? outcome.keyword : undefined;
  const others = matched.filter((k) => k !== winner);
  if (winner && others.every((k) => winner.includes(k))) return false;

  const destinations = new Set(
    matched
      .map((k) => QUEUE_KEYWORDS.find((q) => q.keyword === k)?.destination)
      .filter((d): d is QueueSheetName => d !== undefined),
  );

  if (hasTerminal && destinations.size > 0) return true;
  return destinations.size > 1;
}
