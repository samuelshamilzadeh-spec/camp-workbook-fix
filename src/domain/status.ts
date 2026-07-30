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
  'united refuah',
] as const;

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

  { keyword: 'verify insurance', destination: 'Verify Insurance' }, // 221
] as const;

/**
 * Variants the inspection found that are NOT yet routed, because routing them
 * would be a guess and the brief is explicit: never guess a destination.
 *
 * Each of these is a real patient sitting in no queue. They are listed here so
 * the question is answerable by the office rather than lost — see the open
 * questions in the README.
 *
 *   no insurance on file (5), need insurance (3), doesnt have insurance (2),
 *   no insurance (2), need ins (1), doesnt have ins (1),
 *   need insurance verification (1), incorrect insurance (1), invalid ins (1),
 *   wrote paper has no ins (1), pt doesnt have insurance (1)
 *     -> Verify Insurance, or Missing Info? Both are defensible.
 *
 *   skip (7), not on campium (3), not on campflow (1),
 *   need to confirm dob (1), same w/ line NN (1)
 *     -> unknown. `skip` may well be deliberate exclusion.
 */
export const UNROUTED_VARIANTS: readonly string[] = [
  'no insurance on file',
  'need insurance',
  'doesnt have insurance',
  'no insurance',
  'need ins',
  'doesnt have ins',
  'need insurance verification',
  'incorrect insurance',
  'invalid ins',
  'wrote paper has no ins',
  'pt doesnt have insurance',
  'skip',
  'not on campium',
  'not on campflow',
  'need to confirm dob',
] as const;

export type StatusOutcome =
  | { kind: 'blank' }
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

  // 1. Terminal wins outright, whatever suffix follows and whatever else the
  //    cell contains.
  if (TERMINAL_KEYWORDS.some((k) => normalized.includes(k))) {
    return { kind: 'terminal', matched };
  }

  // 2. First queued keyword in table order.
  for (const { keyword, destination } of QUEUE_KEYWORDS) {
    if (normalized.includes(keyword)) {
      return { kind: 'queued', destination, keyword, matched };
    }
  }

  // 3. Unrecognized is never guessed at.
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
  const destinations = new Set(
    matched
      .map((k) => QUEUE_KEYWORDS.find((q) => q.keyword === k)?.destination)
      .filter((d): d is QueueSheetName => d !== undefined),
  );

  if (hasTerminal && destinations.size > 0) return true;
  return destinations.size > 1;
}
