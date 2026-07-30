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

/** Terminal success states. Presence anywhere in the value suppresses queuing. */
export const TERMINAL_KEYWORDS: readonly string[] = ['ohi', 'lasante'] as const;

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
  { keyword: 'not accepted', destination: 'Not Accepted' },
  { keyword: 'missing info', destination: 'Missing Info' },
  { keyword: 'ineligible', destination: 'Ineligible & Inactive' },
  { keyword: 'inactive', destination: 'Ineligible & Inactive' },
  { keyword: 'verify insurance', destination: 'Verify Insurance' },
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
