import { STYLE, type QueueColumn } from '../config';
import { isBlank } from './cells';
import { toDateKey } from './adopt';

/**
 * Does a queue cell and a daily cell mean different things?
 *
 * The reconciler's rule is that a queue value differing from its source row is a
 * staff edit, and it gets written back to the daily sheet. That rule is only as
 * good as the comparison behind it, and a plain string comparison is not good
 * enough — measured against the live workbook it called **1,179 cells staff
 * edits, and not one of them was**:
 *
 *     770  case only               State "LAKEWOOD"      vs "Lakewood"
 *     369  punctuation only        Phone "(248) 331-3973" vs "248-331-3973"
 *      40  the rest, mostly        Zip   "8701"          vs "08701"
 *
 * The old mirror sheets upper-cased towns and bracketed area codes; staff type
 * neither. Nobody edited anything. Applying those write-backs would have pushed
 * the mirror's formatting over 557 patients' records on the daily sheets — and
 * the zip codes are worse than cosmetic, because the queue cell is numeric and
 * has already lost its leading zero. `08701` is Lakewood, New Jersey. `8701` is
 * not a zip code at all, and the daily sheets are the copy that gets billed
 * from.
 *
 * So comparison is per-field and asks about meaning:
 *
 *   - phone numbers by their digits, with a US country code ignored
 *   - zip codes by their first five digits, leading zeros restored
 *   - dates by the day they denote, in any representation
 *   - everything else case-insensitively, with whitespace collapsed
 *
 * A genuine edit still shows up: a different phone number has different digits.
 * What no longer shows up is the same value written two ways.
 */

/** Digits only, with a leading US country code dropped. */
function phoneKey(value: string): string {
  const digits = value.replace(/\D/g, '');
  return digits.length === 11 && digits.startsWith('1') ? digits.slice(1) : digits;
}

/**
 * First five digits, zero-padded.
 *
 * Excel stores `08701` in a numeric cell as the number 8701, so the queue side
 * arrives already stripped. Padding both sides means the two agree instead of
 * the shorter one being written over the correct one.
 */
function zipKey(value: string): string {
  const digits = value.replace(/\D/g, '');
  if (digits.length === 0) return '';
  return digits.slice(0, 5).padStart(5, '0');
}

function textKey(value: string): string {
  return value.replace(/\s+/g, ' ').trim().toLowerCase();
}

/** The comparable form of one field's value. Exported for testing and reporting. */
export function comparisonKey(field: QueueColumn, value: unknown): string {
  if (isBlank(value)) return '';
  const raw = String(value).trim();

  if (STYLE.dateColumns.includes(field)) {
    // A date is the day it denotes, whether it arrived as a serial or as text.
    return toDateKey(raw) ?? textKey(raw);
  }

  switch (field) {
    case 'Phone Number':
      return phoneKey(raw);
    case 'Zip Code':
      return zipKey(raw);
    default:
      return textKey(raw);
  }
}

/**
 * True when the two sides carry the same information for this field.
 *
 * Blank on both sides counts as agreement. Blank on one side does not: a value
 * appearing where there was none is exactly the edit this system exists to
 * carry, and a value disappearing is worth surfacing rather than silently
 * treating as equal.
 */
export function sameMeaning(field: QueueColumn, a: unknown, b: unknown): boolean {
  if (isBlank(a) && isBlank(b)) return true;
  if (isBlank(a) || isBlank(b)) return false;
  return comparisonKey(field, a) === comparisonKey(field, b);
}
