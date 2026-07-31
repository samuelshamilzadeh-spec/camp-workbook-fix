import { toDateKey } from './adopt';

/**
 * Excel dates, in both directions.
 *
 * `adopt.ts` already converts a serial to a Date, because it had to compare a
 * queue tab's serial against a daily sheet's text. This is the other direction,
 * needed since the office asked for real dates on the queue sheets rather than
 * the text this project was writing.
 *
 * That matters for more than looks. A text date cannot be sorted, filtered by
 * range, or formatted, and `07/05/2026` typed into a text column is a different
 * string from `7/5/2026` while being the same day.
 */

/** Days between the Excel epoch (1899-12-30, 1900 system) and the Unix epoch. */
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const DAY_MS = 86400000;

/**
 * `2009-11-30` or a Date -> 40147. Returns undefined for anything that is not a
 * date, so a caller can leave the original value alone rather than writing a
 * number over something it did not understand.
 */
export function toExcelSerial(value: unknown): number | undefined {
  const key = toDateKey(value);
  if (!key) return undefined;

  const [year, month, day] = key.split('-').map(Number);
  if (!year || !month || !day) return undefined;

  const serial = (Date.UTC(year, month - 1, day) - EXCEL_EPOCH_UTC) / DAY_MS;
  return Number.isInteger(serial) && serial >= 1 ? serial : undefined;
}

/**
 * The value to write into a date cell: a serial when the input is a recognizable
 * date, otherwise the input untouched.
 *
 * Leaving an unrecognized value alone is the point. Dates of birth on the daily
 * sheets are free text and some of them are not dates at all; overwriting one
 * with a guess would destroy the only record of what staff actually typed.
 */
export function toDateCellValue(value: unknown): unknown {
  return toExcelSerial(value) ?? value;
}

/**
 * True when two values denote the same day, in any mix of representations —
 * serial, `11/30/2009`, `2009-11-30`.
 *
 * Needed because the queue sheets now hold serials while the daily sheets hold
 * text. Compared as strings, every single date would look like a staff edit, and
 * the reconciler would write back the same value on every cycle forever.
 */
export function sameDay(a: unknown, b: unknown): boolean {
  const left = toDateKey(a);
  const right = toDateKey(b);
  return left !== undefined && right !== undefined && left === right;
}
