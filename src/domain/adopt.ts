import { LAYOUT, type QueueColumn, type WorkbookLayout } from '../config';
import { isBlank } from './cells';
import type { DailyRow, ParsedDailySheet } from './dailySheets';
import type { QueueRow } from './queueSheets';
import { sameValue } from './reconcile';

/**
 * Adoption: giving the queue rows that already exist the same SyncID as their
 * source row on the daily sheet.
 *
 * The office needs the rows sitting on the live queue tabs today to keep
 * working — staff are editing them right now — and needs an edit made there to
 * reach the daily sheet. That only works if both ends share an ID.
 *
 * The existing rows predate SyncID, so the only link back is the old pointer:
 * `Date of Visit` plus `Source Row`. The build brief is blunt about that
 * pointer — it "breaks silently the moment anyone inserts, deletes, or sorts a
 * row on a daily sheet", and it is the reason SyncID exists at all.
 *
 * So the pointer is used to FIND the candidate and never trusted on its own.
 * The patient's identifying fields are then compared, and a row is adopted only
 * when they agree. A pointer that has drifted produces a mismatch, which is
 * reported for a human rather than silently binding one patient's queue row to
 * another patient's source row — the exact failure the brief calls the primary
 * accuracy risk in the project.
 *
 * This runs once. Afterwards every row has an ID and none of this matters.
 */

/**
 * Identity of a patient within one daily sheet, for finding a row whose pointer
 * has drifted.
 */
export function identityKey(fields: Partial<Record<QueueColumn, unknown>>): string | undefined {
  const parts = IDENTITY_FIELDS.map((field) => {
    const value = fields[field];
    if (isBlank(value)) return '';
    return field === 'Date of Birth'
      ? (toDateKey(value) ?? String(value).trim().toLowerCase())
      : String(value).replace(/\s+/g, ' ').trim().toLowerCase();
  });
  // All three are required: two of three is how siblings get merged.
  return parts.every((part) => part !== '') ? parts.join('|') : undefined;
}

/** Fields compared to confirm a pointer landed on the right patient. */
const IDENTITY_FIELDS: readonly QueueColumn[] = [
  'Last Name',
  'First Name',
  'Date of Birth',
] as const;

export interface AdoptionMatch {
  queueSheet: string;
  queueRow: number;
  sourceSheet: string;
  sourceRow: number;
  syncId: string;
  /** Identity fields that were present on both sides and agreed. */
  confirmedBy: QueueColumn[];
  /**
   * `pointer` — the stored Source Row was correct.
   * `identity` — the pointer had drifted and the patient was found by name and
   * date of birth instead.
   */
  via: 'pointer' | 'identity';
  /** How far the stored pointer was wrong by. Only set when via is identity. */
  drift?: number;
}

export interface AdoptionProblem {
  queueSheet: string;
  queueRow: number;
  reason:
    | 'no-pointer'
    | 'unknown-source-sheet'
    | 'source-row-out-of-range'
    | 'identity-mismatch'
    | 'identity-ambiguous'
    | 'too-little-identity'
    | 'source-already-adopted';
  /** Set when a candidate was found but rejected. */
  sourceSheet?: string;
  sourceRow?: number;
  /** Identity fields that disagreed. Names only, never values. */
  mismatchedFields?: QueueColumn[];
}

export interface AdoptionResult {
  matches: AdoptionMatch[];
  problems: AdoptionProblem[];
}

/**
 * Excel stores a date as days since 1899-12-30 in the 1900 date system. The
 * queue tabs hold `Date of Visit` as a serial (46225) while the daily tabs are
 * named `July 5, 2026`, so one has to be converted to reach the other.
 */
export function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2958465) return null;
  const ms = Math.round(serial) * 86400000;
  const date = new Date(Date.UTC(1899, 11, 30) + ms);
  return Number.isNaN(date.getTime()) ? null : date;
}

function toUtcDay(date: Date): number {
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

/**
 * Canonical `YYYY-MM-DD` for a value that might be a date, or undefined if it
 * is not one.
 *
 * The same birthday is stored two different ways in this workbook: the queue
 * tabs hold an Excel serial (`40147`) and the daily sheets hold text
 * (`11/30/2009`). Both mean 30 November 2009. Comparing their string forms
 * makes every one of them look like a different patient — which is exactly
 * what it did, producing 13 false "identity mismatch" rejections whose only
 * disagreement was the date of birth.
 */
export function toDateKey(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;

  if (typeof value === 'number' || /^\d+(\.\d+)?$/.test(String(value).trim())) {
    const date = excelSerialToDate(Number(value));
    return date ? date.toISOString().slice(0, 10) : undefined;
  }

  const text = String(value).trim();
  if (!text) return undefined;

  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (iso) return calendarKey(Number(iso[1]), Number(iso[2]), Number(iso[3]));

  // US order, which is what the daily sheets use.
  const us = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/.exec(text);
  if (us) {
    const year = Number(us[3]);
    const full = year < 100 ? (year > 30 ? 1900 + year : 2000 + year) : year;
    return calendarKey(full, Number(us[1]), Number(us[2]));
  }

  return undefined;
}

/**
 * A date key, or nothing if those numbers are not a day on the calendar.
 *
 * The validation is the point. `14/24/2014` is a typo somebody made — there is
 * no fourteenth month — and JavaScript's Date rolls it forward to 24 February
 * 2015 without a word. Left unchecked, converting the daily sheets' text dates
 * to real dates turned that typo into a plausible, wrong, unremarkable date of
 * birth: the kind of error nobody spots again, on the record used for billing.
 *
 * Live example: `Not Accepted ` row 377, whose source cell on `July 16, 2026`
 * reads `14/24/2014`.
 *
 * Returning undefined leaves the original text exactly as staff typed it, so
 * the mistake stays visible and fixable.
 */
function calendarKey(year: number, month: number, day: number): string | undefined {
  if (!Number.isInteger(year) || year < 1900 || year > 2200) return undefined;
  if (!Number.isInteger(month) || month < 1 || month > 12) return undefined;
  if (!Number.isInteger(day) || day < 1) return undefined;

  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  if (day > daysInMonth) return undefined;

  return `${year}-${pad(String(month))}-${pad(String(day))}`;
}

function pad(part: string): string {
  return part.padStart(2, '0');
}

/**
 * Compares two values that may both be dates in different representations,
 * falling back to the ordinary comparison when neither is.
 */
export function sameDateOrValue(
  a: unknown,
  b: unknown,
  fallback: (x: unknown, y: unknown) => boolean,
): boolean {
  const left = toDateKey(a);
  const right = toDateKey(b);
  if (left !== undefined && right !== undefined) return left === right;
  return fallback(a, b);
}

/**
 * Resolves a queue row's `Date of Visit` to a daily sheet name. Accepts either
 * the serial the live tabs use or a plain sheet name, since rows this code
 * writes carry the name.
 */
export function resolveSourceSheet(
  dateOfVisit: unknown,
  dailySheets: readonly string[],
  layout: WorkbookLayout = LAYOUT,
): string | undefined {
  if (isBlank(dateOfVisit)) return undefined;

  const text = String(dateOfVisit).trim();
  if (dailySheets.includes(text)) return text;

  const serial = Number(text);
  const date = Number.isFinite(serial) ? excelSerialToDate(serial) : null;
  const target = date ?? layout.parseDailySheetDate(text);
  if (!target) return undefined;

  const wanted = toUtcDay(target);
  return dailySheets.find((sheet) => {
    const parsed = layout.parseDailySheetDate(sheet);
    return parsed !== null && toUtcDay(parsed) === wanted;
  });
}

export interface AdoptionInput {
  queueRows: readonly QueueRow[];
  daily: readonly ParsedDailySheet[];
  /** All daily rows by sheet, including those with no status keyword. */
  dailyRowsBySheet: ReadonlyMap<string, ReadonlyMap<number, DailyRow>>;
  newSyncId: () => string;
  layout?: WorkbookLayout;
}

export function planAdoption(input: AdoptionInput): AdoptionResult {
  const layout = input.layout ?? LAYOUT;
  const matches: AdoptionMatch[] = [];
  const problems: AdoptionProblem[] = [];
  const sheetNames = input.daily.map((sheet) => sheet.sheet);

  // One source row must never be adopted by two queue rows: that would point
  // two patients' edits at the same cells.
  const claimed = new Set<string>();

  for (const queueRow of input.queueRows) {
    if (queueRow.syncId) continue; // already linked

    const sourceSheet = resolveSourceSheet(queueRow.values['Date of Visit'], sheetNames, layout);
    const sourceRowNumber = Number(String(queueRow.values['Source Row'] ?? '').trim());

    if (!sourceSheet || !Number.isInteger(sourceRowNumber) || sourceRowNumber < 1) {
      problems.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        reason: !sourceSheet ? 'unknown-source-sheet' : 'no-pointer',
        sourceSheet,
      });
      continue;
    }

    const rows = input.dailyRowsBySheet.get(sourceSheet);
    let candidate = rows?.get(sourceRowNumber);
    let via: 'pointer' | 'identity' = 'pointer';
    let drift: number | undefined;

    // Does the pointer actually land on this patient? If not, find them by who
    // they are. Observed live: someone inserted one row on one daily sheet and
    // deleted one on another, and 43 pointers drifted by exactly ±1 within
    // ninety minutes. The pointer is a hint; identity is the truth.
    const wanted = identityKey(queueRow.values);
    const landedOnSomeoneElse =
      candidate !== undefined && wanted !== undefined && identityKey(candidate.fields) !== wanted;

    if ((candidate === undefined || landedOnSomeoneElse) && wanted !== undefined && rows) {
      const found: number[] = [];
      for (const [row, daily] of rows) {
        if (identityKey(daily.fields) === wanted) found.push(row);
      }

      if (found.length === 1) {
        candidate = rows.get(found[0]!);
        via = 'identity';
        drift = found[0]! - sourceRowNumber;
      } else if (found.length > 1) {
        // The same patient twice on one day. Picking either would be a guess.
        problems.push({
          queueSheet: queueRow.sheet,
          queueRow: queueRow.row,
          reason: 'identity-ambiguous',
          sourceSheet,
          sourceRow: sourceRowNumber,
        });
        continue;
      }
    }

    if (!candidate) {
      problems.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        reason: 'source-row-out-of-range',
        sourceSheet,
        sourceRow: sourceRowNumber,
      });
      continue;
    }

    // Verify, do not trust. A pointer that drifted still resolves to a row —
    // just the wrong patient's.
    const confirmedBy: QueueColumn[] = [];
    const mismatchedFields: QueueColumn[] = [];
    for (const field of IDENTITY_FIELDS) {
      const queueValue = queueRow.values[field];
      const dailyValue = candidate.fields[field];
      if (isBlank(queueValue) && isBlank(dailyValue)) continue;
      // Date-aware: the same birthday is an Excel serial on the queue tabs and
      // text on the daily sheets.
      if (sameDateOrValue(queueValue, dailyValue, sameValue)) confirmedBy.push(field);
      else mismatchedFields.push(field);
    }

    if (mismatchedFields.length > 0) {
      problems.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        reason: 'identity-mismatch',
        sourceSheet,
        sourceRow: sourceRowNumber,
        mismatchedFields,
      });
      continue;
    }

    // Agreement on nothing is not agreement. Two rows that are both blank in
    // every identity field tell us only that the pointer resolved.
    if (confirmedBy.length === 0) {
      problems.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        reason: 'too-little-identity',
        sourceSheet,
        sourceRow: sourceRowNumber,
      });
      continue;
    }

    const resolvedRow = candidate.row;
    const key = `${sourceSheet}!${resolvedRow}`;
    if (claimed.has(key)) {
      problems.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        reason: 'source-already-adopted',
        sourceSheet,
        sourceRow: resolvedRow,
      });
      continue;
    }
    claimed.add(key);

    matches.push({
      queueSheet: queueRow.sheet,
      queueRow: queueRow.row,
      sourceSheet,
      sourceRow: resolvedRow,
      via,
      ...(drift !== undefined && drift !== 0 ? { drift } : {}),
      // Reuse the source row's existing ID when it has one, so a row already
      // stamped is linked rather than given a second identity.
      syncId: candidate.syncId ?? input.newSyncId(),
      confirmedBy,
    });
  }

  return { matches, problems };
}
