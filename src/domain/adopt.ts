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
}

export interface AdoptionProblem {
  queueSheet: string;
  queueRow: number;
  reason:
    | 'no-pointer'
    | 'unknown-source-sheet'
    | 'source-row-out-of-range'
    | 'identity-mismatch'
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
    const candidate = rows?.get(sourceRowNumber);
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
      if (sameValue(queueValue, dailyValue)) confirmedBy.push(field);
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

    const key = `${sourceSheet}!${sourceRowNumber}`;
    if (claimed.has(key)) {
      problems.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        reason: 'source-already-adopted',
        sourceSheet,
        sourceRow: sourceRowNumber,
      });
      continue;
    }
    claimed.add(key);

    matches.push({
      queueSheet: queueRow.sheet,
      queueRow: queueRow.row,
      sourceSheet,
      sourceRow: sourceRowNumber,
      // Reuse the source row's existing ID when it has one, so a row already
      // stamped is linked rather than given a second identity.
      syncId: candidate.syncId ?? input.newSyncId(),
      confirmedBy,
    });
  }

  return { matches, problems };
}
