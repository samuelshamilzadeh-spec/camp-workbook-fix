import { isBlank } from './cells';
import type { QueueColumn } from '../config';

/**
 * Patient identity across visits.
 *
 * A patient seen twice gets two rows — two visits, two insurance entries, and
 * the office wants both on the daily sheet AND on the queue. But a fix typed
 * once should reach all of them: 373 of the 1,061 queued visits belong to a
 * patient who appears more than once, so without this, a fifth of the queue is
 * the same work re-typed.
 *
 * Two identifiers, doing different jobs:
 *
 *   SyncID    one visit row. Links a queue row to the exact daily row it came
 *             from, so an edit lands in the right cells.
 *   PatientID one person. Shared by every visit of that person, so a fix can
 *             fan out to their other visits.
 *
 * The key is Last Name + First Name + Date of Birth + **Camp**. Camp was the
 * office's addition and it is free: on the live data it produces exactly the
 * same 836 people as name+DOB alone, splitting nobody, while ruling out the
 * case where two children genuinely share a name and a birthday. Fourteen
 * name-pairs in this workbook already cover more than one date of birth, so
 * name alone was never safe.
 *
 * Which way this fails matters. A typo in a name or camp splits one patient
 * into two — the office does the work twice, exactly as today. Merging two
 * different children would put one child's insurance on another's record. The
 * key is built to fail toward the first.
 */

/** Fields that make up a patient's identity. */
export const IDENTITY_COLUMNS: readonly QueueColumn[] = [
  'Last Name',
  'First Name',
  'Date of Birth',
] as const;

function normalizePart(value: unknown): string {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Returns undefined when identity is too thin to be trusted. A row missing a
 * surname or a date of birth is not matched to anything — better a duplicate
 * than a wrong merge. Only 1 of 1,061 queued rows is in that state.
 */
export function patientKey(
  fields: Partial<Record<QueueColumn, unknown>>,
  camp: string | undefined,
): string | undefined {
  const last = normalizePart(fields['Last Name']);
  const first = normalizePart(fields['First Name']);
  const dob = normalizePart(fields['Date of Birth']);
  const campKey = normalizePart(camp);

  if (!last || !first || !dob) return undefined;
  return [last, first, dob, campKey].join('|');
}

export function hasIdentity(
  fields: Partial<Record<QueueColumn, unknown>>,
): boolean {
  return IDENTITY_COLUMNS.every((column) => !isBlank(fields[column]));
}

/**
 * Decides what a field should hold across every visit of one patient.
 *
 * `visits` gives, per visit, the value currently on the queue sheet and the
 * value on the daily sheet it came from.
 *
 * A queue value that differs from its own source is a staff edit — somebody
 * typed it just now — and that is what propagates. Everything else is left
 * alone.
 *
 * Two staff edits disagreeing is not resolvable here, so nothing is written and
 * the conflict is reported. Silently picking one would discard the other, and
 * on this data that is not rare: 72 of the 148 repeat patients already disagree
 * on some filled field, most often a phone number.
 */
export interface VisitValue {
  syncId: string;
  queueValue: unknown;
  sourceValue: unknown;
}

export type FieldResolution =
  | { kind: 'no-change' }
  | { kind: 'propagate'; value: unknown; from: string }
  | { kind: 'conflict'; editedBy: string[] };

export function resolveFieldAcrossVisits(
  visits: readonly VisitValue[],
  sameValue: (a: unknown, b: unknown) => boolean,
): FieldResolution {
  const edits = visits.filter((visit) => !sameValue(visit.queueValue, visit.sourceValue));
  if (edits.length === 0) return { kind: 'no-change' };

  const distinct: unknown[] = [];
  for (const edit of edits) {
    if (!distinct.some((value) => sameValue(value, edit.queueValue))) {
      distinct.push(edit.queueValue);
    }
  }

  if (distinct.length > 1) {
    return { kind: 'conflict', editedBy: edits.map((edit) => edit.syncId) };
  }

  return { kind: 'propagate', value: edits[0]!.queueValue, from: edits[0]!.syncId };
}
