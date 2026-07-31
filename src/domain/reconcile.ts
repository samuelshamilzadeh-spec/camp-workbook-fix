import {
  APPEND_ONLY_DESTINATIONS,
  LAYOUT,
  QUEUE_ONLY_COLUMNS,
  REQUIRED_FIELDS,
  STYLE,
  type QueueColumn,
  type QueueSheetName,
  type WorkbookLayout,
} from '../config';
import { isBlank } from './cells';
import { sameDay } from './dates';
import { campKey, type DailyRow, type ParsedDailySheet } from './dailySheets';
import type { ParsedQueueSheet, QueueRow } from './queueSheets';

/**
 * State-based reconciliation.
 *
 * Every cycle reads what column B says right now and decides where each row
 * belongs. Nothing here tries to detect "it changed from X to Y". That is what
 * makes the job idempotent, makes catch-up after downtime identical to the
 * normal path, and makes a cleared cell need no special case.
 *
 * Running this twice against the same workbook state must produce an identical
 * plan. There is a test asserting exactly that.
 */

export interface StampIdIntent {
  kind: 'stamp-id';
  sheet: string;
  row: number;
  syncId: string;
}

export interface AppendQueueRowIntent {
  kind: 'append-queue-row';
  destination: QueueSheetName;
  syncId: string;
  sourceSheet: string;
  sourceRow: number;
  camp: string | undefined;
  values: Partial<Record<QueueColumn, unknown>>;
  /** Required fields that are blank and therefore get dark red shading. */
  blankRequired: QueueColumn[];
}

export interface WriteBackIntent {
  kind: 'write-back';
  syncId: string;
  sourceSheet: string;
  sourceRow: number;
  queueSheet: QueueSheetName;
  queueRow: number;
  field: QueueColumn;
  value: unknown;
}

export interface RemoveQueueRowIntent {
  kind: 'remove-queue-row';
  syncId: string | undefined;
  queueSheet: QueueSheetName;
  queueRow: number;
  /** Set when the source row is known and its column B must also be cleared. */
  sourceSheet: string | undefined;
  sourceRow: number | undefined;
  reason:
    | 'resolved-on-queue'
    | 'cleared-on-queue'
    | 'no-longer-queued-at-source'
    | 'source-row-missing';
}

export type Intent =
  | StampIdIntent
  | AppendQueueRowIntent
  | WriteBackIntent
  | RemoveQueueRowIntent;

export interface AmbiguityReport {
  sheet: string;
  row: number;
  syncId: string | undefined;
  matched: string[];
  destination: QueueSheetName | 'terminal';
}

export interface OrphanReport {
  queueSheet: QueueSheetName;
  queueRow: number;
  syncId: string | undefined;
  reason: 'unknown-sync-id' | 'missing-sync-id' | 'duplicate-sync-id';
}

/** Somebody typed something in `Resolved` that is not one of the accepted words. */
export interface UnrecognizedResolvedReport {
  queueSheet: QueueSheetName;
  queueRow: number;
  syncId: string;
  /** What they typed. A marker word, never a patient field. */
  raw: string;
}

export interface ReconcilePlan {
  intents: Intent[];
  ambiguous: AmbiguityReport[];
  orphans: OrphanReport[];
  unrecognizedResolved: UnrecognizedResolvedReport[];
  unrecognizedCounts: Record<string, number>;
  counts: Record<string, number>;
}

export interface ReconcileInput {
  daily: ParsedDailySheet[];
  queues: ParsedQueueSheet[];
  /** Injected so tests are deterministic and so IDs are minted in one place. */
  newSyncId: () => string;
  layout?: WorkbookLayout;
}

export function reconcile(input: ReconcileInput): ReconcilePlan {
  const layout = input.layout ?? LAYOUT;
  const intents: Intent[] = [];
  const ambiguous: AmbiguityReport[] = [];
  const orphans: OrphanReport[] = [];
  const unrecognizedResolved: UnrecognizedResolvedReport[] = [];
  const unrecognizedCounts: Record<string, number> = {};

  // --- Index the source side -------------------------------------------------
  //
  // Only rows carrying a recognized keyword are here. That is the whole point of
  // stamping lazily: a patient who never receives a keyword in column B never
  // needs an ID, group headers and instruction blocks are excluded for free, and
  // `run_daily.py` can keep writing rows with no ID.
  const dailyById = new Map<string, DailyRow>();
  const pendingStamps: DailyRow[] = [];

  for (const sheet of input.daily) {
    if (sheet.unrecognizedRows.length > 0) {
      unrecognizedCounts[sheet.sheet] = sheet.unrecognizedRows.length;
    }

    for (const row of sheet.rows) {
      if (row.ambiguous) {
        ambiguous.push({
          sheet: row.sheet,
          row: row.row,
          syncId: row.syncId,
          matched:
            row.outcome.kind === 'queued' || row.outcome.kind === 'terminal'
              ? row.outcome.matched
              : [],
          destination:
            row.outcome.kind === 'queued' ? row.outcome.destination : 'terminal',
        });
      }

      if (row.syncId) {
        dailyById.set(row.syncId, row);
      } else if (row.outcome.kind === 'queued') {
        // Stamp at the moment a row first enters a queue, in the same operation
        // that reads it. Never a bulk walk looking for blanks.
        pendingStamps.push(row);
      }
    }
  }

  // --- Index the queue side --------------------------------------------------
  //
  // EVERY row an ID appears on, not just the last one seen.
  //
  // A patient legitimately sits on two queue tabs for a while: their column B
  // changed, so they belong on the new queue and their old row is waiting for
  // Phase 4 to remove it. Keyed by ID alone, the second tab parsed overwrote the
  // first, and which one survived depended on the order of QUEUE_SHEET_NAMES.
  //
  // That is not a cosmetic problem. If the surviving row was the stale one, the
  // reconciler concluded "wrong queue" and planned an append to the correct
  // queue — where the row already was. Every run planned it again, so every run
  // added another copy of that patient. Caught on the live workbook by
  // re-running the appender against a tab it had just written: 243 rows written,
  // and it still wanted to append one of them.
  const queueById = new Map<string, QueueRow[]>();
  for (const queue of input.queues) {
    for (const row of queue.rows) {
      if (!row.syncId) {
        orphans.push({
          queueSheet: queue.sheet,
          queueRow: row.row,
          syncId: undefined,
          reason: 'missing-sync-id',
        });
        continue;
      }
      const rows = queueById.get(row.syncId) ?? [];
      rows.push(row);
      queueById.set(row.syncId, rows);
    }
  }

  // --- 1. New rows entering a queue -----------------------------------------
  for (const row of pendingStamps) {
    if (row.outcome.kind !== 'queued') continue;
    const syncId = input.newSyncId();
    intents.push({ kind: 'stamp-id', sheet: row.sheet, row: row.row, syncId });
    intents.push(buildAppend(row, syncId, row.outcome.destination));
  }

  // --- 2. Already-stamped source rows ---------------------------------------
  for (const [syncId, dailyRow] of dailyById) {
    const queueRows = queueById.get(syncId) ?? [];

    const removeStale = (row: QueueRow): void => {
      intents.push({
        kind: 'remove-queue-row',
        syncId,
        queueSheet: row.sheet,
        queueRow: row.row,
        sourceSheet: undefined,
        sourceRow: undefined,
        reason: 'no-longer-queued-at-source',
      });
    };

    if (dailyRow.outcome.kind === 'queued') {
      const destination = dailyRow.outcome.destination;
      const onDestination = queueRows.filter((row) => row.sheet === destination);

      // Any row on a queue this patient no longer belongs to goes, whether or
      // not they also have a row on the right one.
      for (const row of queueRows) {
        if (row.sheet !== destination) removeStale(row);
      }

      if (onDestination.length === 0) {
        // Not on the queue they belong to: append. Covers both the never-queued
        // case and the moved-queue case, and the stale row was just removed.
        intents.push(buildAppend(dailyRow, syncId, destination));
        continue;
      }

      if (onDestination.length > 1) {
        // Two rows on one queue for one visit. Nothing here picks which to keep
        // — the first is reconciled and the rest are reported, because deleting
        // a row a human may have been editing is not a guess worth making.
        for (const row of onDestination.slice(1)) {
          orphans.push({
            queueSheet: row.sheet,
            queueRow: row.row,
            syncId,
            reason: 'duplicate-sync-id',
          });
        }
      }

      // An append-only destination is a record, not a work queue: the office
      // confirmed a United Refuah row is copied across and then never changes,
      // and nothing is ever sent back to the daily sheet.
      if (APPEND_ONLY_DESTINATIONS.includes(destination)) continue;

      // Same queue: reconcile field values. A staff edit on the queue sheet is
      // authoritative and gets copied back to the daily sheet. The Resolved
      // marker and the Source Row link stay put.
      const queueRow = onDestination[0]!;
      for (const [column, queueValue] of Object.entries(queueRow.values)) {
        const field = column as QueueColumn;
        if (QUEUE_ONLY_COLUMNS.includes(field)) continue;
        if (!layout.daily.fieldColumns[field]) continue;
        if (sameFieldValue(field, queueValue, dailyRow.fields[field])) continue;

        intents.push({
          kind: 'write-back',
          syncId,
          sourceSheet: dailyRow.sheet,
          sourceRow: dailyRow.row,
          queueSheet: queueRow.sheet,
          queueRow: queueRow.row,
          field,
          value: queueValue,
        });
      }

      // Somebody marked this row done. The write-backs above carry whatever they
      // fixed to the daily sheet; this takes the row off the queue and clears
      // column B at the source, so the patient rejoins the office's normal flow
      // instead of being re-appended within seconds by the next cycle.
      //
      // The order is not optional: the applier writes the fields back BEFORE it
      // clears the status and deletes the row, because the queue row is the only
      // place the fix exists until it has been copied.
      if (queueRow.resolved.kind === 'resolved') {
        intents.push({
          kind: 'remove-queue-row',
          syncId,
          queueSheet: queueRow.sheet,
          queueRow: queueRow.row,
          sourceSheet: dailyRow.sheet,
          sourceRow: dailyRow.row,
          reason: 'resolved-on-queue',
        });
      } else if (queueRow.resolved.kind === 'unrecognized') {
        // Somebody typed something meaning to say something. Acting on it would
        // be a guess; ignoring it silently would leave them thinking the row was
        // dealt with.
        unrecognizedResolved.push({
          queueSheet: queueRow.sheet,
          queueRow: queueRow.row,
          syncId,
          raw: queueRow.resolved.raw,
        });
      }
      continue;
    }

    // Terminal (ohi / lasante) or the keyword was removed at the source: the row
    // belongs on no queue at all, so every copy of it goes.
    for (const row of queueRows) removeStale(row);
  }

  // --- 3. Queue rows with no live source ------------------------------------
  for (const [syncId, queueRows] of queueById) {
    const dailyRow = dailyById.get(syncId);

    if (!dailyRow) {
      // Either the source sheet was outside this cycle's bounded scan, or the
      // source row's column B was cleared, or the row was deleted. Only the
      // scanned-sheet case is safe to act on, and we cannot tell them apart
      // here, so this is reported and never acted on. `selectSheetsToScan`
      // pulls in every sheet referenced by a queue row precisely so that a
      // genuine orphan is rare.
      for (const queueRow of queueRows) {
        orphans.push({
          queueSheet: queueRow.sheet,
          queueRow: queueRow.row,
          syncId,
          reason: 'unknown-sync-id',
        });
      }
      continue;
    }

    // Staff cleared the queue row: remove it and clear column B at the source.
    // The patient stays on their daily sheet, unflagged.
    //
    // OPEN QUESTION (see README): requirement 4 says "when staff clear column B
    // on the queue sheet". Column B on a *daily* sheet is Status, but on a queue
    // sheet the documented layout puts Source Row in column B and has no status
    // column at all. Until that is settled this uses the conservative reading —
    // the whole row blanked — because it cannot misfire, whereas keying deletion
    // off a single column would delete rows on a stray backspace.
    for (const queueRow of queueRows) {
      if (APPEND_ONLY_DESTINATIONS.includes(queueRow.sheet)) continue;
      if (!rowLooksCleared(queueRow)) continue;

      intents.push({
        kind: 'remove-queue-row',
        syncId,
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        sourceSheet: dailyRow.sheet,
        sourceRow: dailyRow.row,
        reason: 'cleared-on-queue',
      });
    }
  }

  return {
    intents,
    ambiguous,
    orphans,
    unrecognizedResolved,
    unrecognizedCounts,
    counts: countIntents(intents),
  };
}

function buildAppend(
  row: DailyRow,
  syncId: string,
  destination: QueueSheetName,
): AppendQueueRowIntent {
  const values: Partial<Record<QueueColumn, unknown>> = { ...row.fields };
  // `Date of Visit` is written as the daily sheet's own name (`July 30, 2026`)
  // rather than as an Excel serial. Both are accepted by `resolveSourceSheet`,
  // and the name needs no locale-dependent date parsing on the way in and tells
  // a human which tab to open on the way out.
  values['Date of Visit'] = row.sheet;
  // A plain row number, per the office's decision that Source Row carries no
  // hyperlink. It was briefly written as `July 30, 2026!B45`, which reads well
  // but is not a number: `planAdoption` parses this cell with `Number()`, so the
  // combined form would have made every row this project appends unadoptable and
  // left the column holding two different things.
  values['Source Row'] = row.row;

  const blankRequired = REQUIRED_FIELDS[destination].filter((field) =>
    isBlank(values[field]),
  );

  return {
    kind: 'append-queue-row',
    destination,
    syncId,
    sourceSheet: row.sheet,
    sourceRow: row.row,
    camp: row.camp,
    values,
    blankRequired,
  };
}

/**
 * A queue row counts as "cleared by staff" only when every patient field is
 * blank. A single blank field is a normal Missing Info row, not a deletion
 * request — treating it as one would delete exactly the patients this system
 * exists to chase.
 */
function rowLooksCleared(queueRow: QueueRow): boolean {
  for (const [column, value] of Object.entries(queueRow.values)) {
    if (QUEUE_ONLY_COLUMNS.includes(column as QueueColumn)) continue;
    if (!isBlank(value)) return false;
  }
  return true;
}

/**
 * Excel round-trips values through JSON as numbers, strings, booleans or null,
 * and a date can come back either way depending on the cell's format. Comparing
 * trimmed string forms avoids a write-back storm where the job rewrites the same
 * value every 5 seconds because `5` !== `"5"`.
 */
export function sameValue(a: unknown, b: unknown): boolean {
  if (isBlank(a) && isBlank(b)) return true;
  if (isBlank(a) || isBlank(b)) return false;
  return String(a).trim() === String(b).trim();
}

/**
 * Compares one field's value on the queue sheet against the daily sheet.
 *
 * Date columns need their own comparison, because the two sides now store a date
 * differently on purpose: the queue sheets hold an Excel serial so the column
 * can be sorted and filtered, and the daily sheets hold whatever text staff
 * typed. `40147` and `11/30/2009` are the same birthday and different strings.
 *
 * Compared as strings, every dated row would look like a staff edit on every
 * cycle — a write-back storm at a five-second cadence, and one that would
 * overwrite the daily sheets' text with serials.
 *
 * The date-aware path is deliberately limited to the columns that hold dates.
 * Applied to everything it would start reading zip codes and phone numbers as
 * serial numbers, and two different values would compare equal.
 */
export function sameFieldValue(field: QueueColumn, a: unknown, b: unknown): boolean {
  if (STYLE.dateColumns.includes(field)) {
    if (isBlank(a) && isBlank(b)) return true;
    if (isBlank(a) || isBlank(b)) return false;
    return sameDay(a, b) || sameValue(a, b);
  }
  return sameValue(a, b);
}

function countIntents(intents: Intent[]): Record<string, number> {
  const counts: Record<string, number> = {
    'stamp-id': 0,
    'append-queue-row': 0,
    'write-back': 0,
    'remove-queue-row': 0,
  };
  for (const intent of intents) {
    counts[intent.kind] = (counts[intent.kind] ?? 0) + 1;
  }
  return counts;
}

/** Groups appends by destination and camp, for logging and for Phase 2 batching. */
export function groupAppends(
  intents: Intent[],
): Map<QueueSheetName, Map<string, AppendQueueRowIntent[]>> {
  const byDestination = new Map<QueueSheetName, Map<string, AppendQueueRowIntent[]>>();
  for (const intent of intents) {
    if (intent.kind !== 'append-queue-row') continue;
    let byCamp = byDestination.get(intent.destination);
    if (!byCamp) {
      byCamp = new Map();
      byDestination.set(intent.destination, byCamp);
    }
    const key = campKey(intent.camp);
    const list = byCamp.get(key) ?? [];
    list.push(intent);
    byCamp.set(key, list);
  }
  return byDestination;
}
