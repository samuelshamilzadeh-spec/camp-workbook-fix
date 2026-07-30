import {
  LAYOUT,
  QUEUE_ONLY_COLUMNS,
  REQUIRED_FIELDS,
  type QueueColumn,
  type QueueSheetName,
  type WorkbookLayout,
} from '../config';
import { isBlank } from './cells';
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
  reason: 'cleared-on-queue' | 'no-longer-queued-at-source' | 'source-row-missing';
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
  reason: 'unknown-sync-id' | 'missing-sync-id';
}

export interface ReconcilePlan {
  intents: Intent[];
  ambiguous: AmbiguityReport[];
  orphans: OrphanReport[];
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
  const queueById = new Map<string, QueueRow>();
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
      queueById.set(row.syncId, row);
    }
  }

  // --- 1. New rows entering a queue -----------------------------------------
  for (const row of pendingStamps) {
    if (row.outcome.kind !== 'queued') continue;
    const syncId = input.newSyncId();
    intents.push({ kind: 'stamp-id', sheet: row.sheet, row: row.row, syncId });
    intents.push(buildAppend(row, syncId, row.outcome.destination, layout));
  }

  // --- 2. Already-stamped source rows ---------------------------------------
  for (const [syncId, dailyRow] of dailyById) {
    const queueRow = queueById.get(syncId);

    if (dailyRow.outcome.kind === 'queued') {
      if (!queueRow) {
        // Stamped but absent from its queue: append. Covers a row whose keyword
        // changed to a different destination as well, because the stale row is
        // removed by the pass below.
        intents.push(buildAppend(dailyRow, syncId, dailyRow.outcome.destination, layout));
        continue;
      }

      if (queueRow.sheet !== dailyRow.outcome.destination) {
        // Wrong queue. Remove here, append there — the append was emitted above
        // only when the row was missing entirely, so emit it now.
        intents.push({
          kind: 'remove-queue-row',
          syncId,
          queueSheet: queueRow.sheet,
          queueRow: queueRow.row,
          sourceSheet: undefined,
          sourceRow: undefined,
          reason: 'no-longer-queued-at-source',
        });
        intents.push(buildAppend(dailyRow, syncId, dailyRow.outcome.destination, layout));
        continue;
      }

      // Same queue: reconcile field values. A staff edit on the queue sheet is
      // authoritative and gets copied back to the daily sheet. Notes and the
      // Source Row link stay put.
      for (const [column, queueValue] of Object.entries(queueRow.values)) {
        const field = column as QueueColumn;
        if (QUEUE_ONLY_COLUMNS.includes(field)) continue;
        if (!layout.daily.fieldColumns[field]) continue;
        if (sameValue(queueValue, dailyRow.fields[field])) continue;

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
      continue;
    }

    // Terminal (ohi / lasante) or the keyword was removed at the source: the row
    // does not belong on any queue.
    if (queueRow) {
      intents.push({
        kind: 'remove-queue-row',
        syncId,
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        sourceSheet: undefined,
        sourceRow: undefined,
        reason: 'no-longer-queued-at-source',
      });
    }
  }

  // --- 3. Queue rows with no live source ------------------------------------
  for (const [syncId, queueRow] of queueById) {
    const dailyRow = dailyById.get(syncId);

    if (!dailyRow) {
      // Either the source sheet was outside this cycle's bounded scan, or the
      // source row's column B was cleared, or the row was deleted. Only the
      // scanned-sheet case is safe to act on, and we cannot tell them apart
      // here, so this is reported and never acted on. `selectSheetsToScan`
      // pulls in every sheet referenced by a queue row precisely so that a
      // genuine orphan is rare.
      orphans.push({
        queueSheet: queueRow.sheet,
        queueRow: queueRow.row,
        syncId,
        reason: 'unknown-sync-id',
      });
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
    if (rowLooksCleared(queueRow)) {
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
    unrecognizedCounts,
    counts: countIntents(intents),
  };
}

function buildAppend(
  row: DailyRow,
  syncId: string,
  destination: QueueSheetName,
  layout: WorkbookLayout,
): AppendQueueRowIntent {
  const values: Partial<Record<QueueColumn, unknown>> = { ...row.fields };
  values['Date of Visit'] = row.sheet;
  values['Source Row'] = `${row.sheet}!${layout.daily.statusColumn}${row.row}`;
  // Notes live only on the queue sheet and are never seeded from the source.
  values['Notes'] = '';

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
