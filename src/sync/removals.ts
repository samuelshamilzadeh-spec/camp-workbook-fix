import { LAYOUT, type WorkbookLayout } from '../config';
import { cellAddress } from '../domain/cells';
import type { RemoveQueueRowIntent, WriteBackIntent } from '../domain/reconcile';
import { normalizeSyncId } from '../domain/syncId';
import type { Logger } from '../logging';

/**
 * Applying removals: the one place in this project that destroys data.
 *
 * Three properties, and each was a defect before it was a property.
 *
 * **Deepest row first, across every kind of removal at once.** Deleting row 40
 * moves row 41 up into it, so a top-down pass deletes the wrong patient from the
 * second row onwards. Doing the marked rows in one pass and the stale ones in
 * another has the same flaw wearing a tidier shape: the second pass holds row
 * numbers from before the first pass moved everything up.
 *
 * **One row at a time, all steps.** Clearing every source status and then
 * starting to delete meant a failure partway through left rows whose column B
 * was already blank and whose queue row still existed — and a source row with a
 * blank column B is invisible to the reconciler, so nothing could ever remove
 * those rows again. Interleaved, a failure strands exactly one.
 *
 * **Delete before clearing the status.** Both orders can fail in the middle;
 * only one fails safely. Clear-then-delete strands the row as above.
 * Delete-then-clear leaves a patient whose fix has already been written home and
 * whose source still carries its keyword, so the next cycle simply puts them
 * back on the queue — visibly, and with their record now complete.
 *
 * And before any of it, the row is checked for being the row it was. Every
 * number here was read minutes ago; staff are in this file continuously, and the
 * build brief records 43 source pointers drifting inside one 90-minute window.
 */

/** The slice of `Workbook` this needs. Narrow so a test can stand in for it. */
export interface RemovalWorkbook {
  getRange(sheetName: string, address: string): Promise<{ values: unknown[][] }>;
  writeRange(sheetName: string, address: string, values: unknown[][]): Promise<void>;
  deleteRows(sheetName: string, firstRow: number, count: number): Promise<void>;
  clearRange(sheetName: string, address: string, applyTo?: 'All' | 'Contents' | 'Formats'): Promise<void>;
}

export interface ApplyRemovalsInput {
  workbook: RemovalWorkbook;
  removals: readonly RemoveQueueRowIntent[];
  writeBacks: readonly WriteBackIntent[];
  /** Queue name -> the tab it actually lives on. */
  tabFor: ReadonlyMap<string, string>;
  log: Logger;
  layout?: WorkbookLayout;
}

export interface ApplyRemovalsResult {
  wrote: number;
  deleted: number;
  cleared: number;
  /** Rows whose SyncID no longer matched. Left alone, reported, retried next run. */
  skipped: number;
}

export async function applyRemovals(input: ApplyRemovalsInput): Promise<ApplyRemovalsResult> {
  const layout = input.layout ?? LAYOUT;
  const { workbook, tabFor, log } = input;

  const ordered = [...input.removals].sort((a, b) => {
    const tabA = tabFor.get(a.queueSheet) ?? '';
    const tabB = tabFor.get(b.queueSheet) ?? '';
    return tabA === tabB ? b.queueRow - a.queueRow : tabA.localeCompare(tabB);
  });

  const writeBacksFor = new Map<string, WriteBackIntent[]>();
  for (const intent of input.writeBacks) {
    const key = `${intent.queueSheet}!${intent.queueRow}`;
    writeBacksFor.set(key, [...(writeBacksFor.get(key) ?? []), intent]);
  }

  const result: ApplyRemovalsResult = { wrote: 0, deleted: 0, cleared: 0, skipped: 0 };

  for (const intent of ordered) {
    const tab = tabFor.get(intent.queueSheet);
    if (!tab) continue;

    // Is this still the row the plan was made about? A row inserted above it
    // since the read means this number now belongs to somebody else, and the
    // next two operations would clear their status and delete them.
    const check = await workbook.getRange(
      tab,
      cellAddress(layout.queue.syncIdColumn, intent.queueRow),
    );
    const actual = normalizeSyncId(check.values?.[0]?.[0]);
    if (!actual || actual !== intent.syncId) {
      result.skipped++;
      log.warn('resolve.row_moved', {
        queueSheet: intent.queueSheet,
        row: intent.queueRow,
        syncId: intent.syncId,
        reason: 'the SyncID at that row is not the one planned; the sheet shifted',
      });
      continue;
    }

    // 1. The fixes go home. Until this lands, the queue row is the only place
    //    they exist.
    for (const write of writeBacksFor.get(`${intent.queueSheet}!${intent.queueRow}`) ?? []) {
      const column = layout.daily.fieldColumns[write.field];
      if (!column) continue;
      await workbook.writeRange(
        write.sourceSheet,
        cellAddress(column, write.sourceRow),
        [[write.value ?? null]],
      );
      result.wrote++;
    }

    // 2. The row goes.
    await workbook.deleteRows(tab, intent.queueRow, 1);
    result.deleted++;

    // 3. The status is cleared — but only for a row a human marked Done. A row
    //    that moved queue or finished at source needs no clearing: its column B
    //    already says the right thing, and blanking it would pull the patient
    //    out of a queue they belong in.
    if (intent.reason === 'resolved-on-queue' && intent.sourceSheet && intent.sourceRow) {
      // `/clear`, not a write of null. A values PATCH containing null returns
      // 200 and leaves the cell exactly as it was, so this step would have
      // silently done nothing and the patient would return on every cycle.
      await workbook.clearRange(
        intent.sourceSheet,
        cellAddress(layout.daily.statusColumn, intent.sourceRow),
        'Contents',
      );
      result.cleared++;
    }
  }

  return result;
}
