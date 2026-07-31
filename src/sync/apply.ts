import { LAYOUT, type QueueSheetName, type RuntimeConfig, type WorkbookLayout } from '../config';
import { planQueueAppend } from '../domain/append';
import { planColumnWrites } from '../domain/stamp';
import type { ParsedDailySheet } from '../domain/dailySheets';
import type { ParsedQueueSheet } from '../domain/queueSheets';
import type {
  AppendQueueRowIntent,
  ReconcilePlan,
  RemoveQueueRowIntent,
  StampIdIntent,
  WriteBackIntent,
} from '../domain/reconcile';
import { cellAddress } from '../domain/cells';
import type { Logger } from '../logging';
import type { Workbook } from '../graph/workbook';
import { applyRemovals } from './removals';

/**
 * Applying a reconciliation plan, in the only order that is safe.
 *
 * The phases are not a roadmap here, they are a gate: `config.phase` decides how
 * much of a plan is allowed to run, so the same code can be moved forward one
 * step at a time against a live workbook.
 *
 *   phase 2  stamp source rows, then append the queue rows they need
 *   phase 3  ... and carry staff edits back to the daily sheets
 *   phase 4  ... and remove rows that are finished, moved, or marked Done
 *
 * **Stamps before appends**, always. The reconciler decides what to append by
 * asking which source rows have no queue row, which it can only ask of a row
 * carrying an ID. Interrupted after stamping, the next cycle appends once under
 * the same ID; interrupted after appending, it mints a NEW ID and appends the
 * same patient again.
 *
 * **Appends before removals.** A patient moving queue produces a removal and an
 * append in one plan. Honour the removal first and, if the cycle dies in
 * between, they are on no queue at all.
 *
 * **Write-backs are carried by the removal applier** for rows being removed, so
 * that a row is never deleted before what is on it has been taken. Only
 * write-backs for rows that are staying are applied here.
 */

export interface ApplyPlanInput {
  workbook: Workbook;
  config: RuntimeConfig;
  plan: ReconcilePlan;
  daily: readonly ParsedDailySheet[];
  queues: readonly ParsedQueueSheet[];
  /** Queue name -> the tab it actually lives on. */
  tabFor: ReadonlyMap<QueueSheetName, string>;
  log: Logger;
  layout?: WorkbookLayout;
}

export interface ApplyPlanResult {
  stamped: number;
  appended: number;
  wroteBack: number;
  removed: number;
  skipped: number;
  /** True when anything at all was written. Drives the loop guard. */
  wrote: boolean;
}

export async function applyPlan(input: ApplyPlanInput): Promise<ApplyPlanResult> {
  const layout = input.layout ?? LAYOUT;
  const { workbook, plan, config, log, tabFor } = input;
  const result: ApplyPlanResult = {
    stamped: 0,
    appended: 0,
    wroteBack: 0,
    removed: 0,
    skipped: 0,
    wrote: false,
  };

  const stamps = plan.intents.filter((i): i is StampIdIntent => i.kind === 'stamp-id');
  const appends = plan.intents.filter(
    (i): i is AppendQueueRowIntent => i.kind === 'append-queue-row',
  );
  const writeBacks = plan.intents.filter((i): i is WriteBackIntent => i.kind === 'write-back');
  const removals = plan.intents.filter(
    (i): i is RemoveQueueRowIntent => i.kind === 'remove-queue-row',
  );

  // --- 1. Stamp the source rows --------------------------------------------
  const stampsBySheet = new Map<string, Map<number, string>>();
  for (const stamp of stamps) {
    const forSheet = stampsBySheet.get(stamp.sheet) ?? new Map<number, string>();
    forSheet.set(stamp.row, stamp.syncId);
    stampsBySheet.set(stamp.sheet, forSheet);
  }

  const existingBySheet = new Map(input.daily.map((sheet) => [sheet.sheet, sheet.syncIdsByRow]));

  for (const [sheet, forSheet] of stampsBySheet) {
    for (const write of planColumnWrites({
      column: layout.daily.syncIdColumn,
      stamps: forSheet,
      existing: existingBySheet.get(sheet) ?? new Map(),
    })) {
      await workbook.writeRange(sheet, write.address, write.values);
      result.wrote = true;
    }
    result.stamped += forSheet.size;
  }

  // --- 2. Append, one destination at a time --------------------------------
  //
  // Per destination, because the planner positions rows against one sheet's
  // camp blocks and its addresses are only valid for that sheet.
  for (const queue of input.queues) {
    const forQueue = appends.filter((intent) => intent.destination === queue.sheet);
    if (forQueue.length === 0) continue;

    const tab = tabFor.get(queue.sheet);
    if (!tab) continue;

    if (!queue.shapeDetected) {
      // The first data row is a guess. Refusing beats writing eight rows into
      // an instruction block.
      log.warn('apply.shape_not_detected', {
        queueSheet: queue.sheet,
        count: forQueue.length,
        reason: 'header row could not be read off the sheet; appends skipped',
      });
      result.skipped += forQueue.length;
      continue;
    }

    const queuePlan = planQueueAppend({ sheet: queue, appends: forQueue, layout });
    for (const operation of queuePlan.operations) {
      switch (operation.kind) {
        case 'insert-rows':
          await workbook.insertRows(tab, operation.row, operation.count);
          break;
        case 'write-cells':
          await workbook.writeRange(tab, operation.address, operation.values);
          break;
        case 'shade':
          await workbook.setFill(tab, operation.address, operation.color);
          break;
      }
      result.wrote = true;
    }
    result.appended += queuePlan.appended;
  }

  if (config.phase < 3) return result;

  // --- 3. Write-backs for rows that are STAYING ----------------------------
  //
  // A row being removed has its write-backs applied by the removal pass, one
  // row at a time and immediately before that row is deleted. Applying them
  // here as well would write the same cell twice, and — worse — would write it
  // long before the identity check that decides whether the row is still the
  // row.
  const removedRows = new Set(
    removals.map((intent) => `${intent.queueSheet}!${intent.queueRow}`),
  );
  const staying = writeBacks.filter(
    (intent) => !removedRows.has(`${intent.queueSheet}!${intent.queueRow}`),
  );

  for (const intent of staying) {
    const column = layout.daily.fieldColumns[intent.field];
    if (!column) continue;
    await workbook.writeRange(
      intent.sourceSheet,
      cellAddress(column, intent.sourceRow),
      [[intent.value ?? null]],
    );
    result.wroteBack++;
    result.wrote = true;
  }

  if (config.phase < 4) return result;

  // --- 4. Removals ---------------------------------------------------------
  //
  // A `wrong-queue` row whose replacement has not been appended yet is left
  // alone: deleting both sides of a move puts the patient on no queue. Since
  // appends ran above, that only happens when the append was skipped.
  const appendedIds = new Set(appends.map((intent) => intent.syncId));
  const safeToRemove = removals.filter((intent) => {
    if (intent.reason === 'cleared-on-queue') return false; // reported, never applied
    if (intent.reason === 'wrong-queue' && appendedIds.has(intent.syncId ?? '')) {
      return result.skipped === 0; // the append ran, so the replacement exists
    }
    return true;
  });

  const removalResult = await applyRemovals({
    workbook,
    removals: safeToRemove,
    writeBacks: writeBacks.filter((intent) =>
      removedRows.has(`${intent.queueSheet}!${intent.queueRow}`),
    ),
    tabFor: tabFor as ReadonlyMap<string, string>,
    log,
    layout,
  });

  result.removed = removalResult.deleted;
  result.wroteBack += removalResult.wrote;
  result.skipped += removalResult.skipped;
  if (removalResult.deleted > 0 || removalResult.wrote > 0) result.wrote = true;

  return result;
}
