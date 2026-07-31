import { LAYOUT, type QueueSheetName, type RuntimeConfig, type WorkbookLayout } from '../config';
import { planCountRefresh, planQueueAppend } from '../domain/append';
import { planColumnWrites } from '../domain/stamp';
import type { ParsedDailySheet } from '../domain/dailySheets';
import { parseQueueSheet, type ParsedQueueSheet } from '../domain/queueSheets';
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
import { planDividerStyle } from '../domain/style';
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
  /** Divider and TOTAL lines rewritten because the tab's row count changed. */
  recounted: number;
  /** Style operations applied to camp dividers this cycle created. */
  styled: number;
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
    recounted: 0,
    styled: 0,
    wrote: false,
  };

  // Tabs whose row layout this cycle changed, so the divider and TOTAL counts
  // on them can be brought back into line at the end.
  const touchedTabs = new Set<string>();
  /** Tab -> camps whose block this cycle created, which need styling. */
  const newCamps = new Map<string, string[]>();

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
  //
  // Iterating the DESTINATIONS that have appends, not the queues that happened
  // to parse. Iterating the parsed queues silently dropped every append bound
  // for a tab that could not be resolved — and a tab can stop resolving simply
  // by being renamed, which this office has already done once. Those appends
  // vanished without being counted as skipped, which let the removal guard
  // below conclude the replacement row existed.
  const appendedOk = new Set<string>();
  const queueBySheet = new Map(input.queues.map((queue) => [queue.sheet, queue]));
  const destinations = new Set(appends.map((intent) => intent.destination));

  for (const destination of destinations) {
    const forQueue = appends.filter((intent) => intent.destination === destination);
    const queue = queueBySheet.get(destination);
    const tab = tabFor.get(destination);

    if (!queue || !tab) {
      log.warn('apply.destination_unavailable', {
        queueSheet: destination,
        count: forQueue.length,
        reason: 'no tab in the workbook resolves to this queue; its rows have nowhere to go',
      });
      result.skipped += forQueue.length;
      continue;
    }

    if (!queue.shapeDetected) {
      // The first data row is a guess. Refusing beats writing eight rows into
      // an instruction block.
      log.warn('apply.shape_not_detected', {
        queueSheet: destination,
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
    for (const intent of forQueue) appendedOk.add(intent.syncId);
    touchedTabs.add(tab);
    // Camps that did not exist on this tab a moment ago. Their divider rows are
    // written as bare text and need dressing, or a new camp turns up as an
    // unbanded label between banded ones.
    for (const placement of queuePlan.placements) {
      if (placement.isNew) newCamps.set(tab, [...(newCamps.get(tab) ?? []), placement.camp]);
    }
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
  // A `wrong-queue` row is only safe to delete once THAT PATIENT'S replacement
  // exists. Asked per SyncID, not by a global counter: one destination failing
  // must not authorise deleting rows whose replacements did land, and — far
  // worse — must not be invisible to rows whose replacements did not.
  const appendPlanned = new Set(appends.map((intent) => intent.syncId));
  const safeToRemove = removals.filter((intent) => {
    if (intent.reason === 'cleared-on-queue') return false; // reported, never applied
    if (intent.reason !== 'wrong-queue') return true;

    const syncId = intent.syncId ?? '';
    // No append planned means they already have a row on the right queue.
    if (!appendPlanned.has(syncId)) return true;
    if (appendedOk.has(syncId)) return true;

    log.warn('apply.removal_deferred', {
      queueSheet: intent.queueSheet,
      row: intent.queueRow,
      syncId,
      reason: 'the replacement row was not written, so this is the only row this patient has',
    });
    return false;
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

  for (const intent of safeToRemove) {
    const tab = tabFor.get(intent.queueSheet);
    if (tab) touchedTabs.add(tab);
  }

  // --- 5. Bring the counts back into line ----------------------------------
  //
  // `Achim - 6 patients` is a snapshot taken when the row was written, which is
  // what the office asked for. It goes stale the moment a row is added or
  // removed, and nothing in a cycle was putting it right — only the migrate
  // script did, by hand. So the first time the timer removed a resolved row,
  // every divider above it started lying and stayed that way.
  //
  // Re-read rather than reason: the tab has just had rows inserted and deleted,
  // and the row numbers this cycle started with are exactly the ones that are no
  // longer true.
  for (const tab of touchedTabs) {
    const queue = [...queueBySheet.values()].find((q) => tabFor.get(q.sheet) === tab);
    if (!queue) continue;

    const fresh = parseQueueSheet(queue.sheet, await workbook.getUsedRange(tab), layout);
    for (const operation of planCountRefresh(fresh, layout)) {
      await workbook.writeRange(tab, operation.address, operation.values);
      result.wrote = true;
      result.recounted++;
    }

    // Dress the dividers this cycle created, and the TOTAL, from the same fresh
    // parse — the row numbers moved when the rows went in.
    const camps = newCamps.get(tab);
    if (!camps || camps.length === 0) continue;

    for (const operation of planDividerStyle(fresh, camps, layout)) {
      switch (operation.kind) {
        case 'fill':
          if (operation.fill) await workbook.setFill(tab, operation.address, operation.fill);
          break;
        case 'font':
          if (operation.font) await workbook.setFont(tab, operation.address, operation.font);
          break;
        case 'format':
          if (operation.format) await workbook.setRangeFormat(tab, operation.address, operation.format);
          break;
        case 'border':
          if (operation.border) {
            await workbook.setBorder(tab, operation.address, operation.border.edge, {
              style: operation.border.style,
              color: operation.border.color,
              weight: operation.border.weight,
            });
          }
          break;
        default:
          break;
      }
      result.wrote = true;
      result.styled++;
    }
  }

  return result;
}
