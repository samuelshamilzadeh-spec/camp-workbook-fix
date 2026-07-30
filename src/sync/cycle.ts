import {
  LAYOUT,
  QUEUE_SHEET_NAMES,
  assertLayoutVerified,
  type QueueSheetName,
  type RuntimeConfig,
} from '../config';
import { mapWithConcurrency } from '../domain/concurrency';
import { parseDailySheet, planScan } from '../domain/dailySheets';
import { parseQueueSheet, assertSyncIdColumnIsClear } from '../domain/queueSheets';
import { reconcile, type ReconcilePlan } from '../domain/reconcile';
import { newSyncId } from '../domain/syncId';
import { describeError, type Logger } from '../logging';
import type { StateStore, SyncState } from '../state/store';
import type { Workbook } from '../graph/workbook';

export interface CycleDeps {
  config: RuntimeConfig;
  workbook: Workbook;
  state: StateStore;
  log: Logger;
  /** Overrides the queue sheet names, for the parallel run under temporary names. */
  queueSheetNames?: readonly QueueSheetName[];
}

export type CycleResult =
  | { status: 'skipped'; reason: 'unchanged' | 'self-write'; durationMs: number }
  | {
      status: 'planned';
      plan: ReconcilePlan;
      scannedSheets: string[];
      applied: boolean;
      durationMs: number;
    };

/**
 * One reconciliation cycle.
 *
 * Order matters for cost. The `lastModifiedDateTime` check comes first and is
 * the only call most of the ~17,000 daily invocations make.
 */
export async function runCycle(deps: CycleDeps): Promise<CycleResult> {
  const { config, workbook, state, log } = deps;
  const startedAt = Date.now();
  const previous = await state.read();

  // --- Requirement 6: cheap change detection --------------------------------
  const metadata = await workbook.getFileMetadata();

  if (
    previous.lastSeenModified &&
    metadata.lastModifiedDateTime === previous.lastSeenModified
  ) {
    log.debug('cycle.skipped', { reason: 'unchanged', lastModified: metadata.lastModifiedDateTime });
    return { status: 'skipped', reason: 'unchanged', durationMs: Date.now() - startedAt };
  }

  // --- Requirement 5: loop guard --------------------------------------------
  //
  // Our own writes bump lastModifiedDateTime and would otherwise make the next
  // cycle look like a real change. Because the logic is state-based, doing the
  // work anyway is safe rather than destructive — but it is a write storm, and
  // at a 5-second cadence a write storm is a throttling incident.
  if (
    previous.lastSelfWriteModified &&
    metadata.lastModifiedDateTime === previous.lastSelfWriteModified
  ) {
    log.debug('cycle.skipped', { reason: 'self-write' });
    await state.write({ ...previous, lastSeenModified: metadata.lastModifiedDateTime });
    return { status: 'skipped', reason: 'self-write', durationMs: Date.now() - startedAt };
  }

  assertSyncIdColumnIsClear();

  // Read-only work uses a non-persisted session: faster than session-less calls
  // and structurally incapable of modifying the file.
  const willWrite = shouldApply(config);
  if (willWrite) assertLayoutVerified(config);

  await workbook.createSession(willWrite);

  try {
    const worksheets = await workbook.listWorksheets();
    const sheetNames = worksheets.map((sheet) => sheet.name);
    const queueNames = (deps.queueSheetNames ?? QUEUE_SHEET_NAMES).filter((name) =>
      sheetNames.includes(name),
    );

    // Queue sheets first: they tell us which daily sheets a bounded scan must
    // include beyond the most recent N.
    const queues = await mapWithConcurrency(queueNames, config.readConcurrency, async (name) =>
      parseQueueSheet(name, await workbook.getUsedRange(name)),
    );

    const referencedSheets = new Set<string>();
    for (const queue of queues) {
      for (const row of queue.rows) {
        const dateOfVisit = row.values['Date of Visit'];
        if (typeof dateOfVisit === 'string' && dateOfVisit.trim()) {
          referencedSheets.add(dateOfVisit.trim());
        }
      }
    }
    // --- Requirement 7: bounded scan ---------------------------------------
    //
    // Hot window every cycle, plus a rotating slice of everything else, so the
    // whole year is covered without reading every sheet every five seconds.
    const scan = planScan({
      allSheetNames: sheetNames,
      referencedSheets,
      today: new Date(),
      hotDaysBack: config.hotDaysBack,
      hotDaysForward: config.hotDaysForward,
      coldBatchSize: config.coldBatchSize,
      maxSheetsPerCycle: config.maxSheetsPerCycle,
      cursor: previous.scanCursor ?? 0,
    });

    const scannedSheets = scan.sheets;

    log.debug('scan.plan', {
      count: scannedSheets.length,
      sheets: scan.hot,
      skipped: scan.totalDaily - scannedSheets.length,
      counts: {
        hot: scan.hot.length,
        cold: scan.cold.length,
        totalDaily: scan.totalDaily,
        sweepCycles: scan.sweepCycles,
      },
      reason: scan.full ? 'full-scan' : 'tiered',
    });

    if (!scan.full) {
      // Crossing this line changes the freshness guarantee from "every sheet,
      // every cycle" to "every sheet within sweepCycles", so it is worth saying
      // out loud rather than discovering later.
      log.warn('scan.tiering_engaged', {
        count: scan.totalDaily,
        counts: { maxSheetsPerCycle: config.maxSheetsPerCycle, sweepCycles: scan.sweepCycles },
        reason: 'daily sheet count exceeds the full-scan threshold',
      });
    }

    if (scan.unparseable.length > 0) {
      // These still get scanned, on the tail of the rotation. But a daily sheet
      // whose name will not parse as a date cannot be placed in the hot window,
      // so it is only ever picked up on its rotation turn.
      log.warn('scan.unparseable_sheet_names', {
        count: scan.unparseable.length,
        sheets: scan.unparseable,
        reason: 'name matches the daily pattern but does not parse as a date',
      });
    }

    // Concurrent, because ~60 sequential reads at a few hundred milliseconds
    // each does not fit in a five-second cycle. See mapWithConcurrency.
    const daily = await mapWithConcurrency(
      scannedSheets,
      config.readConcurrency,
      async (name) => parseDailySheet(name, await workbook.getUsedRange(name)),
    );

    const plan = reconcile({ daily, queues, newSyncId });

    logPlan(log, plan, config);

    if (!willWrite) {
      const next: SyncState = {
        ...previous,
        lastSeenModified: metadata.lastModifiedDateTime,
        lastSeenETag: metadata.eTag,
        lastScannedSheets: scannedSheets,
        scanCursor: scan.nextCursor,
        lastFullCycleAt: new Date().toISOString(),
      };
      await state.write(next);
      return {
        status: 'planned',
        plan,
        scannedSheets,
        applied: false,
        durationMs: Date.now() - startedAt,
      };
    }

    // Phases 0 and 2-5 attach their appliers here. Nothing writes yet; see the
    // phase table in the README.
    throw new Error(
      `No applier is wired for phase ${config.phase}. Build phases in order; ` +
        'see "Build in this order" in the README.',
    );
  } catch (error) {
    log.error('cycle.failed', describeError(error));
    throw error;
  } finally {
    await workbook.closeSession();
  }
}

/** Writes happen only above phase 1 and only when dry run is off. */
export function shouldApply(config: RuntimeConfig): boolean {
  return config.phase >= 2 && !config.dryRun;
}

/**
 * Phase 1's deliverable: structured logging of intended changes.
 *
 * Summary counts go out at info on every cycle that found work. Per-intent
 * detail goes out at debug, because at 17,000 invocations a day an info line per
 * intent is the thing that generates a surprising Application Insights bill.
 * Ambiguities are the exception and are always warned about — they are the cells
 * most likely to represent a data entry problem.
 */
function logPlan(log: Logger, plan: ReconcilePlan, config: RuntimeConfig): void {
  const total = plan.intents.length;

  if (total > 0 || plan.ambiguous.length > 0 || plan.orphans.length > 0) {
    log.info('cycle.plan', {
      phase: config.phase,
      dryRun: config.dryRun,
      wouldWrite: shouldApply(config),
      counts: plan.counts,
      count: total,
    });
  }

  for (const report of plan.ambiguous) {
    log.warn('status.multi_match', {
      sheet: report.sheet,
      row: report.row,
      syncId: report.syncId,
      matched: report.matched,
      destination: report.destination,
    });
  }

  for (const orphan of plan.orphans) {
    log.warn('queue.orphan_row', {
      queueSheet: orphan.queueSheet,
      row: orphan.queueRow,
      syncId: orphan.syncId,
      reason: orphan.reason,
    });
  }

  for (const [sheet, count] of Object.entries(plan.unrecognizedCounts)) {
    log.info('status.unrecognized', { sheet, count });
  }

  for (const intent of plan.intents) {
    switch (intent.kind) {
      case 'stamp-id':
        log.debug('plan.stamp_id', {
          action: intent.kind,
          sheet: intent.sheet,
          row: intent.row,
          syncId: intent.syncId,
        });
        break;
      case 'append-queue-row':
        log.debug('plan.append', {
          action: intent.kind,
          destination: intent.destination,
          syncId: intent.syncId,
          sourceSheet: intent.sourceSheet,
          sourceRow: intent.sourceRow,
          camp: intent.camp,
          fields: intent.blankRequired,
        });
        break;
      case 'write-back':
        log.debug('plan.write_back', {
          action: intent.kind,
          syncId: intent.syncId,
          sourceSheet: intent.sourceSheet,
          sourceRow: intent.sourceRow,
          queueSheet: intent.queueSheet,
          field: intent.field,
        });
        break;
      case 'remove-queue-row':
        log.debug('plan.remove', {
          action: intent.kind,
          syncId: intent.syncId,
          queueSheet: intent.queueSheet,
          row: intent.queueRow,
          reason: intent.reason,
        });
        break;
    }
  }
}
