import {
  LAYOUT,
  QUEUE_SHEET_NAMES,
  QUEUE_SHEET_TABS,
  assertLayoutVerified,
  type QueueSheetName,
  type RuntimeConfig,
} from '../config';
import { mapWithConcurrency } from '../domain/concurrency';
import { parseDailySheet, planScan } from '../domain/dailySheets';
import {
  parseQueueSheet,
  assertSyncIdColumnIsClear,
  resolveSheetName,
} from '../domain/queueSheets';
import { reconcile, type ReconcilePlan } from '../domain/reconcile';
import { newSyncId } from '../domain/syncId';
import { applyPlan } from './apply';
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

  // Reused across cycles, never closed here. The cold open costs ~9s against
  // this workbook; every later read in the same session is ~200ms.
  const coldStart = !workbook.hasSession;
  await workbook.ensureSession(willWrite);
  if (coldStart) log.info('session.opened', { wouldWrite: willWrite });

  try {
    const worksheets = await workbook.listWorksheets();
    const sheetNames = worksheets.map((sheet) => sheet.name);
    // Each queue's canonical name maps to whatever the tab is actually called:
    // `Missing Info (New)`, and `Not Accepted ` with a trailing space. Exact
    // matching reported both as absent, so two of the four queues were silently
    // never read.
    const wanted = deps.queueSheetNames ?? QUEUE_SHEET_NAMES;
    const bindings: { status: QueueSheetName; tab: string }[] = [];
    const missingTabs: QueueSheetName[] = [];

    for (const status of wanted) {
      const tab = resolveSheetName(QUEUE_SHEET_TABS[status] ?? status, sheetNames);
      if (tab) bindings.push({ status, tab });
      else missingTabs.push(status);
    }

    if (missingTabs.length > 0) {
      // A queue with no tab cannot be populated, and its patients go nowhere.
      // Verify Insurance is in this state today: 271 rows and no sheet.
      log.warn('queue.sheet_missing', {
        count: missingTabs.length,
        sheets: missingTabs,
        reason: 'no matching tab in the workbook; these rows have nowhere to go',
      });
    }

    // Queue sheets first: they tell us which daily sheets a bounded scan must
    // include beyond the most recent N.
    const queues = await mapWithConcurrency(bindings, config.readConcurrency, async (binding) =>
      parseQueueSheet(binding.status, await workbook.getUsedRange(binding.tab)),
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

    // --- Apply -------------------------------------------------------------
    const applied = await applyPlan({
      workbook,
      config,
      plan,
      daily,
      queues,
      tabFor: new Map(bindings.map((binding) => [binding.status, binding.tab])),
      log,
    });

    log.info('cycle.applied', {
      phase: config.phase,
      count: applied.stamped + applied.appended + applied.wroteBack + applied.removed,
      counts: {
        stamped: applied.stamped,
        appended: applied.appended,
        wroteBack: applied.wroteBack,
        removed: applied.removed,
        skipped: applied.skipped,
      },
    });

    // Read the file's timestamp back so the NEXT cycle recognizes this write as
    // ours and does not treat it as a staff change. Without this the job wakes
    // itself every five seconds forever, which at this cadence is a throttling
    // incident rather than merely wasteful.
    const after = applied.wrote ? await workbook.getFileMetadata() : metadata;

    await state.write({
      ...previous,
      lastSeenModified: after.lastModifiedDateTime,
      lastSeenETag: after.eTag,
      lastSelfWriteModified: applied.wrote
        ? after.lastModifiedDateTime
        : previous.lastSelfWriteModified,
      lastScannedSheets: scannedSheets,
      scanCursor: scan.nextCursor,
      lastFullCycleAt: new Date().toISOString(),
    });

    return {
      status: 'planned',
      plan,
      scannedSheets,
      applied: true,
      durationMs: Date.now() - startedAt,
    };
  } catch (error) {
    log.error('cycle.failed', describeError(error));
    throw error;
  }
  // Deliberately no closeSession here: see Workbook.ensureSession. Graph
  // expires an idle session on its own, and the next cycle reopens one.
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

  for (const report of plan.unrecognizedResolved) {
    // Somebody typed something in Resolved meaning to say something. Acting on
    // it would be a guess, so it is warned about and left alone.
    log.warn('queue.resolved_unrecognized', {
      queueSheet: report.queueSheet,
      row: report.queueRow,
      syncId: report.syncId,
      keyword: report.raw,
    });
  }

  for (const report of plan.blankedOnQueue) {
    // A field blank on the queue and filled at the source. Never written back —
    // an empty string clears a cell, and the daily sheet is the billing copy.
    log.warn('queue.blank_not_written_back', {
      queueSheet: report.queueSheet,
      row: report.queueRow,
      syncId: report.syncId,
      field: report.field,
    });
  }

  for (const report of plan.unmarkedEdits) {
    // A queue value differing from its source on a row nobody marked Resolved.
    // Never written — the reconciler cannot tell a staff edit from a stale
    // mirror value, and about 700 rows were adopted from mirrors that stopped
    // updating a month ago. Surfaced so the difference is visible to a human.
    log.info('queue.unmarked_edit', {
      queueSheet: report.queueSheet,
      row: report.queueRow,
      syncId: report.syncId,
      field: report.field,
      reason: report.fillsABlank ? 'would fill a blank' : 'would replace a value',
    });
  }

  for (const report of plan.duplicateSources) {
    // Two daily rows sharing one ID, almost always a copied row. Neither is
    // reconciled, because picking one silently hides the other visit forever.
    log.warn('source.duplicate_sync_id', {
      sheet: report.sheet,
      row: report.row,
      syncId: report.syncId,
    });
  }

  for (const report of plan.wouldDuplicate) {
    log.warn('queue.append_suppressed', {
      queueSheet: report.queueSheet,
      row: report.queueRow,
      syncId: report.syncId,
      reason: report.reason,
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
