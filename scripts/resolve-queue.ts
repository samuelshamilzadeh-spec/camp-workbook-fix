/**
 * Phases 3 and 4 — taking rows off a queue.
 *
 *   npm run resolve-rows                     # dry run, every queue
 *   npm run resolve-rows -- "Missing Info"   # dry run, one queue
 *   npm run resolve-rows -- --apply
 *
 * Two things put a row on the removal list, and they are not the same job.
 *
 * **A staff member marked it Done.** They filled in what was missing, so three
 * things have to happen and the order is not negotiable:
 *
 *   1. **Write the fix back to the daily sheet.** Until this runs, the queue row
 *      is the ONLY place that phone number exists. Deleting first would destroy
 *      the very work the marker is reporting.
 *   2. **Clear column B at the source.** The reconciler decides everything from
 *      what column B says right now, so a row still reading `missing info` is
 *      re-appended within seconds. Blank means "not processed yet", which puts
 *      the patient back into the office's normal flow to be entered into an EMR.
 *   3. **Delete the queue row.** Whole-row, bottom-up per sheet, so deleting one
 *      row cannot move another row out from under a later delete.
 *
 * **The source says they are finished.** Column B on the daily sheet went
 * terminal or was cleared there, so the daily sheet is already right, there is
 * nothing to write back and nothing to clear — the queue row is just out of
 * date. It is deleted. Without this the queues only ever grow.
 *
 * **The patient moved queue.** Their column B still flags them, but for a
 * different tab — nine live rows sat on Missing Info reading `no insurance` and
 * `incorrect insurance`, which route to Not Accepted and to Ineligible &
 * Inactive. The old copy is deleted ONLY once the new one exists. The
 * reconciler emits the append in the same plan and this script does not apply
 * appends, so a row whose replacement is still pending is skipped and picked up
 * on the next run. Delete both sides of a move and the patient is on no queue at
 * all.
 *
 * A third reason, `cleared-on-queue`, is reported and NEVER applied. It fires
 * when every patient field on a queue row is blank, which was the conservative
 * reading of "staff clear the row to signal removal" back when there was no
 * column to say so. The `Resolved` column says it explicitly now, and a row
 * that is blank because somebody is midway through retyping it looks exactly
 * like one they meant to delete.
 *
 * Nothing here acts on inference about a patient. A row is removed because a
 * human marked it, or because the daily sheet — the record of what actually
 * happened — no longer flags them.
 */

import {
  LAYOUT,
  QUEUE_SHEET_NAMES,
  QUEUE_SHEET_TABS,
  assertLayoutVerified,
  loadConfig,
  type QueueSheetName,
} from '../src/config';
import { cellAddress } from '../src/domain/cells';
import { mapWithConcurrency } from '../src/domain/concurrency';
import { isDailySheetName, parseDailySheet } from '../src/domain/dailySheets';
import { normalizeSheetName, parseQueueSheet, resolveSheetName } from '../src/domain/queueSheets';
import {
  reconcile,
  type RemoveQueueRowIntent,
  type WriteBackIntent,
} from '../src/domain/reconcile';
import { newSyncId } from '../src/domain/syncId';
import { consoleSink, createLogger, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';

const out = (line = '') => process.stdout.write(`${line}\n`);

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const wanted = args.find((arg) => !arg.startsWith('--'));
  const apply = args.includes('--apply');

  const only = wanted
    ? QUEUE_SHEET_NAMES.find((name) => normalizeSheetName(name) === normalizeSheetName(wanted))
    : undefined;
  if (wanted && !only) {
    throw new Error(
      `"${wanted}" is not a queue. Choose one of: ${QUEUE_SHEET_NAMES.map((n) => `"${n}"`).join(', ')}`,
    );
  }

  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 3, dryRun: !apply });
  if (apply) assertLayoutVerified(config);

  const silent = { debug() {}, info() {}, warn() {}, error() {} };
  const { workbook } = createSyncContext(silent, { phase: 3, dryRun: !apply });
  await workbook.ensureSession(apply);

  try {
    const names = (await workbook.listWorksheets()).map((sheet) => sheet.name);

    const bindings: { status: QueueSheetName; tab: string }[] = [];
    for (const status of QUEUE_SHEET_NAMES) {
      const tab = resolveSheetName(QUEUE_SHEET_TABS[status] ?? status, names);
      if (tab) bindings.push({ status, tab });
    }
    const tabFor = new Map(bindings.map((binding) => [binding.status, binding.tab]));

    const queues = await mapWithConcurrency(bindings, config.readConcurrency, async (binding) =>
      parseQueueSheet(binding.status, await workbook.getUsedRange(binding.tab)),
    );

    const dailyNames = names.filter((name) => isDailySheetName(name));
    const daily = await mapWithConcurrency(dailyNames, config.readConcurrency, async (sheet) =>
      parseDailySheet(sheet, await workbook.getUsedRange(sheet)),
    );

    const plan = reconcile({ daily, queues, newSyncId });

    const removals = plan.intents.filter(
      (intent): intent is RemoveQueueRowIntent =>
        intent.kind === 'remove-queue-row' && (!only || intent.queueSheet === only),
    );

    // A human marked it Done: write their fix back, clear the status at source,
    // delete the row.
    const resolved = removals.filter((intent) => intent.reason === 'resolved-on-queue');

    // The source says they are finished — column B went terminal or was cleared
    // there. Nothing to write back and nothing to clear; the daily sheet is
    // already right and the queue row is simply out of date. Without this the
    // queues only ever grow.
    const finished = removals.filter((intent) => intent.reason === 'no-longer-queued-at-source');

    // Still queued, on the wrong tab. Deleting one of these is only safe once
    // the replacement row exists on the correct tab — and the reconciler emits
    // that append in the SAME plan, which this script does not apply. So a row
    // whose replacement is still pending is skipped and picked up next run,
    // after `npm run append` has put it where it belongs. Delete both sides of
    // a move and the patient is on no queue at all.
    const pendingAppend = new Set(
      plan.intents.filter((i) => i.kind === 'append-queue-row').map((i) => i.syncId),
    );
    const movedAlready = removals.filter(
      (intent) => intent.reason === 'wrong-queue' && !pendingAppend.has(intent.syncId ?? ''),
    );
    const movePending = removals.filter(
      (intent) => intent.reason === 'wrong-queue' && pendingAppend.has(intent.syncId ?? ''),
    );

    // Deliberately NEVER applied. This fires when every patient field on a queue
    // row is blank, which the README calls the conservative reading of "staff
    // clear the row to signal removal". The `Resolved` column now says that
    // explicitly, so this heuristic is superseded — and a row that is blank
    // because somebody is midway through retyping it is indistinguishable from
    // one they meant to delete. Reported, never acted on.
    const clearedHeuristic = removals.filter((intent) => intent.reason === 'cleared-on-queue');

    const resolvedRows = new Set(resolved.map((intent) => `${intent.queueSheet}!${intent.queueRow}`));

    // Write-backs belonging to the marked rows, and only those. A field that
    // differs on a row nobody marked is not this script's business.
    const writeBacks = plan.intents.filter(
      (intent): intent is WriteBackIntent =>
        intent.kind === 'write-back' &&
        resolvedRows.has(`${intent.queueSheet}!${intent.queueRow}`),
    );

    // --- report -------------------------------------------------------------
    out();
    out(`queues:               ${only ?? 'all five'}`);
    out(`marked Done by staff: ${resolved.length}   (write back, clear column B, delete)`);
    out(`finished at source:   ${finished.length}   (delete only)`);
    out(`moved queue, replaced:${String(movedAlready.length).padStart(4)}   (delete the old copy)`);
    out(`moved, not yet placed:${String(movePending.length).padStart(4)}   (SKIPPED — run append first)`);
    out(`blank-row heuristic:  ${clearedHeuristic.length}   (reported, never applied)`);
    out(`fields to write back: ${writeBacks.length}`);
    out(`unrecognized markers: ${plan.unrecognizedResolved.length}`);

    const actionable = [...resolved, ...finished, ...movedAlready];

    if (actionable.length > 0) {
      out();
      out('rows to remove:');
      const byQueue = new Map<string, { done: number; finished: number; moved: number }>();
      const bump = (q: string, k: 'done' | 'finished' | 'moved') => {
        const e = byQueue.get(q) ?? { done: 0, finished: 0, moved: 0 };
        e[k]++;
        byQueue.set(q, e);
      };
      for (const i of resolved) bump(i.queueSheet, 'done');
      for (const i of finished) bump(i.queueSheet, 'finished');
      for (const i of movedAlready) bump(i.queueSheet, 'moved');
      for (const [queue, e] of byQueue) {
        out(`  ${queue.padEnd(24)} ${e.done} done, ${e.finished} finished, ${e.moved} moved`);
      }
    }

    for (const report of plan.unrecognizedResolved) {
      out(
        `  ${report.queueSheet} row ${report.queueRow}: ${JSON.stringify(report.raw)} ` +
          '— not an accepted marker, left alone',
      );
    }
    for (const intent of clearedHeuristic) {
      out(`  ${intent.queueSheet} row ${intent.queueRow}: every field blank — NOT removed`);
    }
    out();

    if (actionable.length === 0) {
      out('Nothing to remove.');
      return;
    }

    if (!apply) {
      out('DRY RUN — nothing written. Re-run with --apply.');
      return;
    }

    log.warn('resolve.apply_mode', {
      counts: { resolved: resolved.length, finished: finished.length, moved: movedAlready.length },
    });

    // --- 1. the fixes go home ------------------------------------------------
    //
    // Before anything is cleared or deleted. The queue row is the only copy.
    for (const intent of writeBacks) {
      const column = LAYOUT.daily.fieldColumns[intent.field];
      if (!column) continue;
      await workbook.writeRange(
        intent.sourceSheet,
        cellAddress(column, intent.sourceRow),
        [[intent.value ?? null]],
      );
    }
    out(`wrote back ${writeBacks.length} fields to the daily sheets`);

    // --- 2. clear the status at source ---------------------------------------
    // `/clear`, not a write of null. Graph accepts a values PATCH containing
    // nulls, returns 200, and leaves the cells exactly as they were — so this
    // step would have silently done nothing, the keyword would have stayed in
    // column B, and the patient would have been re-appended on every cycle
    // forever. See Workbook.clearRange.
    let cleared = 0;
    for (const intent of resolved) {
      if (!intent.sourceSheet || !intent.sourceRow) continue;
      await workbook.clearRange(
        intent.sourceSheet,
        cellAddress(LAYOUT.daily.statusColumn, intent.sourceRow),
        'Contents',
      );
      cleared++;
    }
    out(`cleared column B on ${cleared} source rows`);

    // --- 3. delete the queue rows -------------------------------------------
    //
    // Bottom-up within each tab. Deleting row 40 moves row 41 up to 40, so a
    // top-down pass would delete the wrong patient from the second row onwards.
    //
    // Both kinds are merged into ONE pass before sorting. Deleting the marked
    // rows first and the stale rows afterwards would leave the second set
    // holding row numbers from before the first set moved everything up — the
    // same failure, arrived at by being tidy.
    const byTab = new Map<string, number[]>();
    for (const intent of actionable) {
      const tab = tabFor.get(intent.queueSheet);
      if (!tab) continue;
      const rows = byTab.get(tab) ?? [];
      rows.push(intent.queueRow);
      byTab.set(tab, rows);
    }

    let deleted = 0;
    for (const [tab, rows] of byTab) {
      for (const row of [...new Set(rows)].sort((a, b) => b - a)) {
        await workbook.deleteRows(tab, row, 1);
        deleted++;
      }
    }
    out(`deleted ${deleted} queue rows`);

    log.info('resolve.complete', {
      count: actionable.length,
      counts: {
        resolved: resolved.length,
        finished: finished.length,
        moved: movedAlready.length,
        writeBacks: writeBacks.length,
        cleared,
        deleted,
      },
    });

    out();
    out('The camp divider counts and the TOTAL line are now stale on any tab that');
    out('lost a row. Re-run `npm run migrate -- "<queue>"` to redraw them.');
  } catch (error) {
    log.error('resolve.failed', describeError(error));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'resolve.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
