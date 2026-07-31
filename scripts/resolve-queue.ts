/**
 * Phases 3 and 4, driven by the `Resolved` marker.
 *
 *   npm run resolve-rows                     # dry run, every queue
 *   npm run resolve-rows -- "Missing Info"   # dry run, one queue
 *   npm run resolve-rows -- --apply
 *
 * When a staff member fills in what was missing on a queue row and marks it
 * Done, three things have to happen and the order is not negotiable:
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
 * Nothing here acts on inference. A row is only touched when a human put one of
 * the accepted words in the Resolved column; anything else in that cell is
 * reported and left exactly as it is.
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

    // Only rows a human marked. `no-longer-queued-at-source` and
    // `cleared-on-queue` removals are a different decision and are not applied
    // here — this script exists to honour an explicit signal, nothing else.
    const resolved = plan.intents.filter(
      (intent): intent is RemoveQueueRowIntent =>
        intent.kind === 'remove-queue-row' &&
        intent.reason === 'resolved-on-queue' &&
        (!only || intent.queueSheet === only),
    );
    const resolvedRows = new Set(resolved.map((intent) => `${intent.queueSheet}!${intent.queueRow}`));

    // Write-backs belonging to those rows, and only those.
    const writeBacks = plan.intents.filter(
      (intent): intent is WriteBackIntent =>
        intent.kind === 'write-back' &&
        resolvedRows.has(`${intent.queueSheet}!${intent.queueRow}`),
    );

    // --- report -------------------------------------------------------------
    out();
    out(`queues:           ${only ?? 'all five'}`);
    out(`rows marked done: ${resolved.length}`);
    out(`fields to write back: ${writeBacks.length}`);
    out(`statuses to clear:    ${new Set(resolved.map((r) => `${r.sourceSheet}!${r.sourceRow}`)).size}`);
    out(`unrecognized markers: ${plan.unrecognizedResolved.length}`);

    if (resolved.length > 0) {
      out();
      out('rows to remove, by queue:');
      const byQueue = new Map<string, number>();
      for (const intent of resolved) {
        byQueue.set(intent.queueSheet, (byQueue.get(intent.queueSheet) ?? 0) + 1);
      }
      for (const [queue, count] of byQueue) out(`  ${queue.padEnd(24)} ${count}`);
    }

    for (const report of plan.unrecognizedResolved) {
      out(
        `  ${report.queueSheet} row ${report.queueRow}: ${JSON.stringify(report.raw)} ` +
          '— not an accepted marker, left alone',
      );
    }
    out();

    if (resolved.length === 0) {
      out('Nothing marked done. Nothing to do.');
      return;
    }

    if (!apply) {
      out('DRY RUN — nothing written. Re-run with --apply.');
      return;
    }

    log.warn('resolve.apply_mode', { count: resolved.length });

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
    const byTab = new Map<string, number[]>();
    for (const intent of resolved) {
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
      count: resolved.length,
      counts: { writeBacks: writeBacks.length, cleared, deleted },
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
