/**
 * Phase 2b — append the patients who have no queue row yet.
 *
 *   npm run append -- "United Refuah"                    # dry run, prints the plan
 *   npm run append -- "United Refuah" --apply            # writes
 *   npm run append -- "United Refuah" --apply --allow-inserts
 *
 * One destination per run, on purpose. The reconciler computes appends for all
 * five queues at once, but a queue at a time is what makes the first live write
 * reviewable: run `United Refuah` — empty, append-only, nothing written back to
 * a daily sheet — read the result, then do the next.
 *
 * Two things about the order of writes are load-bearing.
 *
 * **Daily SyncIDs are stamped BEFORE the queue rows are appended.** The
 * reconciler decides what to append by asking which source rows have no queue
 * row, and it can only ask that of a row carrying an ID. Stamp first and a crash
 * before the append leaves a stamped row with no queue row — which the next run
 * appends, once, under the same ID. Append first and a crash leaves an unstamped
 * row, which the next run mints a *new* ID for and appends again. Same failure,
 * opposite outcome: one is recoverable, the other silently duplicates patients.
 *
 * **Row inserts are refused unless asked for.** A plan that only writes past the
 * end of a sheet cannot misalign columns; a plan that opens gaps in the middle
 * can. `--allow-inserts` is how that becomes a decision rather than a surprise.
 */

import {
  LAYOUT,
  QUEUE_SHEET_NAMES,
  QUEUE_SHEET_TABS,
  assertLayoutVerified,
  loadConfig,
  type QueueSheetName,
} from '../src/config';
import { planQueueAppend, type QueueAppendPlan } from '../src/domain/append';
import { cellFromGrid, parseAddress } from '../src/domain/cells';
import { mapWithConcurrency } from '../src/domain/concurrency';
import { isDailySheetName, parseDailySheet } from '../src/domain/dailySheets';
import {
  assertSyncIdColumnIsClear,
  normalizeSheetName,
  parseQueueSheet,
  resolveSheetName,
} from '../src/domain/queueSheets';
import { reconcile, type AppendQueueRowIntent, type StampIdIntent } from '../src/domain/reconcile';
import { planColumnWrites } from '../src/domain/stamp';
import { newSyncId, normalizeSyncId } from '../src/domain/syncId';
import { consoleSink, createLogger, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';
import type { ParsedQueueSheet } from '../src/domain/queueSheets';

const out = (line = '') => process.stdout.write(`${line}\n`);

function resolveDestination(wanted: string): QueueSheetName {
  const target = normalizeSheetName(wanted);
  const match = QUEUE_SHEET_NAMES.find((name) => normalizeSheetName(name) === target);
  if (!match) {
    throw new Error(
      `"${wanted}" is not a queue. Choose one of: ${QUEUE_SHEET_NAMES.map((n) => `"${n}"`).join(', ')}`,
    );
  }
  return match;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const wanted = args.find((arg) => !arg.startsWith('--'));
  const apply = args.includes('--apply');
  const allowInserts = args.includes('--allow-inserts');

  if (!wanted) {
    throw new Error('Usage: npm run append -- "United Refuah" [--apply] [--allow-inserts]');
  }
  const destination = resolveDestination(wanted);

  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 2, dryRun: !apply });

  if (apply) assertLayoutVerified(config);
  assertSyncIdColumnIsClear();

  // The cycle's own logger would print a line per intent; this script prints its
  // own report instead.
  const silent = { debug() {}, info() {}, warn() {}, error() {} };
  const { workbook } = createSyncContext(silent, { phase: 2, dryRun: !apply });
  await workbook.ensureSession(apply);

  try {
    const worksheets = await workbook.listWorksheets();
    const names = worksheets.map((sheet) => sheet.name);

    const tab = resolveSheetName(QUEUE_SHEET_TABS[destination] ?? destination, names);
    if (!tab) {
      throw new Error(
        `No tab in the workbook matches "${QUEUE_SHEET_TABS[destination] ?? destination}". ` +
          'Create it first: npm run create-sheet -- "<name>" --apply',
      );
    }

    // --- read ---------------------------------------------------------------
    //
    // Every queue tab, not just the target: a source row already sitting on
    // another queue must not be appended to this one as well.
    const bindings: { status: QueueSheetName; tab: string }[] = [];
    for (const status of QUEUE_SHEET_NAMES) {
      const found = resolveSheetName(QUEUE_SHEET_TABS[status] ?? status, names);
      if (found) bindings.push({ status, tab: found });
    }

    const queues = await mapWithConcurrency(bindings, config.readConcurrency, async (binding) =>
      parseQueueSheet(binding.status, await workbook.getUsedRange(binding.tab)),
    );

    const dailyNames = names.filter((name) => isDailySheetName(name));
    // Existing SyncIDs for EVERY daily row, not only the queued ones. A range
    // write replaces every cell it covers, so a row inside a stamp block that is
    // not being stamped has to be written back with the ID it already holds —
    // and an unqueued row can perfectly well hold one.
    const existingDailyIds = new Map<string, Map<number, string>>();

    const daily = await mapWithConcurrency(dailyNames, config.readConcurrency, async (sheet) => {
      const used = await workbook.getUsedRange(sheet);
      const { startRow, startColumn } = parseAddress(used.address);
      const ids = new Map<number, string>();
      for (let i = 0; i < used.values.length; i++) {
        const row = startRow + i;
        if (row < LAYOUT.daily.firstDataRow) continue;
        const id = normalizeSyncId(
          cellFromGrid(used.values, startRow, startColumn, row, LAYOUT.daily.syncIdColumn),
        );
        if (id) ids.set(row, id);
      }
      existingDailyIds.set(sheet, ids);
      return parseDailySheet(sheet, used);
    });

    // --- plan ---------------------------------------------------------------
    const reconciled = reconcile({ daily, queues, newSyncId });

    const appends = reconciled.intents.filter(
      (intent): intent is AppendQueueRowIntent =>
        intent.kind === 'append-queue-row' && intent.destination === destination,
    );
    const wantedIds = new Set(appends.map((intent) => intent.syncId));
    const stamps = reconciled.intents.filter(
      (intent): intent is StampIdIntent =>
        intent.kind === 'stamp-id' && wantedIds.has(intent.syncId),
    );

    const target = queues.find((queue) => queue.sheet === destination);
    if (!target) throw new Error(`"${tab}" resolved but did not parse.`);

    const plan = planQueueAppend({ sheet: target, appends });

    // Daily-side stamp writes, one column, gaps wider than 50 rows split.
    const stampsBySheet = new Map<string, Map<number, string>>();
    for (const stamp of stamps) {
      const forSheet = stampsBySheet.get(stamp.sheet) ?? new Map<number, string>();
      forSheet.set(stamp.row, stamp.syncId);
      stampsBySheet.set(stamp.sheet, forSheet);
    }

    const stampWrites: { sheet: string; address: string; values: unknown[][] }[] = [];
    for (const [sheet, forSheet] of stampsBySheet) {
      for (const write of planColumnWrites({
        column: LAYOUT.daily.syncIdColumn,
        stamps: forSheet,
        existing: existingDailyIds.get(sheet) ?? new Map(),
      })) {
        stampWrites.push({ sheet, address: write.address, values: write.values });
      }
    }

    report({ destination, tab, target, plan, stampWrites, stamps: stamps.length });

    // --- gates --------------------------------------------------------------
    if (plan.appended === 0) {
      out('Nothing to append. Every queued patient for this destination already has a row.');
      return;
    }

    if (!target.shapeDetected) {
      throw new Error(
        `Could not find the header row on "${tab}", so the first data row is a guess ` +
          `(row ${target.firstDataRow}, from src/config.ts). Refusing to write into a sheet ` +
          'whose shape has not been read off the sheet itself.',
      );
    }

    if (!apply) {
      out('DRY RUN — nothing written. Re-run with --apply to write.');
      return;
    }

    if (plan.requiresInserts && !allowInserts) {
      throw new Error(
        `This plan opens ${plan.operations.filter((op) => op.kind === 'insert-rows').length} ` +
          'gap(s) in the middle of the sheet, which shifts live rows. Re-run with ' +
          '--allow-inserts once that is what you mean to do.',
      );
    }

    log.warn('append.apply_mode', { destination, count: plan.appended });

    // --- write --------------------------------------------------------------
    //
    // Stamps first. See the note at the top of this file: the crash-safe order
    // is the one where an interrupted run under-appends rather than duplicating.
    for (const write of stampWrites) {
      await workbook.writeRange(write.sheet, write.address, write.values);
    }
    out(`stamped ${stamps.length} source rows across ${stampWrites.length} range writes`);

    let done = 0;
    for (const operation of plan.operations) {
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
      done++;
      if (done % 25 === 0) out(`  ${done}/${plan.operations.length} operations done`);
    }

    log.info('append.complete', {
      destination,
      count: plan.appended,
      counts: { operations: plan.operations.length, stamps: stamps.length },
    });

    // --- verify -------------------------------------------------------------
    //
    // Reading back inside the same session is fresh; reading back without one
    // returns the pre-write picture and looks like the write silently failed.
    const after = parseQueueSheet(destination, await workbook.getUsedRange(tab));
    out();
    out('--- after ---');
    out(`patient rows:  ${after.rows.length}   (expected ${plan.totalAfter})`);
    out(`camp blocks:   ${after.groups.length}`);
    out(
      `TOTAL row:     ${after.totalRow ?? 'absent'}` +
        (after.totalRow ? ` declaring ${after.totalDeclared ?? '?'}` : ''),
    );
    out(`rows carrying a SyncID: ${after.rows.filter((row) => row.syncId).length}`);

    if (after.rows.length !== plan.totalAfter) {
      process.exitCode = 1;
      out();
      out('MISMATCH: the sheet does not hold the number of rows the plan expected.');
      out('Do not re-run until the tab has been looked at by hand.');
    }
  } catch (error) {
    log.error('append.failed', describeError(error));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

function report(args: {
  destination: QueueSheetName;
  tab: string;
  target: ParsedQueueSheet;
  plan: QueueAppendPlan;
  stampWrites: { sheet: string; address: string; values: unknown[][] }[];
  stamps: number;
}): void {
  const { plan, target } = args;
  const inserts = plan.operations.filter((op) => op.kind === 'insert-rows');

  out();
  out(`destination:   ${args.destination}`);
  out(`tab:           ${JSON.stringify(args.tab)}`);
  out(
    `shape:         header row ${target.headerRow}, data from ${target.firstDataRow}` +
      (target.shapeDetected ? ' (read off the sheet)' : ' (GUESSED from src/config.ts)'),
  );
  out(`existing:      ${target.rows.length} patient rows in ${target.groups.length} camp blocks`);
  out(`TOTAL row:     ${target.totalRow ?? 'absent — one will be written'}`);
  out();
  out(`to append:     ${plan.appended} patients`);
  out(`rows added:    ${plan.rowsAdded} (patients plus one divider per new camp)`);
  out(`total after:   ${plan.totalAfter}`);
  const stampCells = args.stampWrites.reduce((sum, write) => sum + write.values.length, 0);
  out(
    `source stamps: ${args.stamps} SyncIDs across ${args.stampWrites.length} range writes ` +
      `(${stampCells} cells, the rest written back unchanged)`,
  );
  out(`shading:       ${plan.shadedCells} blank required cells`);
  out(
    `row inserts:   ${inserts.length}` +
      (inserts.length === 0 ? ' — writes only past the end of the sheet' : ' — SHIFTS LIVE ROWS'),
  );
  out(`operations:    ${plan.operations.length}`);

  if (plan.placements.length > 0) {
    out();
    out('camp blocks:');
    for (const placement of [...plan.placements].sort((a, b) => a.atRow - b.atRow)) {
      out(
        `  ${placement.isNew ? 'new ' : 'grow'}  ${placement.camp.padEnd(24)} ` +
          `+${String(placement.added).padStart(3)}  (had ${placement.existing}, at row ${placement.atRow})`,
      );
    }
  }

  if (inserts.length > 0) {
    out();
    out('inserts, bottom of the sheet first:');
    for (const insert of inserts) {
      out(`  row ${String(insert.row).padStart(5)}  x${insert.count}  ${insert.reason}`);
    }
  }
  out();
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'append.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
