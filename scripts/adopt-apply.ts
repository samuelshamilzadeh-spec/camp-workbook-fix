/**
 * Phase 2a — stamp SyncIDs onto the queue rows that already exist and onto the
 * daily rows they came from.
 *
 *   npm run adopt:apply              # dry run, reports every write it would make
 *   npm run adopt:apply -- --apply   # writes
 *
 * This is deliberately the FIRST thing that writes to the live workbook,
 * because it is the least dangerous write available:
 *
 *   - It only ever touches column BA, which is empty on every sheet. No patient
 *     field is written, moved, or read back into.
 *   - It moves no rows, so nothing can fall out of alignment.
 *   - It is reversible: clearing column BA returns the workbook to exactly its
 *     current state.
 *   - It is idempotent. A row that already carries the right ID is written back
 *     unchanged, so running it twice changes nothing.
 *
 * And it unlocks everything else: until both ends of a row share an ID, an edit
 * on a queue sheet cannot find its way back to the daily sheet.
 */

import { LAYOUT, QUEUE_SHEET_NAMES, QUEUE_SHEET_TABS, assertLayoutVerified, loadConfig, type QueueColumn } from '../src/config';
import { planAdoption } from '../src/domain/adopt';
import { cellFromGrid, parseAddress } from '../src/domain/cells';
import { mapWithConcurrency } from '../src/domain/concurrency';
import { isDailySheetName } from '../src/domain/dailySheets';
import type { DailyRow } from '../src/domain/dailySheets';
import { parseQueueSheet, resolveSheetName } from '../src/domain/queueSheets';
import { planColumnWrites, totalCells, type ColumnWrite } from '../src/domain/stamp';
import { classifyStatus } from '../src/domain/status';
import { newSyncId, normalizeSyncId } from '../src/domain/syncId';
import { consoleSink, createLogger, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';

const out = (line = '') => process.stdout.write(`${line}\n`);

async function main(): Promise<void> {
  const apply = process.argv.includes('--apply');
  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 2, dryRun: !apply });

  if (apply) {
    // The layout gate. Refuses unless someone has confirmed the LAYOUT block
    // against the real workbook.
    assertLayoutVerified(config);
    log.warn('adopt.apply_mode', { reason: 'writing SyncIDs to the live workbook' });
  }

  const silent = { debug() {}, info() {}, warn() {}, error() {} };
  const { workbook } = createSyncContext(silent, { phase: 2, dryRun: !apply });
  await workbook.ensureSession(apply);

  try {
    const worksheets = await workbook.listWorksheets();
    const names = worksheets.map((sheet) => sheet.name);
    const dailyNames = names.filter((name) => isDailySheetName(name));

    // Index every daily row, not only the queued ones: a queue row may point at
    // a row whose column B has since been cleared, and it still needs finding.
    const dailyRowsBySheet = new Map<string, Map<number, DailyRow>>();
    await mapWithConcurrency(dailyNames, config.readConcurrency, async (sheet) => {
      const used = await workbook.getUsedRange(sheet);
      const { startRow, startColumn } = parseAddress(used.address);
      const rows = new Map<number, DailyRow>();
      for (let i = 0; i < used.values.length; i++) {
        const row = startRow + i;
        if (row < LAYOUT.daily.firstDataRow) continue;
        const at = (column: string) => cellFromGrid(used.values, startRow, startColumn, row, column);
        const fields: Partial<Record<QueueColumn, unknown>> = {};
        for (const [queueColumn, dailyColumn] of Object.entries(LAYOUT.daily.fieldColumns)) {
          if (dailyColumn) fields[queueColumn as QueueColumn] = at(dailyColumn);
        }
        rows.set(row, {
          sheet,
          row,
          syncId: normalizeSyncId(at(LAYOUT.daily.syncIdColumn)),
          statusRaw: at(LAYOUT.daily.statusColumn),
          outcome: classifyStatus(at(LAYOUT.daily.statusColumn)),
          ambiguous: false,
          camp: String(at(LAYOUT.daily.campColumn) ?? '').trim() || undefined,
          fields,
        });
      }
      dailyRowsBySheet.set(sheet, rows);
      return null;
    });

    // --- plan -------------------------------------------------------------
    const queueStamps = new Map<string, Map<number, string>>();
    const queueExisting = new Map<string, Map<number, string>>();
    const dailyStamps = new Map<string, Map<number, string>>();
    let matched = 0;
    const problems = new Map<string, number>();

    for (const status of QUEUE_SHEET_NAMES) {
      const tab = resolveSheetName(QUEUE_SHEET_TABS[status] ?? status, names);
      if (!tab) continue;

      const parsed = parseQueueSheet(status, await workbook.getUsedRange(tab));
      const result = planAdoption({
        queueRows: parsed.rows,
        daily: [...dailyRowsBySheet.keys()].map((sheet) => ({ sheet, rows: [], unrecognizedRows: [] })),
        dailyRowsBySheet,
        newSyncId,
      });

      matched += result.matches.length;
      for (const problem of result.problems) {
        problems.set(problem.reason, (problems.get(problem.reason) ?? 0) + 1);
      }

      const qs = new Map<number, string>();
      for (const match of result.matches) {
        qs.set(match.queueRow, match.syncId);
        const ds = dailyStamps.get(match.sourceSheet) ?? new Map<number, string>();
        ds.set(match.sourceRow, match.syncId);
        dailyStamps.set(match.sourceSheet, ds);
      }
      queueStamps.set(tab, qs);

      const qe = new Map<number, string>();
      for (const row of parsed.rows) if (row.syncId) qe.set(row.row, row.syncId);
      queueExisting.set(tab, qe);
    }

    // --- turn into writes --------------------------------------------------
    const writes: { sheet: string; write: ColumnWrite }[] = [];

    for (const [tab, stamps] of queueStamps) {
      for (const write of planColumnWrites({
        column: LAYOUT.queue.syncIdColumn,
        stamps,
        existing: queueExisting.get(tab) ?? new Map(),
      })) {
        writes.push({ sheet: tab, write });
      }
    }

    for (const [sheet, stamps] of dailyStamps) {
      const existing = new Map<number, string>();
      for (const [row, daily] of dailyRowsBySheet.get(sheet) ?? []) {
        if (daily.syncId) existing.set(row, daily.syncId);
      }
      for (const write of planColumnWrites({
        column: LAYOUT.daily.syncIdColumn,
        stamps,
        existing,
      })) {
        writes.push({ sheet, write });
      }
    }

    const newStamps = writes.reduce((sum, w) => sum + w.write.newValues, 0);
    out();
    out(`adoption matches:      ${matched}`);
    out(`unlinked (untouched):  ${[...problems].map(([r, n]) => `${r}=${n}`).join(', ') || 'none'}`);
    out(`range writes planned:  ${writes.length}  across ${new Set(writes.map((w) => w.sheet)).size} sheets`);
    out(`cells written:         ${totalCells(writes.map((w) => w.write))}  (${newStamps} new IDs, rest written back unchanged)`);
    out(`column used:           queue ${LAYOUT.queue.syncIdColumn} / daily ${LAYOUT.daily.syncIdColumn}`);
    out();

    if (!apply) {
      out('DRY RUN — nothing written. Re-run with --apply to stamp.');
      return;
    }

    let done = 0;
    for (const { sheet, write } of writes) {
      await workbook.writeRange(sheet, write.address, write.values);
      done++;
      if (done % 10 === 0) out(`  ${done}/${writes.length} range writes done`);
    }
    log.info('adopt.complete', { count: newStamps, counts: { writes: writes.length } });
    out(`\nStamped ${newStamps} IDs across ${writes.length} range writes.`);
  } catch (error) {
    log.error('adopt.failed', describeError(error));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'adopt_apply.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
