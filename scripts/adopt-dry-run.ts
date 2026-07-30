/**
 * Reports how many of the queue rows already sitting in the workbook can be
 * linked to their source row on a daily sheet.
 *
 *   npm run adopt
 *
 * Read-only. Writes nothing, ever — there is no --apply.
 *
 * This is the gate on Phase 2. Those rows are live, staff are editing them, and
 * they only keep working if each one is tied to the daily row it came from. If
 * a meaningful number cannot be linked, the old pointers have drifted and the
 * plan has to change before anything is written.
 */

import { LAYOUT, QUEUE_SHEET_NAMES, QUEUE_SHEET_TABS, type QueueColumn } from '../src/config';
import { planAdoption } from '../src/domain/adopt';
import { cellFromGrid, parseAddress } from '../src/domain/cells';
import { mapWithConcurrency } from '../src/domain/concurrency';
import { isDailySheetName } from '../src/domain/dailySheets';
import { parseQueueSheet, resolveSheetName } from '../src/domain/queueSheets';
import { classifyStatus } from '../src/domain/status';
import { newSyncId, normalizeSyncId } from '../src/domain/syncId';
import type { DailyRow } from '../src/domain/dailySheets';
import { consoleSink, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';

const out = (line = '') => process.stdout.write(`${line}\n`);

async function main(): Promise<void> {
  const silent = { debug() {}, info() {}, warn() {}, error() {} };
  const { workbook, config } = createSyncContext(silent, { phase: 1, dryRun: true });
  await workbook.ensureSession(false);

  try {
    const worksheets = await workbook.listWorksheets();
    const names = worksheets.map((sheet) => sheet.name);

    // Every row on every daily sheet, keyed by row number — not just the ones
    // carrying a status keyword. An old queue row may point at a row whose
    // column B has since been cleared or completed, and that still has to be
    // findable.
    const dailyNames = names.filter((name) => isDailySheetName(name));
    const dailyRowsBySheet = new Map<string, Map<number, DailyRow>>();

    await mapWithConcurrency(dailyNames, config.readConcurrency, async (sheet) => {
      const used = await workbook.getUsedRange(sheet);
      const { startRow, startColumn } = parseAddress(used.address);
      const rows = new Map<number, DailyRow>();

      for (let i = 0; i < used.values.length; i++) {
        const row = startRow + i;
        if (row < LAYOUT.daily.firstDataRow) continue;
        const at = (column: string) =>
          cellFromGrid(used.values, startRow, startColumn, row, column);

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

    out(`daily sheets read: ${dailyRowsBySheet.size}`);
    out(
      `rows indexed: ${[...dailyRowsBySheet.values()].reduce((sum, rows) => sum + rows.size, 0)}`,
    );
    out();

    let totalMatched = 0;
    let totalProblems = 0;
    const reasons = new Map<string, number>();
    const samples: string[] = [];

    for (const status of QUEUE_SHEET_NAMES) {
      const tab = resolveSheetName(QUEUE_SHEET_TABS[status] ?? status, names);
      if (!tab) {
        out(`${status.padEnd(24)} no tab`);
        continue;
      }

      const parsed = parseQueueSheet(status, await workbook.getUsedRange(tab));
      const result = planAdoption({
        queueRows: parsed.rows,
        daily: [...dailyRowsBySheet.keys()].map((sheet) => ({
          sheet,
          rows: [],
          unrecognizedRows: [],
        })),
        dailyRowsBySheet,
        newSyncId,
      });

      totalMatched += result.matches.length;
      totalProblems += result.problems.length;
      for (const problem of result.problems) {
        reasons.set(problem.reason, (reasons.get(problem.reason) ?? 0) + 1);
        if (samples.length < 15 && problem.reason === 'identity-mismatch') {
          samples.push(
            `  ${tab} row ${problem.queueRow} -> ${problem.sourceSheet} row ${problem.sourceRow}` +
              `  disagrees on: ${problem.mismatchedFields?.join(', ')}`,
          );
        }
      }

      const total = parsed.rows.length;
      const pct = total > 0 ? ((result.matches.length / total) * 100).toFixed(1) : '—';
      out(
        `${status.padEnd(24)} rows ${String(total).padStart(4)}   ` +
          `linked ${String(result.matches.length).padStart(4)} (${pct}%)   ` +
          `unlinked ${String(result.problems.length).padStart(4)}`,
      );
    }

    out();
    out(`TOTAL linked: ${totalMatched}   unlinked: ${totalProblems}`);
    if (reasons.size > 0) {
      out('\nwhy rows could not be linked:');
      for (const [reason, count] of [...reasons].sort((a, b) => b[1] - a[1])) {
        out(`  ${String(count).padStart(5)}  ${reason}`);
      }
    }
    if (samples.length > 0) {
      out('\nsample identity mismatches (field names only, no values):');
      samples.forEach((line) => out(line));
    }
  } finally {
    await workbook.closeSession();
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'adopt.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
