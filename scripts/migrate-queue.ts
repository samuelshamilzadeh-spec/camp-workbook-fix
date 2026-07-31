/**
 * Brings a queue tab up to the current schema and appearance.
 *
 *   npm run migrate -- "United Refuah"            # dry run, prints the plan
 *   npm run migrate -- "United Refuah" --apply
 *
 * Five things, in this order, because each depends on the one before:
 *
 *   1. Insert the `Resolved` column between `Source Row` and `Last Name`,
 *      shifting every later column right. Skipped for United Refuah, which is an
 *      append-only record with nothing to resolve, and skipped when the column
 *      is already there.
 *   2. Rewrite the header row from the configured column list.
 *   3. Convert `Date of Visit` and `Date of Birth` from text to Excel serials.
 *   4. Apply the styling: header bar, camp divider bands, total row, centred
 *      columns, date formats, column widths.
 *
 * The `Resolved` dropdown and the frozen header row are NOT done here: Graph
 * refuses both against this workbook. They are one-off manual settings — see
 * src/domain/style.ts.
 *
 * Step 1 is the only irreversible-ish one and the only one that touches patient
 * data — everything after it is formatting. It is a whole-column insert, so
 * nothing can shear: a patient's name and their insurance move together or not
 * at all.
 */

import {
  LAYOUT,
  QUEUE_SHEET_NAMES,
  QUEUE_SHEET_TABS,
  STYLE,
  assertLayoutVerified,
  loadConfig,
  queueColumnsFor,
  type QueueColumn,
  type QueueSheetName,
} from '../src/config';
import { cellFromGrid, offsetColumn, parseAddress, rangeAddress } from '../src/domain/cells';
import { toExcelSerial } from '../src/domain/dates';
import {
  detectQueueShape,
  lastQueueColumnLetter,
  normalizeSheetName,
  parseQueueSheet,
  resolveSheetName,
} from '../src/domain/queueSheets';
import { planCountRefresh } from '../src/domain/append';
import { planQueueStyle, type StyleOperation } from '../src/domain/style';
import { consoleSink, createLogger, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';
import type { Workbook } from '../src/graph/workbook';

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

/** Does the header row already carry `Resolved` where it belongs? */
function hasResolvedColumn(headerValues: unknown[], firstIndex: number): boolean {
  const label = String(headerValues[firstIndex + 2] ?? '')
    .trim()
    .toLowerCase();
  return label === 'resolved';
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const wanted = args.find((arg) => !arg.startsWith('--'));
  const apply = args.includes('--apply');

  if (!wanted) throw new Error('Usage: npm run migrate -- "United Refuah" [--apply]');
  const destination = resolveDestination(wanted);

  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 2, dryRun: !apply });
  if (apply) assertLayoutVerified(config);

  const silent = { debug() {}, info() {}, warn() {}, error() {} };
  const { workbook } = createSyncContext(silent, { phase: 2, dryRun: !apply });
  await workbook.ensureSession(apply);

  try {
    const names = (await workbook.listWorksheets()).map((sheet) => sheet.name);
    const tab = resolveSheetName(QUEUE_SHEET_TABS[destination] ?? destination, names);
    if (!tab) throw new Error(`No tab in the workbook matches "${destination}".`);

    const columns = queueColumnsFor(destination);
    const first = LAYOUT.queue.firstColumn;
    const firstIndex = 0;

    let used = await workbook.getUsedRange(tab);
    let shape = detectQueueShape(used);
    if (!shape.detected) {
      throw new Error(
        `Could not find the header row on "${tab}". Refusing to restructure a sheet whose ` +
          'shape has not been read off the sheet itself.',
      );
    }

    const headerValues = used.values[shape.headerRow - parseAddress(used.address).startRow] ?? [];
    const wantsResolved = columns.includes('Resolved');
    const alreadyHasResolved = hasResolvedColumn(headerValues, firstIndex);
    const needsColumnInsert = wantsResolved && !alreadyHasResolved;

    // --- report -------------------------------------------------------------
    out();
    out(`destination:   ${destination}`);
    out(`tab:           ${JSON.stringify(tab)}`);
    out(`usedRange:     ${used.address}`);
    out(`header row:    ${shape.headerRow} (data from ${shape.firstDataRow})`);
    out(`columns:       ${columns.length}  A..${lastQueueColumnLetter(destination)}`);
    out(
      `Resolved:      ${
        !wantsResolved
          ? 'not used on this tab (append-only record)'
          : alreadyHasResolved
            ? 'already present'
            : `WILL BE INSERTED at ${offsetColumn(first, 2)}, shifting everything right`
      }`,
    );

    if (!apply && needsColumnInsert) {
      out();
      out('The column insert changes where patient data lives, so the styling plan');
      out('below is computed against the sheet as it is TODAY. Re-run the dry run');
      out('after applying to see the final layout.');
    }

    // --- 1. insert the Resolved column --------------------------------------
    //
    // A whole-column insert is the right operation and not an in-memory reshuffle
    // of A..Q, because these tabs carry columns this schema knows nothing about
    // — `Ineligible & Inactive` has `Updated Insurance Carrier`, `Updated
    // Insurance ID #` and `Updated Medicaid #` out at R, S and T. Rewriting a
    // fixed window would overwrite them; an insert carries them along.
    //
    // It also carries along the SyncID column, which is the catch: BA shifts to
    // BB, and every link between a queue row and its patient's daily row is
    // suddenly in a column nothing reads. So the IDs are read first and put back
    // where they belong afterwards.
    if (apply && needsColumnInsert) {
      const syncIdColumn = LAYOUT.queue.syncIdColumn;
      const { startRow } = parseAddress(used.address);
      const lastRow = startRow + used.values.length - 1;
      const ids: unknown[][] = [];
      for (let row = startRow; row <= lastRow; row++) {
        ids.push([cellFromGrid(used.values, startRow, parseAddress(used.address).startColumn, row, syncIdColumn) ?? null]);
      }

      await workbook.insertColumns(tab, offsetColumn(first, 2), 1);
      log.info('migrate.column_inserted', { sheet: tab, column: offsetColumn(first, 2) });
      out(`\ninserted ${offsetColumn(first, 2)}`);

      const shifted = offsetColumn(syncIdColumn, 1);
      await workbook.writeRange(
        tab,
        rangeAddress(syncIdColumn, startRow, syncIdColumn, lastRow),
        ids,
      );
      // `/clear`, not a write of nulls: a values PATCH containing null returns
      // 200 and changes nothing. See Workbook.clearRange.
      await workbook.clearRange(tab, rangeAddress(shifted, startRow, shifted, lastRow), 'All');
      out(`moved ${ids.filter((v) => v[0]).length} SyncIDs back from ${shifted} to ${syncIdColumn}`);

      used = await workbook.getUsedRange(tab);
      shape = detectQueueShape(used);
    }

    // --- 2. the header row ---------------------------------------------------
    const headerAddress = rangeAddress(
      first,
      shape.headerRow,
      lastQueueColumnLetter(destination),
      shape.headerRow,
    );
    if (apply) {
      await workbook.writeRange(tab, headerAddress, [columns as unknown as unknown[]]);
      out(`header rewritten: ${headerAddress}`);
    }

    // --- 3. dates as dates ---------------------------------------------------
    const dateWrites = planDateWrites(used, shape.firstDataRow, destination);
    out(
      `date cells:    ${dateWrites.reduce((n, w) => n + w.values.length, 0)} to convert to serials ` +
        `across ${dateWrites.length} range writes`,
    );

    if (apply) {
      for (const write of dateWrites) {
        await workbook.writeRange(tab, write.address, write.values);
      }
    }

    // --- 4 and 5. styling and the dropdown ----------------------------------
    const parsed = parseQueueSheet(destination, apply ? await workbook.getUsedRange(tab) : used);
    const style = planQueueStyle({ sheet: parsed });

    // Divider and TOTAL counts are snapshots and go stale the moment a row is
    // removed, which happens every time somebody marks one Done. Recomputed from
    // what the sheet holds now, so migrate doubles as a redraw.
    const counts = planCountRefresh(parsed);
    out(`stale counts:  ${counts.length} divider/total lines to rewrite`);

    if (apply) {
      for (const operation of counts) {
        await workbook.writeRange(tab, operation.address, operation.values);
      }
    }

    out(`style ops:     ${style.operations.length} over ${style.rows} rows`);
    out(`camp bands:    ${parsed.groups.length}`);
    out(`total row:     ${parsed.totalRow ?? 'absent'}`);
    out();
    for (const [kind, n] of countBy(style.operations)) out(`  ${kind.padEnd(15)} ${n}`);
    out();

    if (!apply) {
      out('DRY RUN — nothing written. Re-run with --apply.');
      return;
    }

    let done = 0;
    for (const operation of style.operations) {
      try {
        await applyStyle(workbook, tab, operation);
      } catch (error) {
        // Graph error bodies are dropped by the logger on purpose — a failed
        // range write echoes the payload back, and that payload is patient data.
        // So a bare 400 says nothing about which of seventy-odd calls failed.
        // The operation's own description is safe to print and is the only thing
        // that makes this diagnosable.
        out(
          `\nFAILED on style op ${done + 1}/${style.operations.length}: ` +
            `${operation.kind} ${operation.address} — ${operation.what}`,
        );
        throw error;
      }
      done++;
      if (done % 25 === 0) out(`  ${done}/${style.operations.length} style ops done`);
    }

    log.info('migrate.complete', {
      sheet: tab,
      destination,
      counts: { style: style.operations.length, dates: dateWrites.length },
    });
    out(`\nDone. ${style.operations.length} style operations applied to ${JSON.stringify(tab)}.`);
  } catch (error) {
    log.error('migrate.failed', describeError(error));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

/**
 * Converts the date columns from whatever text they hold into Excel serials.
 *
 * Written one column at a time and only where the value actually parses as a
 * date: a range write replaces every cell it covers, so a cell holding something
 * unrecognizable is written back exactly as it was rather than blanked or
 * guessed at.
 */
function planDateWrites(
  used: { address: string; values: unknown[][] },
  firstDataRow: number,
  destination: QueueSheetName,
): { address: string; values: unknown[][] }[] {
  const { startRow, startColumn } = parseAddress(used.address);
  const columns = queueColumnsFor(destination);
  const first = LAYOUT.queue.firstColumn;
  const lastRow = startRow + used.values.length - 1;
  const writes: { address: string; values: unknown[][] }[] = [];

  for (const column of STYLE.dateColumns) {
    const index = columns.indexOf(column as QueueColumn);
    if (index === -1) continue;
    const letter = offsetColumn(first, index);

    const values: unknown[][] = [];
    let changed = false;
    for (let row = firstDataRow; row <= lastRow; row++) {
      const current = cellFromGrid(used.values, startRow, startColumn, row, letter);
      const serial = toExcelSerial(current);
      // A divider row's first cell is its label, not a date. toExcelSerial
      // returns undefined for it, so it is written back unchanged.
      if (serial !== undefined && serial !== current) changed = true;
      values.push([serial ?? (current ?? null)]);
    }

    if (changed && values.length > 0) {
      writes.push({
        address: rangeAddress(letter, firstDataRow, letter, lastRow),
        values,
      });
    }
  }

  return writes;
}

async function applyStyle(workbook: Workbook, tab: string, op: StyleOperation): Promise<void> {
  switch (op.kind) {
    case 'fill':
      if (op.fill) await workbook.setFill(tab, op.address, op.fill);
      return;
    case 'font':
      if (op.font) await workbook.setFont(tab, op.address, op.font);
      return;
    case 'format':
      if (op.format) await workbook.setRangeFormat(tab, op.address, op.format);
      return;
    case 'border':
      if (op.border) {
        await workbook.setBorder(tab, op.address, op.border.edge, {
          style: op.border.style,
          color: op.border.color,
          weight: op.border.weight,
        });
      }
      return;
    case 'number-format':
      if (op.numberFormat) {
        await workbook.setNumberFormat(
          tab,
          op.address,
          op.numberFormat.format,
          op.numberFormat.rows,
          op.numberFormat.columns,
        );
      }
      return;
  }
}

function countBy(operations: readonly StyleOperation[]): [string, number][] {
  const counts = new Map<string, number>();
  for (const operation of operations) {
    counts.set(operation.kind, (counts.get(operation.kind) ?? 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'migrate.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
