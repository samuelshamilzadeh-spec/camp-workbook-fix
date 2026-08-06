/**
 * Splits a queue tab into one tab per month of the season.
 *
 *   npm run split -- "Missing Info"                      # dry run: audit + plan
 *   npm run split -- "Missing Info" --apply              # moves the rows
 *   npm run split -- "Missing Info" --apply --archive    # ... and retires the source tab
 *
 * The office asked to work `Missing Info` and `Ineligible & Inactive` a month at
 * a time. A row's month is the month of its VISIT, so `Missing Info` becomes
 * `Missing Info - June`, `- July`, `- August`, and `- Other` for a visit outside
 * those months or a date that will not parse.
 *
 * Three things shape how this is done.
 *
 * **Every fill on these tabs is derived, so it is regenerated rather than
 * copied.** The red is `BLANK_REQUIRED_FILL` on a blank required field; the
 * bands are camp dividers and the TOTAL; the bar is the header. All of it comes
 * out of `REQUIRED_FIELDS` and `planQueueStyle`, so re-deriving it on the
 * destination reproduces the source exactly — and costs four Graph calls per tab
 * instead of one per cell. What that would miss is anything a person coloured by
 * hand, so the run does not assume: it PROBES the source tab's fills first and
 * refuses to write if it finds one the rules do not predict.
 *
 * **The move is idempotent.** A row whose SyncID is already on the destination is
 * skipped, and an unstamped row already there is matched on identity the same
 * way `reconcile` does it. An interrupted run is re-run, not unpicked.
 *
 * **The source tab is not touched until the destination has been read back.**
 * Rows are copied, the result is verified row by row, and only then does
 * `--archive` rename the original to `... (old)` and hide it. Nothing is ever
 * deleted: the archived tab still holds every row exactly as it stood, which is
 * the only copy of that state.
 */

import {
  LAYOUT,
  QUEUE_SHEET_TABS,
  REQUIRED_FIELDS,
  SEGMENTS,
  SPLIT_FAMILIES,
  STYLE,
  assertLayoutVerified,
  familyOf,
  loadConfig,
  queueColumnsFor,
  queueTabFor,
  type QueueColumn,
  type QueueSheetName,
  type Segment,
  type SplitFamily,
} from '../src/config';
import { BLANK_REQUIRED_FILL } from '../src/config';
import { planQueueAppend } from '../src/domain/append';
import {
  columnToIndex,
  indexToColumn,
  isBlank,
  offsetColumn,
  rangeAddress,
} from '../src/domain/cells';
import { mapWithConcurrency } from '../src/domain/concurrency';
import { identityKey } from '../src/domain/adopt';
import { campKey } from '../src/domain/dailySheets';
import {
  assertQueueTabNamesFit,
  lastQueueColumnLetter,
  normalizeSheetName,
  parseQueueSheet,
  resolveSheetName,
  type ParsedQueueSheet,
  type QueueRow,
} from '../src/domain/queueSheets';
import { segmentForVisit } from '../src/domain/segments';
import { planQueueStyle, type StyleOperation } from '../src/domain/style';
import type { AppendQueueRowIntent } from '../src/domain/reconcile';
import { consoleSink, createLogger, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';
import type { Workbook } from '../src/graph/workbook';

const out = (line = '') => process.stdout.write(`${line}\n`);

/** Where a family's rows live before the split, and where they go afterwards. */
function sourceTabNames(family: SplitFamily): string[] {
  // The pre-split tab, then the archived form — so a re-run after `--archive`
  // still finds the rows rather than reporting the family as absent.
  return [family, `${family} (old)`];
}

function resolveFamily(wanted: string): SplitFamily {
  const target = normalizeSheetName(wanted);
  const match = SPLIT_FAMILIES.find((name) => normalizeSheetName(name) === target);
  if (!match) {
    throw new Error(
      `"${wanted}" is not a split queue. Choose one of: ` +
        SPLIT_FAMILIES.map((n) => `"${n}"`).join(', '),
    );
  }
  return match;
}

// --- The fill audit ---------------------------------------------------------

/**
 * A range whose colour the style rules can predict, and what they predict.
 *
 * `expected: undefined` means "no fill". An unfilled cell reads back as white,
 * and nothing in `STYLE` ever deliberately fills a cell white, so the two are
 * treated as the same thing.
 */
interface FillProbe {
  address: string;
  row: number;
  firstColumn: string;
  lastColumn: string;
  expected: string | undefined;
  what: string;
}

interface FillFinding {
  address: string;
  what: string;
  expected: string;
  found: string;
}

const NO_FILL = 'none';

function normalizeColor(color: string | undefined): string {
  if (!color) return NO_FILL;
  const upper = color.trim().toUpperCase();
  return upper === '#FFFFFF' || upper === 'FFFFFF' ? NO_FILL : upper;
}

/**
 * Every range on the source tab whose colour is predictable, in probe order.
 *
 * A patient row is cut into runs of "should be red" and "should be plain" rather
 * than probed cell by cell: a row with no blank required fields is one call, and
 * the usual one-or-two-blanks row is three. Probing all 18 cells of 195 rows
 * would be 3,500 calls to answer a question three usually settles.
 */
function planFillProbes(sheet: ParsedQueueSheet): FillProbe[] {
  const columns = queueColumnsFor(sheet.sheet);
  const required = REQUIRED_FIELDS[familyOf(sheet.sheet)];
  const first = LAYOUT.queue.firstColumn;
  const last = lastQueueColumnLetter(sheet.sheet);
  const probes: FillProbe[] = [];

  const add = (row: number, from: number, to: number, expected: string | undefined, what: string) => {
    const firstColumn = offsetColumn(first, from);
    const lastColumn = offsetColumn(first, to);
    probes.push({
      address: rangeAddress(firstColumn, row, lastColumn, row),
      row,
      firstColumn,
      lastColumn,
      expected,
      what,
    });
  };

  probes.push({
    address: rangeAddress(first, sheet.headerRow, last, sheet.headerRow),
    row: sheet.headerRow,
    firstColumn: first,
    lastColumn: last,
    expected: STYLE.headerFill,
    what: 'header bar',
  });

  for (const group of sheet.groups) {
    probes.push({
      address: rangeAddress(first, group.headerRow, last, group.headerRow),
      row: group.headerRow,
      firstColumn: first,
      lastColumn: last,
      expected: STYLE.dividerFill,
      what: `divider band: ${group.camp}`,
    });
  }

  if (sheet.totalRow !== undefined) {
    probes.push({
      address: rangeAddress(first, sheet.totalRow, last, sheet.totalRow),
      row: sheet.totalRow,
      firstColumn: first,
      lastColumn: last,
      expected: STYLE.totalFill,
      what: 'total row',
    });
  }

  for (const row of sheet.rows) {
    // Which of this row's cells the rules say are red, as column indexes.
    const red = new Set<number>();
    for (const field of required) {
      if (!isBlank(row.values[field])) continue;
      const index = columns.indexOf(field as QueueColumn);
      if (index !== -1) red.add(index);
    }

    // Walk the row, emitting one probe per run of same-expectation cells.
    let runStart = 0;
    for (let index = 1; index <= columns.length; index++) {
      const same = index < columns.length && red.has(index) === red.has(runStart);
      if (same) continue;
      const isRed = red.has(runStart);
      add(
        row.row,
        runStart,
        index - 1,
        isRed ? BLANK_REQUIRED_FILL : undefined,
        isRed ? 'blank required (red)' : 'patient cells (plain)',
      );
      runStart = index;
    }
  }

  return probes;
}

/**
 * Probes the planned ranges and reports every fill the rules do not predict.
 *
 * A multi-cell range whose cells disagree comes back with no colour, which for a
 * range means "mixed" and for a single cell means "unfilled". So a mixed answer
 * is not a finding, it is an instruction to look closer: the run is re-probed one
 * cell at a time, where the answer is unambiguous.
 */
async function auditFills(
  workbook: Workbook,
  tab: string,
  sheet: ParsedQueueSheet,
  concurrency: number,
): Promise<{ findings: FillFinding[]; probes: number }> {
  /** Splits a multi-cell probe into one probe per cell, same expectation. */
  const cellsOf = (probe: FillProbe): FillProbe[] => {
    const from = columnToIndex(probe.firstColumn);
    const to = columnToIndex(probe.lastColumn);
    const cells: FillProbe[] = [];
    for (let index = from; index <= to; index++) {
      const letter = indexToColumn(index);
      cells.push({
        ...probe,
        address: rangeAddress(letter, probe.row, letter, probe.row),
        firstColumn: letter,
        lastColumn: letter,
      });
    }
    return cells;
  };

  const isRange = (probe: FillProbe): boolean => probe.firstColumn !== probe.lastColumn;

  const check = async (probe: FillProbe): Promise<{ color: string | undefined }> =>
    workbook.getFill(tab, probe.address);

  const verdict = (probe: FillProbe, color: string | undefined): FillFinding | undefined => {
    const found = normalizeColor(color);
    const expected = normalizeColor(probe.expected);
    return found === expected
      ? undefined
      : { address: probe.address, what: probe.what, expected, found };
  };

  // Pass one: every planned range. A range whose cells disagree answers with no
  // colour at all, which is not a finding — it is an instruction to look closer.
  const planned = planFillProbes(sheet);
  const firstPass = await mapWithConcurrency(planned, concurrency, check);

  const findings: FillFinding[] = [];
  const mixed: FillProbe[] = [];
  planned.forEach((probe, index) => {
    const color = firstPass[index]!.color;
    if (color === undefined && isRange(probe)) {
      mixed.push(probe);
      return;
    }
    const finding = verdict(probe, color);
    if (finding) findings.push(finding);
  });

  // Pass two: the disagreeing ranges, one cell at a time. Done as a second pass
  // rather than by recursing inside the first, so the number of requests in
  // flight stays bounded by `concurrency` instead of squaring it.
  const cells = mixed.flatMap(cellsOf);
  const secondPass = await mapWithConcurrency(cells, concurrency, check);
  cells.forEach((probe, index) => {
    const finding = verdict(probe, secondPass[index]!.color);
    if (finding) findings.push(finding);
  });

  return { findings, probes: planned.length + cells.length };
}

// --- Moving the rows --------------------------------------------------------

/**
 * A source row as something the appender can write.
 *
 * `blankRequired` is recomputed rather than carried across, because it is what
 * reproduces the red shading on the destination and it must describe the row as
 * it is NOW — a field somebody filled in since the shading was applied should
 * arrive unshaded.
 *
 * A row with no SyncID keeps none: `''` writes an empty cell, so it lands on the
 * new tab exactly as unstamped as it left, and shows up in the same orphan
 * report it always did. Inventing an ID here would silently link it to a daily
 * row nothing had matched it to.
 */
function toIntent(row: QueueRow, destination: QueueSheetName): AppendQueueRowIntent {
  const required = REQUIRED_FIELDS[familyOf(destination)];
  return {
    kind: 'append-queue-row',
    destination,
    syncId: row.syncId ?? '',
    sourceSheet: String(row.values['Date of Visit'] ?? ''),
    sourceRow: Number(row.values['Source Row'] ?? 0),
    camp: row.camp,
    values: row.values,
    blankRequired: required.filter((field) => isBlank(row.values[field])),
  };
}

/** Groups a tab's rows by the month of their visit, in sheet order. */
function bucketBySegment(sheet: ParsedQueueSheet): Map<Segment, QueueRow[]> {
  const buckets = new Map<Segment, QueueRow[]>();
  for (const segment of SEGMENTS) buckets.set(segment, []);
  for (const row of sheet.rows) {
    const segment = segmentForVisit(row.values['Date of Visit']);
    buckets.get(segment)!.push(row);
  }
  return buckets;
}

/**
 * Rows already on the destination, so a re-run adds nothing twice.
 *
 * Keyed both ways, because both kinds of row can already be there: by SyncID for
 * a stamped one, and by identity for an unstamped one that has no ID to be
 * recognized by. Missing the second is how the reconciler used to plan a second
 * physical row for a patient who was already sitting on the tab.
 */
function alreadyThere(destination: ParsedQueueSheet): { ids: Set<string>; identities: Set<string> } {
  const ids = new Set<string>();
  const identities = new Set<string>();
  for (const row of destination.rows) {
    if (row.syncId) ids.add(row.syncId);
    const key = identityKey(row.values);
    if (key) identities.add(key);
  }
  return { ids, identities };
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

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const wanted = args.find((arg) => !arg.startsWith('--'));
  const apply = args.includes('--apply');
  const archive = args.includes('--archive');
  const ignoreFills = args.includes('--ignore-unexpected-fills');

  if (!wanted) {
    throw new Error(
      'Usage: npm run split -- "Missing Info" [--apply] [--archive] ' +
        '[--ignore-unexpected-fills]',
    );
  }
  const family = resolveFamily(wanted);

  // Before anything else: this script is the one that creates tabs, so a name
  // Excel would refuse must fail here rather than halfway through a live move.
  assertQueueTabNamesFit();

  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 2, dryRun: !apply });
  if (apply) assertLayoutVerified(config);

  const silent = { debug() {}, info() {}, warn() {}, error() {} };
  const { workbook } = createSyncContext(silent, { phase: 2, dryRun: !apply });
  await workbook.ensureSession(apply);

  try {
    const names = (await workbook.listWorksheets()).map((sheet) => sheet.name);

    const sourceTab = sourceTabNames(family)
      .map((name) => resolveSheetName(name, names))
      .find((name): name is string => name !== undefined);
    if (!sourceTab) {
      throw new Error(
        `No tab in the workbook matches "${family}" or "${family} (old)". ` +
          'Nothing to split.',
      );
    }

    // Parsed under the family's own name: `familyOf` maps it to itself, so the
    // column list and required fields come out right for a tab that predates
    // the split.
    const sourceName = queueTabFor(family, undefined);
    const source = parseQueueSheet(sourceName, await workbook.getUsedRange(sourceTab));
    if (!source.shapeDetected) {
      throw new Error(
        `Could not find the header row on "${sourceTab}". Refusing to move rows off a ` +
          'sheet whose shape has not been read off the sheet itself.',
      );
    }

    const buckets = bucketBySegment(source);

    out();
    out(`family:        ${family}`);
    out(`source tab:    ${JSON.stringify(sourceTab)}`);
    out(`header row:    ${source.headerRow} (data from ${source.firstDataRow}, last ${source.lastRow})`);
    out(`patient rows:  ${source.rows.length} across ${source.groups.length} camps`);
    out();
    out('  segment    rows   camps');
    for (const segment of SEGMENTS) {
      const rows = buckets.get(segment)!;
      const camps = new Set(rows.map((row) => campKey(row.camp)));
      out(
        `  ${segment.padEnd(9)} ${String(rows.length).padStart(5)}   ${String(camps.size).padStart(5)}` +
          `   -> ${queueTabFor(family, segment)}`,
      );
    }
    out(`  ${'TOTAL'.padEnd(9)} ${String(source.rows.length).padStart(5)}`);

    const strays = buckets.get('Other')!;
    if (strays.length > 0) {
      out();
      out(`${strays.length} row(s) have no month. Their Date of Visit values:`);
      const seen = new Map<string, number>();
      for (const row of strays) {
        const raw = String(row.values['Date of Visit'] ?? '(blank)');
        seen.set(raw, (seen.get(raw) ?? 0) + 1);
      }
      for (const [raw, count] of [...seen.entries()].sort((a, b) => b[1] - a[1]).slice(0, 20)) {
        out(`  ${String(count).padStart(4)}  ${raw}`);
      }
    }

    // --- The fill audit ------------------------------------------------------
    out();
    out('Auditing the source tab\'s fills against the style rules...');
    const audit = await auditFills(workbook, sourceTab, source, config.readConcurrency);
    out(`  ${audit.probes} range(s) probed, ${audit.findings.length} unexpected`);

    if (audit.findings.length > 0) {
      out();
      out('  UNEXPECTED FILLS — these are colours the style rules do not predict,');
      out('  so regenerating the styling on the new tabs would NOT reproduce them:');
      for (const finding of audit.findings.slice(0, 40)) {
        out(
          `    ${finding.address.padEnd(12)} ${finding.what.padEnd(24)} ` +
            `expected ${finding.expected}, found ${finding.found}`,
        );
      }
      if (audit.findings.length > 40) {
        out(`    ... and ${audit.findings.length - 40} more`);
      }
    }

    const blocked = audit.findings.length > 0 && !ignoreFills;

    // --- Plan the move -------------------------------------------------------
    const moves: {
      segment: Segment;
      destination: QueueSheetName;
      tab: string;
      exists: boolean;
      intents: AppendQueueRowIntent[];
      skipped: number;
    }[] = [];

    // Every segment, including the ones with no rows today.
    //
    // A month tab that does not exist is a destination the cycle cannot reach:
    // it logs `queue.sheet_missing` and every row bound for it is skipped. That
    // is fine on the day of the migration, when June happens to be empty, and
    // wrong the moment a late June row turns up. The set of tabs is a property
    // of the split, not of what the data looks like this afternoon.
    for (const segment of SEGMENTS) {
      const rows = buckets.get(segment)!;
      const destination = queueTabFor(family, segment);
      const tab = QUEUE_SHEET_TABS[destination];
      const existingTab = resolveSheetName(tab, names);

      let intents = rows.map((row) => toIntent(row, destination));
      let skipped = 0;

      if (existingTab) {
        const parsed = parseQueueSheet(destination, await workbook.getUsedRange(existingTab));
        const { ids, identities } = alreadyThere(parsed);
        const before = intents.length;
        intents = intents.filter((intent) => {
          if (intent.syncId && ids.has(intent.syncId)) return false;
          const key = identityKey(intent.values);
          return !(key && identities.has(key));
        });
        skipped = before - intents.length;
      }

      moves.push({
        segment,
        destination,
        tab,
        exists: existingTab !== undefined,
        intents,
        skipped,
      });
    }

    out();
    out('  destination                       tab      to move   already there');
    for (const move of moves) {
      out(
        `  ${move.destination.padEnd(32)} ${(move.exists ? 'exists' : 'CREATE').padEnd(8)} ` +
          `${String(move.intents.length).padStart(7)}   ${String(move.skipped).padStart(13)}`,
      );
    }

    if (!apply) {
      out();
      if (blocked) {
        out('DRY RUN — and it would REFUSE to write: the fills above are not ones the');
        out('style rules can regenerate. Decide what they mean, then either clear them');
        out('or re-run with --ignore-unexpected-fills to move the rows anyway.');
      } else {
        out('DRY RUN — nothing written. Re-run with --apply.');
      }
      return;
    }

    if (blocked) {
      throw new Error(
        `Refusing to move rows off "${sourceTab}": ${audit.findings.length} cell range(s) ` +
          'carry a fill the style rules do not predict, so the new tabs would not ' +
          'reproduce them. Re-run with --ignore-unexpected-fills once you have decided ' +
          'what they mean.',
      );
    }

    // --- Move ----------------------------------------------------------------
    //
    // The column list is a property of the family, so it is the same on every one
    // of its monthly tabs and can be taken from the source's own name.
    const columns = queueColumnsFor(sourceName);
    const headerAddress = rangeAddress(
      LAYOUT.queue.firstColumn,
      LAYOUT.queue.headerRow,
      lastQueueColumnLetter(sourceName),
      LAYOUT.queue.headerRow,
    );

    for (const move of moves) {
      out();
      out(`--- ${move.destination} ---`);

      if (!move.exists) {
        await workbook.addWorksheet(move.tab);
        await workbook.writeRange(move.tab, headerAddress, [columns as unknown as unknown[]]);
        out(`created, header written to ${headerAddress}`);
      }

      if (move.intents.length === 0 && move.exists) {
        out('nothing to move.');
        continue;
      }

      if (move.intents.length === 0) {
        // A month with no rows yet. It still gets dressed, because an undressed
        // tab sitting beside styled ones is exactly the moment a queue looks
        // broken to somebody opening it — and because the rows will come.
        const bare = parseQueueSheet(move.destination, await workbook.getUsedRange(move.tab));
        const style = planQueueStyle({ sheet: bare });
        for (const operation of style.operations) {
          await applyStyle(workbook, move.tab, operation);
        }
        out(`no rows for this month yet; styled with ${style.operations.length} operations`);
        continue;
      }

      const parsed = parseQueueSheet(move.destination, await workbook.getUsedRange(move.tab));
      const plan = planQueueAppend({
        sheet: parsed,
        appends: move.intents,
        // The office ordered these camps. Alphabetising them would reorder the
        // whole tab on a job whose premise is that nothing changes but the month.
        newBlockOrder: 'as-given',
      });

      out(
        `${plan.appended} rows into ${plan.placements.length} camp blocks, ` +
          `${plan.rowsAdded} rows written, ${plan.shadedCells} cells shaded`,
      );

      for (const operation of plan.operations) {
        switch (operation.kind) {
          case 'insert-rows':
            await workbook.insertRows(move.tab, operation.row, operation.count);
            break;
          case 'write-cells':
            await workbook.writeRange(move.tab, operation.address, operation.values);
            break;
          case 'shade':
            await workbook.setFill(move.tab, operation.address, operation.color);
            break;
        }
      }

      // --- Verify before anything is retired --------------------------------
      const after = parseQueueSheet(move.destination, await workbook.getUsedRange(move.tab));
      const landed = alreadyThere(after);

      // A row with neither a SyncID nor enough of a name to key on cannot be
      // looked for by identity — that is what "too little identity" means, and
      // it is the same reason adoption could not link it. Counted and reported
      // rather than quietly passed, so the verification never claims more than
      // it checked.
      let unverifiable = 0;
      const missing = move.intents.filter((intent) => {
        if (intent.syncId) return !landed.ids.has(intent.syncId);
        const key = identityKey(intent.values);
        if (!key) {
          unverifiable++;
          return false;
        }
        return !landed.identities.has(key);
      });

      if (missing.length > 0) {
        throw new Error(
          `${missing.length} of ${move.intents.length} rows did not land on "${move.tab}". ` +
            'The source tab has NOT been touched — re-run to finish the move.',
        );
      }
      out(
        `verified: ${after.rows.length} rows on the tab, ` +
          `${move.intents.length - unverifiable} of ${move.intents.length} confirmed by ID or identity`,
      );
      if (unverifiable > 0) {
        out(
          `  ${unverifiable} row(s) carry neither a SyncID nor a full name and date of birth, ` +
            'so they were written but could not be looked up afterwards. The row count above ' +
            'is the check for those.',
        );
      }

      // --- Dress it ----------------------------------------------------------
      const style = planQueueStyle({ sheet: after });
      for (const operation of style.operations) {
        try {
          await applyStyle(workbook, move.tab, operation);
        } catch (error) {
          // Graph echoes a failed range payload back, and the logger drops error
          // bodies for exactly that reason — so a bare 400 says nothing about
          // which of seventy-odd calls failed. The operation's description is
          // safe to print and is the only thing that makes this diagnosable.
          out(`FAILED on style op: ${operation.kind} ${operation.address} — ${operation.what}`);
          throw error;
        }
      }
      out(`styled: ${style.operations.length} operations`);

      log.info('split.moved', {
        sheet: move.tab,
        destination: move.destination,
        count: move.intents.length,
        counts: { styled: style.operations.length, shaded: plan.shadedCells },
      });
    }

    // --- Retire the source ---------------------------------------------------
    out();
    if (!archive) {
      out(`Done. "${sourceTab}" is untouched and still holds all ${source.rows.length} rows.`);
      out('It is already out of QUEUE_SHEET_NAMES, so the sync neither reads nor writes');
      out('it — but staff can still see it. Once you have checked the new tabs:');
      out(`  npm run split -- ${JSON.stringify(family)} --apply --archive`);
      return;
    }

    const archived = `${family} (old)`;
    if (normalizeSheetName(sourceTab) === normalizeSheetName(archived)) {
      out(`"${sourceTab}" is already the archived copy. Nothing to rename.`);
    } else {
      await workbook.updateWorksheet(sourceTab, { name: archived, visibility: 'Hidden' });
      log.info('split.archived', { sheet: sourceTab, renamedTo: archived });
      out(`Renamed ${JSON.stringify(sourceTab)} to ${JSON.stringify(archived)} and hid it.`);
      out('Every row it held is still on it, exactly as it stood before the split.');
    }

    out();
    out('TWO SETTINGS PER NEW TAB, BY HAND. Graph refuses both against this workbook:');
    out('  Data > Data Validation > Allow: List > Source: Done   (on the Resolved column)');
    out('  View > Freeze Panes > Freeze Top Row');
  } catch (error) {
    log.error('split.failed', describeError(error));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'split.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
