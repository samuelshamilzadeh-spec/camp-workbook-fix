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
 * **Fills are COPIED, cell by cell, never re-derived.** The first version of this
 * script regenerated them from `REQUIRED_FIELDS` and `planQueueStyle`, on the
 * reasoning that every fill on these tabs came from those rules in the first
 * place. Auditing the live tab disproved it: 570 cells carry a colour the rules
 * know nothing about — amber in `Date of Visit`, orange in `Medications`, purple
 * in `Allergies`, green across the three insurance columns — which is staff
 * colour-coding built up by hand. Regenerating would have erased all of it.
 *
 * So the source tab's fills are read cell by cell and replayed onto the new tab
 * at the row each patient landed on. Nothing is added: a blank required field
 * that is not shaded today does not become shaded, because what is on the sheet
 * now is what goes. `planQueueStyle` still supplies the fonts, borders, column
 * widths, row heights and number formats, which it reproduces exactly — its fill
 * operations are dropped, since the copied fills are the truth.
 *
 * Excel's own copy would be better than any of this, and Graph refuses it: both
 * `worksheets/copy` and `range/copyFrom` answer "Resource not found for the
 * segment", the same way `dataValidation` and `freezePanes` do on this workbook.
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
  BLANK_REQUIRED_FILL,
  LAYOUT,
  QUEUE_SHEET_TABS,
  SEGMENTS,
  SPLIT_FAMILIES,
  STYLE,
  assertLayoutVerified,
  loadConfig,
  queueTabFor,
  type QueueSheetName,
  type Segment,
  type SplitFamily,
} from '../src/config';
import { planQueueAppend } from '../src/domain/append';
import {
  columnToIndex,
  indexToColumn,
  isBlank,
  parseAddress,
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

// --- Capturing the fills ----------------------------------------------------

/**
 * Every filled cell on a tab, keyed `row!COLUMN`.
 *
 * Colours are read rather than derived because the live tabs carry a colour
 * scheme this codebase did not put there and cannot predict. See the header
 * comment: 570 cells on `Missing Info` alone.
 */
export type FillMap = Map<string, string>;

const DEFAULT = 'default';

/**
 * What counts as "no colour here", per property.
 *
 * An unfilled cell reads back as white and nothing ever deliberately fills white,
 * so for a FILL those are the same thing. For a FONT the default is black, and
 * an explicit black is likewise indistinguishable from — and equivalent to — no
 * colour at all.
 */
function normalizeColor(color: string | undefined, kind: 'fill' | 'font'): string {
  if (!color) return DEFAULT;
  const upper = color.trim().toUpperCase();
  const neutral = kind === 'fill' ? ['#FFFFFF', 'FFFFFF'] : ['#000000', '000000'];
  return neutral.includes(upper) ? DEFAULT : upper;
}

const fillKey = (row: number, column: string): string => `${row}!${column}`;

/**
 * Reads the fill of every cell in the tab's data columns.
 *
 * Graph has no call that returns a grid of formats — `cellProperties` is not
 * exposed, and `copyFrom` is refused — so this is one request per cell in the
 * worst case. It is kept to roughly one per ROW in the common case by asking
 * about the whole row first: Graph answers a multi-cell range with a colour only
 * when every cell agrees, so a uniformly unfilled row costs a single call and
 * only a row that actually varies is opened up.
 *
 * A null answer means "they disagree" for a range and "no fill" for a single
 * cell, which is why the second pass is per cell, where the reading is
 * unambiguous.
 */
async function readColors(
  workbook: Workbook,
  tab: string,
  rows: readonly number[],
  firstColumn: string,
  lastColumn: string,
  concurrency: number,
  kind: 'fill' | 'font',
): Promise<{ colors: FillMap; calls: number }> {
  const colors: FillMap = new Map();
  const from = columnToIndex(firstColumn);
  const to = columnToIndex(lastColumn);
  const read = (address: string) =>
    kind === 'fill' ? workbook.getFill(tab, address) : workbook.getFontColor(tab, address);

  const wholeRows = await mapWithConcurrency(rows, concurrency, (row) =>
    read(rangeAddress(firstColumn, row, lastColumn, row)),
  );

  const mixed: number[] = [];
  rows.forEach((row, index) => {
    if (wholeRows[index]!.color === undefined) {
      // Either the whole row is default or its cells disagree. Both are possible
      // and indistinguishable at range width, so look closer.
      mixed.push(row);
      return;
    }
    const color = normalizeColor(wholeRows[index]!.color, kind);
    if (color === DEFAULT) return;
    for (let index2 = from; index2 <= to; index2++) {
      colors.set(fillKey(row, indexToColumn(index2)), color);
    }
  });

  const cells: { row: number; column: string }[] = [];
  for (const row of mixed) {
    for (let index = from; index <= to; index++) {
      cells.push({ row, column: indexToColumn(index) });
    }
  }

  const perCell = await mapWithConcurrency(cells, concurrency, (cell) =>
    read(rangeAddress(cell.column, cell.row, cell.column, cell.row)),
  );
  cells.forEach((cell, index) => {
    const color = normalizeColor(perCell[index]!.color, kind);
    if (color !== DEFAULT) colors.set(fillKey(cell.row, cell.column), color);
  });

  return { colors, calls: rows.length + cells.length };
}

/** What the fills look like, for the dry-run report. Colours only, never values. */
function summarizeFills(fills: FillMap): { color: string; count: number }[] {
  const counts = new Map<string, number>();
  for (const color of fills.values()) counts.set(color, (counts.get(color) ?? 0) + 1);
  return [...counts.entries()]
    .map(([color, count]) => ({ color, count }))
    .sort((a, b) => b.count - a.count);
}

/**
 * Groups a row's filled cells into contiguous same-colour runs, so three green
 * insurance columns are one `setFill` rather than three.
 */
function fillRuns(
  fills: FillMap,
  sourceRow: number,
  destinationRow: number,
  firstColumn: string,
  lastColumn: string,
): { address: string; color: string }[] {
  const from = columnToIndex(firstColumn);
  const to = columnToIndex(lastColumn);
  const runs: { address: string; color: string }[] = [];

  let start: number | undefined;
  let color: string | undefined;

  const close = (end: number) => {
    if (start === undefined || color === undefined) return;
    runs.push({
      address: rangeAddress(indexToColumn(start), destinationRow, indexToColumn(end), destinationRow),
      color,
    });
    start = undefined;
    color = undefined;
  };

  for (let index = from; index <= to; index++) {
    const here = fills.get(fillKey(sourceRow, indexToColumn(index)));
    if (here !== color) {
      close(index - 1);
      if (here !== undefined) {
        start = index;
        color = here;
      }
    }
  }
  close(to);

  return runs;
}

// --- How wide the tab really is ---------------------------------------------

/**
 * The SyncID column and everything right of it. Never part of the copied width:
 * BA is written separately by the appender, and past it lies whatever debris a
 * script has left — `Missing Info` still carries a `sync-probe` header out at CA
 * from a concurrency run in July.
 */
const isBeyondData = (column: string, layout = LAYOUT): boolean =>
  columnToIndex(column) >= columnToIndex(layout.queue.syncIdColumn);

/**
 * Columns to the right of the schema that the sheet actually uses.
 *
 * `QUEUE_COLUMNS` describes A..R and these tabs are wider than that. Anything
 * out there holding a HEADER is carried across; anything holding DATA stops the
 * run, because `planQueueAppend` writes the schema's columns and would drop it.
 */
function extraColumns(
  sheet: ParsedQueueSheet,
  used: { address: string; values: unknown[][] },
): { lastColumn: string | undefined; labels: string[]; withData: string[] } {
  const { startRow, startColumn } = parseAddress(used.address);
  const schemaLast = columnToIndex(lastQueueColumnLetter(sheet.sheet));
  const width = used.values.reduce((max, row) => Math.max(max, row.length), 0);
  const offset = columnToIndex(startColumn);

  const labels: string[] = [];
  const withData: string[] = [];
  let last: number | undefined;

  for (let index = 0; index < width; index++) {
    const column = indexToColumn(offset + index);
    if (offset + index <= schemaLast || isBeyondData(column)) continue;

    const header = used.values[sheet.headerRow - startRow]?.[index];
    const label = String(header ?? '').trim();

    let filled = false;
    for (const row of sheet.rows) {
      const value = used.values[row.row - startRow]?.[index];
      if (!isBlank(value)) {
        filled = true;
        break;
      }
    }

    if (filled) withData.push(`${column} (${label || 'unlabelled'})`);
    if (label || filled) {
      labels.push(`${column} "${label}"`);
      last = offset + index;
    }
  }

  return {
    lastColumn: last === undefined ? undefined : indexToColumn(last),
    labels,
    withData,
  };
}

/** The source tab's own header row, so labels the schema lacks travel with it. */
function headerValues(
  sheet: ParsedQueueSheet,
  used: { address: string; values: unknown[][] },
  lastColumn: string,
): unknown[] {
  const { startRow, startColumn } = parseAddress(used.address);
  const from = columnToIndex(LAYOUT.queue.firstColumn);
  const to = columnToIndex(lastColumn);
  const row = used.values[sheet.headerRow - startRow] ?? [];
  const offset = columnToIndex(startColumn);

  const values: unknown[] = [];
  for (let index = from; index <= to; index++) {
    const cell = row[index - offset];
    values.push(cell === undefined || cell === '' ? null : cell);
  }
  return values;
}

// --- Moving the rows --------------------------------------------------------

/**
 * A source row as something the appender can write.
 *
 * `blankRequired` is deliberately EMPTY, which turns off the appender's own red
 * shading. This is a migration, not a new row: whatever colour the cell carries
 * today is copied across afterwards, and a blank required field that nobody
 * shaded stays unshaded. Letting the appender add its red here would paint 84
 * cells on `Missing Info` alone that the office never had.
 *
 * A row with no SyncID keeps none: `''` writes an empty cell, so it lands on the
 * new tab exactly as unstamped as it left, and shows up in the same orphan
 * report it always did. Inventing an ID here would silently link it to a daily
 * row nothing had matched it to.
 */
function toIntent(row: QueueRow, destination: QueueSheetName): AppendQueueRowIntent {
  return {
    kind: 'append-queue-row',
    destination,
    syncId: row.syncId ?? '',
    sourceSheet: String(row.values['Date of Visit'] ?? ''),
    sourceRow: Number(row.values['Source Row'] ?? 0),
    camp: row.camp,
    values: row.values,
    blankRequired: [],
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

  if (!wanted) {
    throw new Error('Usage: npm run split -- "Missing Info" [--apply] [--archive]');
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
    const sourceUsed = await workbook.getUsedRange(sourceTab);
    const source = parseQueueSheet(sourceName, sourceUsed);
    if (!source.shapeDetected) {
      throw new Error(
        `Could not find the header row on "${sourceTab}". Refusing to move rows off a ` +
          'sheet whose shape has not been read off the sheet itself.',
      );
    }

    // How far right to copy.
    //
    // The schema stops at R, and these tabs do not: `Ineligible & Inactive`
    // carries `Updated Insurance Carrier`, `Updated Insurance ID #` and
    // `Updated Medicaid #` at S, T and U. They hold no data today, and copying
    // only A..R would drop them the day somebody uses one — so the width is
    // measured off the sheet instead of taken from the config.
    const extra = extraColumns(source, sourceUsed);
    const copyLastColumn =
      extra.lastColumn ?? lastQueueColumnLetter(sourceName);

    if (extra.withData.length > 0) {
      // Values for these are not carried: `planQueueAppend` writes the schema's
      // columns and nothing else. Rather than drop them quietly, stop.
      throw new Error(
        `"${sourceTab}" has data in ${extra.withData.join(', ')}, which are past the ` +
          `columns this migration knows how to move (A..${lastQueueColumnLetter(sourceName)}). ` +
          'Add them to QUEUE_COLUMNS, or clear them, before splitting this tab.',
      );
    }
    if (extra.labels.length > 0) {
      out();
      out(`  columns past the schema, carried as headers only (all empty): ${extra.labels.join(', ')}`);
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

    // --- Capture every colour on the tab -------------------------------------
    const captureRows = [
      source.headerRow,
      ...source.groups.map((group) => group.headerRow),
      ...source.rows.map((row) => row.row),
      ...(source.totalRow === undefined ? [] : [source.totalRow]),
    ].sort((a, b) => a - b);

    const describe = (color: string): string =>
      color === STYLE.headerFill.toUpperCase()
        ? '  (header bar)'
        : color === STYLE.dividerFill.toUpperCase()
          ? '  (camp divider band)'
          : color === STYLE.totalFill.toUpperCase()
            ? '  (total row)'
            : color === BLANK_REQUIRED_FILL.toUpperCase()
              ? '  (blank required, the office red)'
              : '';

    out();
    out('Reading every colour on the source tab (Graph has no bulk format read)...');

    const fills = await readColors(
      workbook,
      sourceTab,
      captureRows,
      LAYOUT.queue.firstColumn,
      copyLastColumn,
      config.readConcurrency,
      'fill',
    );
    out(`  fills: ${fills.calls} request(s), ${fills.colors.size} coloured cells`);
    for (const { color, count } of summarizeFills(fills.colors)) {
      out(`    ${String(count).padStart(5)}  ${color}${describe(color)}`);
    }

    // Text colour matters here too. `Medicaid #` reading `inactive` in red is a
    // staff member saying something, and `planQueueStyle` paints the whole body
    // black — so the non-black ones are captured and put back after it runs.
    const fonts = await readColors(
      workbook,
      sourceTab,
      captureRows,
      LAYOUT.queue.firstColumn,
      copyLastColumn,
      config.readConcurrency,
      'font',
    );
    out(`  font colours: ${fonts.calls} request(s), ${fonts.colors.size} non-black cells`);
    for (const { color, count } of summarizeFills(fonts.colors)) {
      out(`    ${String(count).padStart(5)}  ${color}${color === STYLE.headerFont.toUpperCase() ? '  (header text)' : ''}`);
    }

    out('  Every one of these is copied to the new tabs as-is. Nothing is added,');
    out('  nothing is re-derived.');


    // --- Plan the move -------------------------------------------------------
    const moves: {
      segment: Segment;
      destination: QueueSheetName;
      tab: string;
      exists: boolean;
      rows: QueueRow[];
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

      let keep = rows;
      let skipped = 0;

      if (existingTab) {
        const parsed = parseQueueSheet(destination, await workbook.getUsedRange(existingTab));
        const { ids, identities } = alreadyThere(parsed);
        keep = rows.filter((row) => {
          if (row.syncId && ids.has(row.syncId)) return false;
          const key = identityKey(row.values);
          return !(key && identities.has(key));
        });
        skipped = rows.length - keep.length;
      }

      moves.push({
        segment,
        destination,
        tab,
        exists: existingTab !== undefined,
        rows: keep,
        intents: keep.map((row) => toIntent(row, destination)),
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
      out('DRY RUN — nothing written. Re-run with --apply.');
      return;
    }

    // --- Move ----------------------------------------------------------------
    //
    // The header row is copied from the SOURCE rather than written from the
    // configured column list, so labels the schema does not know about —
    // `Ineligible & Inactive` carries `Updated Insurance Carrier` and two more at
    // S, T and U — arrive on the new tab intact.
    const headerAddress = rangeAddress(
      LAYOUT.queue.firstColumn,
      LAYOUT.queue.headerRow,
      copyLastColumn,
      LAYOUT.queue.headerRow,
    );
    const sourceHeader = headerValues(source, sourceUsed, copyLastColumn);

    for (const move of moves) {
      out();
      out(`--- ${move.destination} ---`);

      if (!move.exists) {
        await workbook.addWorksheet(move.tab);
        await workbook.writeRange(move.tab, headerAddress, [sourceHeader]);
        out(`created, header copied to ${headerAddress}`);
      }

      // The header row's own colours, so the new tab's bar matches the old one.
      for (const run of fillRuns(fills.colors, source.headerRow, LAYOUT.queue.headerRow, LAYOUT.queue.firstColumn, copyLastColumn)) {
        await workbook.setFill(move.tab, run.address, run.color);
      }
      for (const run of fillRuns(fonts.colors, source.headerRow, LAYOUT.queue.headerRow, LAYOUT.queue.firstColumn, copyLastColumn)) {
        await workbook.setFont(move.tab, run.address, { color: run.color });
      }

      const dressWithoutFills = async (sheet: ParsedQueueSheet): Promise<number> => {
        // Fonts, borders, widths, row heights and number formats are regenerated
        // — `planQueueStyle` reproduces those exactly, because migrate applied
        // them from these same rules. Its FILL operations are dropped: the
        // copied fills are the truth, and re-deriving them is what would have
        // erased the office's colour coding.
        const operations = planQueueStyle({ sheet }).operations.filter(
          (operation) => operation.kind !== 'fill',
        );
        for (const operation of operations) {
          try {
            await applyStyle(workbook, move.tab, operation);
          } catch (error) {
            // Graph echoes a failed range payload back, and the logger drops
            // error bodies for exactly that reason — so a bare 400 says nothing
            // about which of seventy-odd calls failed. The operation's own
            // description is safe to print and is what makes this diagnosable.
            out(`FAILED on style op: ${operation.kind} ${operation.address} — ${operation.what}`);
            throw error;
          }
        }
        return operations.length;
      };

      if (move.intents.length === 0) {
        if (move.exists) {
          out('nothing to move.');
          continue;
        }
        // A month with no rows yet. It still gets dressed, because an undressed
        // tab beside styled ones is the moment a queue looks broken to somebody
        // opening it — and because the rows will come.
        const bare = parseQueueSheet(move.destination, await workbook.getUsedRange(move.tab));
        out(`no rows for this month yet; styled with ${await dressWithoutFills(bare)} operations`);
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

      out(`${plan.appended} rows into ${plan.placements.length} camp blocks`);

      for (const operation of plan.operations) {
        switch (operation.kind) {
          case 'insert-rows':
            await workbook.insertRows(move.tab, operation.row, operation.count);
            break;
          case 'write-cells':
            await workbook.writeRange(move.tab, operation.address, operation.values);
            break;
          case 'shade':
            // Never reached: `toIntent` sends no blankRequired, so the appender
            // plans no shading. Fills come from the source, not from a rule.
            await workbook.setFill(move.tab, operation.address, operation.color);
            break;
        }
      }

      // --- Verify, and learn where every row landed --------------------------
      const after = parseQueueSheet(move.destination, await workbook.getUsedRange(move.tab));

      if (after.rows.length !== move.rows.length) {
        throw new Error(
          `"${move.tab}" holds ${after.rows.length} rows but ${move.rows.length} were moved. ` +
            'The source tab has NOT been touched — look at the tab before re-running.',
        );
      }

      // Positional, then PROVEN by SyncID rather than assumed. The appender
      // preserves the order it was given and `parseQueueSheet` returns rows in
      // sheet order, so row i out is row i in — but the fill replay depends on
      // that mapping being exactly right, and a silently wrong one would paint
      // one patient's colours onto another.
      let proven = 0;
      const rowMap = new Map<number, number>();
      move.rows.forEach((sourceRow, index) => {
        const destinationRow = after.rows[index]!;
        if (sourceRow.syncId || destinationRow.syncId) {
          if (sourceRow.syncId !== destinationRow.syncId) {
            throw new Error(
              `Row ${index + 1} on "${move.tab}" is not the row that was written ` +
                '(SyncID mismatch). Refusing to copy fills onto the wrong patient.',
            );
          }
          proven++;
        }
        rowMap.set(sourceRow.row, destinationRow.row);
      });
      out(`verified: ${after.rows.length} rows, ${proven} confirmed by SyncID`);

      // Camp dividers, matched by name so the band and any colour on it follow.
      const destinationGroups = new Map(after.groups.map((group) => [campKey(group.camp), group]));
      for (const group of source.groups) {
        const match = destinationGroups.get(campKey(group.camp));
        if (match) rowMap.set(group.headerRow, match.headerRow);
      }
      if (source.totalRow !== undefined && after.totalRow !== undefined) {
        rowMap.set(source.totalRow, after.totalRow);
      }

      // --- Dress, THEN replay the colours ------------------------------------
      //
      // Order matters. `planQueueStyle` sets the whole body to black Arial 10,
      // so a captured red `inactive` has to go on after it or the styling pass
      // would wipe the very thing this is here to preserve.
      const styled = await dressWithoutFills(after);
      out(`styled: ${styled} operations (fonts, borders, widths, number formats)`);

      const runsFor = (map: FillMap): { address: string; color: string }[] => {
        const runs: { address: string; color: string }[] = [];
        for (const [sourceRow, destinationRow] of rowMap) {
          runs.push(
            ...fillRuns(map, sourceRow, destinationRow, LAYOUT.queue.firstColumn, copyLastColumn),
          );
        }
        return runs;
      };

      const fillWork = runsFor(fills.colors);
      const fontWork = runsFor(fonts.colors);
      let painted = 0;
      const total = fillWork.length + fontWork.length;

      for (const run of fillWork) {
        await workbook.setFill(move.tab, run.address, run.color);
        if (++painted % 100 === 0) out(`  ${painted}/${total} colour ranges copied`);
      }
      for (const run of fontWork) {
        await workbook.setFont(move.tab, run.address, { color: run.color });
        if (++painted % 100 === 0) out(`  ${painted}/${total} colour ranges copied`);
      }
      out(
        `colours copied: ${fillWork.length} fill range(s) and ${fontWork.length} font range(s), ` +
          'exactly as they were',
      );

      log.info('split.moved', {
        sheet: move.tab,
        destination: move.destination,
        count: move.rows.length,
        counts: { fills: fillWork.length, fonts: fontWork.length, styled },
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

    // Where the tab sits in the bar, read BEFORE anything moves. The monthly
    // tabs take this spot: staff have reached for `Missing Info` in the same
    // place all season, and leaving its replacements stranded at the end of a
    // 61-sheet workbook is its own kind of broken.
    const sheets = await workbook.listWorksheets();
    const anchor = sheets.find((sheet) => sheet.name === sourceTab)?.position;

    const archived = `${family} (old)`;
    if (normalizeSheetName(sourceTab) === normalizeSheetName(archived)) {
      out(`"${sourceTab}" is already the archived copy. Nothing to rename.`);
    } else {
      await workbook.updateWorksheet(sourceTab, { name: archived, visibility: 'Hidden' });
      log.info('split.archived', { sheet: sourceTab, renamedTo: archived });
      out(`Renamed ${JSON.stringify(sourceTab)} to ${JSON.stringify(archived)} and hid it.`);
      out('Every row it held is still on it, exactly as it stood before the split.');
    }

    if (anchor !== undefined) {
      // Left to right with consecutive targets: setting a position removes the
      // sheet and re-inserts it there, so each tab lands exactly at its target
      // and the ones after it shuffle around it.
      // LAST tab first, every one of them to the SAME index.
      //
      // Inserting at an index pushes whatever is there to the right, so placing
      // Other, then August, then July, then June — all at `anchor` — leaves them
      // in June, July, August, Other order without this code needing a theory of
      // what Graph does to the indices in between. Two attempts at being clever
      // about that failed differently: consecutive targets 3,4,5,6 landed at
      // 4,6,8,10, and chaining off where each tab actually settled interleaved
      // them with `Verify Insurance` and two daily sheets.
      const byName = new Map(sheets.map((sheet) => [sheet.name, sheet]));
      for (const segment of [...SEGMENTS].reverse()) {
        const sheet = byName.get(QUEUE_SHEET_TABS[queueTabFor(family, segment)]);
        if (sheet) await workbook.moveWorksheet(sheet.id, anchor);
      }

      // Read the order back rather than trust the 200s. The name-addressed form
      // of this PATCH returns success and moves nothing, which is exactly the
      // kind of failure that gets reported as a job well done.
      //
      // Retried, because the position a read reports right after a move can lag
      // the truth on this workbook — the first version of this check called four
      // correctly-placed tabs a failure.
      let positions: number[] = [];
      let contiguous = false;
      for (let attempt = 0; attempt < 3 && !contiguous; attempt++) {
        const after = await workbook.listWorksheets();
        positions = SEGMENTS.map(
          (segment) =>
            after.find((sheet) => sheet.name === QUEUE_SHEET_TABS[queueTabFor(family, segment)])
              ?.position,
        ).filter((position): position is number => position !== undefined);
        contiguous =
          positions.length > 0 &&
          positions.every((value, index) => value === positions[0]! + index);
      }
      out(
        contiguous
          ? `Moved the ${positions.length} monthly tabs to position ${positions[0]}, where ${JSON.stringify(family)} sat.`
          : `WARNING: the tabs did not move as asked — they sit at ${positions.join(', ')}. ` +
            'Drag them into place in Excel.',
      );
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
