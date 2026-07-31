import {
  BLANK_REQUIRED_FILL,
  LAYOUT,
  STYLE,
  queueColumnsFor,
  type QueueColumn,
  type QueueSheetName,
  type WorkbookLayout,
} from '../config';
import { toDateCellValue } from './dates';
import { offsetColumn, rangeAddress } from './cells';
import { campKey } from './dailySheets';
import {
  formatGrandTotal,
  formatGroupHeader,
  lastQueueColumnLetter,
  type CampGroup,
  type ParsedQueueSheet,
} from './queueSheets';
import type { AppendQueueRowIntent } from './reconcile';

/**
 * Phase 2b — turning `append-queue-row` intents into an ordered list of
 * workbook operations.
 *
 * This is the pure half: no Graph calls, no side effects, the same output every
 * time for a given sheet state and intent list. `scripts/append-queue.ts`
 * executes what comes out of here, which means the whole plan can be printed
 * and read by a human before anything is written.
 *
 * Three constraints shape the design.
 *
 * **Rows move as whole rows.** The office's hard rule is that one row is one
 * patient. So a row is never squeezed into a cell range that would push some
 * columns and not others, and a new row's columns go in a single range write so
 * a human never sees a half-populated patient.
 *
 * **Row numbers must not go stale.** Inserting rows shifts everything below,
 * invalidating any row number already computed for the rows underneath. Rather
 * than recomputing offsets — which is where this kind of code goes wrong — the
 * plan is emitted BOTTOM-UP. Every address is in the coordinates of the sheet as
 * it was read, and because each operation only touches rows at or below the ones
 * still to be processed, an insert higher up carries already-written content
 * down without invalidating anything still pending.
 *
 * **Appending nothing must plan nothing.** A second run over the same workbook
 * state produces no operations, because the reconciler only emits an append for
 * a source row that has no queue row — which is why the applier stamps the daily
 * sheet's SyncID before it appends, never after.
 */

// --- Operations -------------------------------------------------------------

export interface InsertRowsOperation {
  kind: 'insert-rows';
  /** First row of the gap to open. Everything from here down shifts. */
  row: number;
  count: number;
  /** Why this insert exists, for the dry-run report. */
  reason: string;
}

export interface WriteCellsOperation {
  kind: 'write-cells';
  address: string;
  values: unknown[][];
  /** First row covered, for the dry-run report. */
  row: number;
  count: number;
  purpose: 'rows' | 'sync-ids' | 'divider' | 'total';
}

export interface ShadeOperation {
  kind: 'shade';
  address: string;
  color: string;
  /** Which required fields are blank here. Column names only, never values. */
  columns: QueueColumn[];
}

export type QueueOperation = InsertRowsOperation | WriteCellsOperation | ShadeOperation;

/** What a plan does to one camp block, for the dry-run report. */
export interface CampPlacement {
  camp: string;
  campKey: string;
  /** True when this camp had no block on the sheet and one is being created. */
  isNew: boolean;
  added: number;
  /** Rows this camp held before the append. */
  existing: number;
  /** Where the new rows go, in the coordinates of the sheet as read. */
  atRow: number;
}

export interface QueueAppendPlan {
  destination: QueueSheetName;
  operations: QueueOperation[];
  placements: CampPlacement[];
  /** Patient rows appended. Divider and TOTAL rows are not counted. */
  appended: number;
  /** Every row added, dividers included. This is what shifts the sheet. */
  rowsAdded: number;
  /** Patient rows the sheet holds once the plan is applied. */
  totalAfter: number;
  /** True when the plan opens gaps in the middle of the sheet. */
  requiresInserts: boolean;
  /** Patient cells that get dark red shading because a required field is blank. */
  shadedCells: number;
}

export interface PlanQueueAppendInput {
  sheet: ParsedQueueSheet;
  appends: readonly AppendQueueRowIntent[];
  layout?: WorkbookLayout;
}

/**
 * Patients whose camp cell is blank still need somewhere to sit. They get their
 * own block rather than inflating a real camp's count.
 *
 * The label is exactly what `campKey` produces for a missing camp, so the block
 * this run writes is the block the next run finds. A prettier label would not
 * match its own key and every run would create another block below the last.
 */
export const NO_CAMP_LABEL = '(no camp)';

/**
 * A camp block about to be written or grown. `atRow` is where the new rows go
 * in the sheet's original coordinates: for an existing camp, straight after its
 * last row; for a new camp, at the end of the body.
 */
interface Block {
  campKey: string;
  camp: string;
  isNew: boolean;
  atRow: number;
  existing: number;
  intents: AppendQueueRowIntent[];
  /** Header row to rewrite with the new count. Only set for an existing camp. */
  headerRow: number | undefined;
}

export function planQueueAppend(input: PlanQueueAppendInput): QueueAppendPlan {
  const layout = input.layout ?? LAYOUT;
  const { sheet } = input;
  const appends = input.appends.filter((intent) => intent.destination === sheet.sheet);

  const firstColumn = layout.queue.firstColumn;
  const columns = queueColumnsFor(sheet.sheet);
  const existingTotal = sheet.rows.length;

  if (appends.length === 0) {
    return {
      destination: sheet.sheet,
      operations: [],
      placements: [],
      appended: 0,
      rowsAdded: 0,
      totalAfter: existingTotal,
      requiresInserts: false,
      shadedCells: 0,
    };
  }

  // --- Work out where each camp's rows go ------------------------------------

  const groupsByKey = new Map<string, CampGroup>();
  for (const group of sheet.groups) {
    // First wins. Two blocks for one camp is a pre-existing mess on the sheet,
    // and growing the first keeps this run's patients together rather than
    // splitting them across both.
    if (!groupsByKey.has(campKey(group.camp))) groupsByKey.set(campKey(group.camp), group);
  }

  const byCamp = new Map<string, AppendQueueRowIntent[]>();
  for (const intent of appends) {
    const key = campKey(intent.camp);
    const list = byCamp.get(key) ?? [];
    list.push(intent);
    byCamp.set(key, list);
  }

  // Where a brand new camp block goes.
  //
  // This has to sit at or below every EXISTING camp block, because new blocks
  // are emitted first (they end up bottom-most) and the whole bottom-up
  // ordering argument depends on nothing above them being emitted later. Taking
  // it from the TOTAL row alone was not enough: a tab whose TOTAL sits above a
  // camp block — which a staff member can produce by adding a block underneath
  // it — put the new block ABOVE a grown one, and its insert then shifted that
  // grown block down while the planner still held the old addresses for it.
  //
  // So it is measured from the content: below every patient row and every camp
  // header, and above the TOTAL when the TOTAL is where it belongs.
  const floor = Math.max(sheet.lastRow, sheet.firstDataRow - 1);
  const contentEnd = Math.max(
    sheet.firstDataRow - 1,
    ...sheet.rows.map((row) => row.row),
    ...sheet.groups.map((group) => group.headerRow),
  );
  const appendPoint =
    sheet.totalRow !== undefined && sheet.totalRow > contentEnd
      ? sheet.totalRow
      : contentEnd + 1;

  const blocks: Block[] = [];
  for (const [key, intents] of byCamp) {
    const group = groupsByKey.get(key);
    if (group) {
      const lastRowOfGroup = group.rows.reduce(
        (max, row) => Math.max(max, row.row),
        group.headerRow,
      );
      blocks.push({
        campKey: key,
        // Keep whatever spelling is already on the sheet. Staff put it there,
        // and the inspection found `elky`/`Elky` and `bnos naale`/`Bnos Naale`
        // both in use — renaming a divider is not this job's business.
        camp: group.camp,
        isNew: false,
        atRow: lastRowOfGroup + 1,
        existing: group.rows.length,
        intents,
        headerRow: group.headerRow,
      });
    } else {
      blocks.push({
        campKey: key,
        camp: displayCamp(intents),
        isNew: true,
        atRow: appendPoint,
        existing: 0,
        intents,
        headerRow: undefined,
      });
    }
  }

  // New camps all stack at one point, so their order among themselves is this
  // plan's to choose: alphabetical, which is stable across runs and needs no
  // knowledge of how the office ordered the tab. Camps already on the sheet keep
  // their place and are never reordered.
  const newBlocks = blocks
    .filter((block) => block.isNew)
    .sort((a, b) => a.campKey.localeCompare(b.campKey));
  // Deepest first, so each insert only moves rows that have already been dealt
  // with.
  const grownBlocks = blocks
    .filter((block) => !block.isNew)
    .sort((a, b) => b.atRow - a.atRow);

  // --- Emit operations, bottom of the sheet first ---------------------------

  const operations: QueueOperation[] = [];
  const totalAfter = existingTotal + appends.length;
  let shadedCells = 0;
  let rowsAdded = 0;

  // How far down the sheet is occupied *right now*, as the plan is built. It
  // starts at the sheet's last used row and grows with every block emitted. It
  // is the whole basis for deciding whether a write needs a gap opened first,
  // and it is why the last existing camp and a brand new camp can both want row
  // 21 without one silently overwriting the other.
  let occupiedTo = floor;

  const emit = (args: {
    firstRow: number;
    grid: unknown[][];
    syncIds: unknown[][];
    shades: { offset: number; columns: QueueColumn[] }[];
    reason: string;
  }): void => {
    const { firstRow, grid, syncIds, shades } = args;
    if (grid.length === 0) return;
    const lastRow = firstRow + grid.length - 1;

    // Anything at or below this row has to move down first. When the block
    // starts past everything the sheet holds, no gap is needed — and that is not
    // an optimisation, it is what makes the first live run safe. A tab populated
    // from empty never opens a gap, so the one operation that could misalign
    // columns is not exercised until a tab that already holds rows is done,
    // deliberately and with someone watching.
    const needsInsert = firstRow <= occupiedTo;
    if (needsInsert) {
      operations.push({
        kind: 'insert-rows',
        row: firstRow,
        count: grid.length,
        reason: args.reason,
      });
    }

    operations.push({
      kind: 'write-cells',
      address: rangeAddress(firstColumn, firstRow, lastQueueColumnLetter(sheet.sheet, layout), lastRow),
      values: grid,
      row: firstRow,
      count: grid.length,
      purpose: 'rows',
    });

    operations.push({
      kind: 'write-cells',
      address: rangeAddress(
        layout.queue.syncIdColumn,
        firstRow,
        layout.queue.syncIdColumn,
        lastRow,
      ),
      values: syncIds,
      row: firstRow,
      count: syncIds.length,
      purpose: 'sync-ids',
    });

    for (const shade of shades) {
      for (const run of contiguousRuns(shade.columns, sheet.sheet, layout)) {
        operations.push({
          kind: 'shade',
          address: rangeAddress(
            run.firstColumn,
            firstRow + shade.offset,
            run.lastColumn,
            firstRow + shade.offset,
          ),
          color: BLANK_REQUIRED_FILL,
          columns: run.columns,
        });
      }
      shadedCells += shade.columns.length;
    }

    rowsAdded += grid.length;
    occupiedTo = Math.max(occupiedTo + (needsInsert ? grid.length : 0), lastRow);
  };

  // The TOTAL line, when the sheet already has one, is the lowest thing on it.
  // Rewriting it first means every insert happens afterwards and simply carries
  // it further down. When there is no TOTAL line it has to be written last
  // instead, once the final row is known — see the end of this function.
  if (sheet.totalRow !== undefined) {
    operations.push({
      kind: 'write-cells',
      address: rangeAddress(firstColumn, sheet.totalRow, firstColumn, sheet.totalRow),
      values: [[formatGrandTotal(totalAfter)]],
      row: sheet.totalRow,
      count: 1,
      purpose: 'total',
    });
  }

  // New camps go in at the append point as one contiguous run of dividers and
  // rows. Emitted before the grown blocks because they end up below them: a
  // grown last camp inserting at the same row pushes this run further down,
  // which is exactly where a brand new camp belongs.
  if (newBlocks.length > 0) {
    const grid: unknown[][] = [];
    const syncIds: unknown[][] = [];
    const shades: { offset: number; columns: QueueColumn[] }[] = [];

    for (const block of newBlocks) {
      grid.push(dividerRow(formatGroupHeader(block.camp, block.intents.length), columns));
      syncIds.push([null]);
      for (const intent of block.intents) {
        if (intent.blankRequired.length > 0) {
          shades.push({ offset: grid.length, columns: intent.blankRequired });
        }
        grid.push(patientRow(intent, columns));
        syncIds.push([intent.syncId]);
      }
    }

    emit({
      firstRow: appendPoint,
      grid,
      syncIds,
      shades,
      reason: `new camp block${newBlocks.length === 1 ? '' : 's'}: ${newBlocks
        .map((block) => block.camp)
        .join(', ')}`,
    });
  }

  for (const block of grownBlocks) {
    const grid: unknown[][] = [];
    const syncIds: unknown[][] = [];
    const shades: { offset: number; columns: QueueColumn[] }[] = [];

    for (const intent of block.intents) {
      if (intent.blankRequired.length > 0) {
        shades.push({ offset: grid.length, columns: intent.blankRequired });
      }
      grid.push(patientRow(intent, columns));
      syncIds.push([intent.syncId]);
    }

    emit({
      firstRow: block.atRow,
      grid,
      syncIds,
      shades,
      reason: `${block.intents.length} row${block.intents.length === 1 ? '' : 's'} into the existing ${block.camp} block`,
    });

    // The divider's count is a snapshot, so it is rewritten whenever the block
    // grows. Its row has not moved: every operation so far was below it.
    if (block.headerRow !== undefined) {
      operations.push({
        kind: 'write-cells',
        address: rangeAddress(firstColumn, block.headerRow, firstColumn, block.headerRow),
        values: [[formatGroupHeader(block.camp, block.existing + block.intents.length)]],
        row: block.headerRow,
        count: 1,
        purpose: 'divider',
      });
    }
  }

  // No TOTAL line on the sheet yet, so one is created below everything. This is
  // the single operation that cannot be addressed in the sheet's original
  // coordinates — it lands under rows that do not exist until the plan has run —
  // so it goes last, when nothing further can shift it.
  if (sheet.totalRow === undefined) {
    const row = occupiedTo + 1;
    operations.push({
      kind: 'write-cells',
      address: rangeAddress(firstColumn, row, firstColumn, row),
      values: [[formatGrandTotal(totalAfter)]],
      row,
      count: 1,
      purpose: 'total',
    });
  }

  return {
    destination: sheet.sheet,
    operations,
    placements: [...newBlocks, ...grownBlocks].map((block) => ({
      camp: block.camp,
      campKey: block.campKey,
      isNew: block.isNew,
      added: block.intents.length,
      existing: block.existing,
      atRow: block.atRow,
    })),
    appended: appends.length,
    rowsAdded,
    totalAfter,
    requiresInserts: operations.some((operation) => operation.kind === 'insert-rows'),
    shadedCells,
  };
}

/**
 * Rewrites every camp divider and the TOTAL line to match what the sheet
 * actually holds right now.
 *
 * The counts are snapshots taken when a row was written, which is what the
 * office asked for — but a snapshot goes stale the moment a row is removed, and
 * rows are removed whenever somebody marks one Done. So this recomputes them
 * from the parsed sheet and is safe to run at any time: on a sheet nobody has
 * touched it produces nothing.
 */
export function planCountRefresh(
  sheet: ParsedQueueSheet,
  layout: WorkbookLayout = LAYOUT,
): WriteCellsOperation[] {
  const first = layout.queue.firstColumn;
  const operations: WriteCellsOperation[] = [];

  for (const group of sheet.groups) {
    if (group.declaredCount === group.rows.length) continue;
    operations.push({
      kind: 'write-cells',
      address: rangeAddress(first, group.headerRow, first, group.headerRow),
      values: [[formatGroupHeader(group.camp, group.rows.length)]],
      row: group.headerRow,
      count: 1,
      purpose: 'divider',
    });
  }

  // The office's own tabs have no TOTAL line, so one is added below the body the
  // first time this runs. A plain write past the last used row — nothing is
  // inserted and nothing moves. Once it exists the branch above keeps it current.
  const totalRow = sheet.totalRow ?? Math.max(sheet.lastRow, sheet.firstDataRow - 1) + 1;
  if (sheet.totalDeclared !== sheet.rows.length) {
    operations.push({
      kind: 'write-cells',
      address: rangeAddress(first, totalRow, first, totalRow),
      values: [[formatGrandTotal(sheet.rows.length)]],
      row: totalRow,
      count: 1,
      purpose: 'total',
    });
  }

  return operations;
}

/**
 * Groups a row's blank required fields into contiguous column runs, so shading
 * `Last Name`, `First Name` and `Date of Birth` is one call rather than three.
 *
 * Graph takes a single rectangular range per fill request, so a run within one
 * row is as far as this can be batched.
 */
function contiguousRuns(
  columns: readonly QueueColumn[],
  destination: QueueSheetName,
  layout: WorkbookLayout,
): { firstColumn: string; lastColumn: string; columns: QueueColumn[] }[] {
  const order = queueColumnsFor(destination);
  const indexed = columns
    .map((column) => ({ column, index: order.indexOf(column) }))
    .filter((entry) => entry.index >= 0)
    .sort((a, b) => a.index - b.index);

  const runs: { firstColumn: string; lastColumn: string; columns: QueueColumn[] }[] = [];
  let current: { first: number; last: number; columns: QueueColumn[] } | undefined;

  for (const entry of indexed) {
    if (current && entry.index === current.last + 1) {
      current.last = entry.index;
      current.columns.push(entry.column);
      continue;
    }
    if (current) runs.push(toRun(current, layout));
    current = { first: entry.index, last: entry.index, columns: [entry.column] };
  }
  if (current) runs.push(toRun(current, layout));

  return runs;
}

function toRun(
  run: { first: number; last: number; columns: QueueColumn[] },
  layout: WorkbookLayout,
): { firstColumn: string; lastColumn: string; columns: QueueColumn[] } {
  return {
    firstColumn: offsetColumn(layout.queue.firstColumn, run.first),
    lastColumn: offsetColumn(layout.queue.firstColumn, run.last),
    columns: run.columns,
  };
}

/**
 * One patient, every queue column for that destination in order, blanks as null
 * rather than "".
 *
 * Date columns are converted to Excel serials on the way in. The office asked
 * for real dates, and a text date cannot be sorted, filtered by range or
 * formatted — `07/05/2026` and `7/5/2026` are two different strings and the same
 * day. A value that does not parse as a date is written through untouched.
 */
function patientRow(intent: AppendQueueRowIntent, columns: readonly QueueColumn[]): unknown[] {
  return columns.map((column) => {
    const value = intent.values[column];
    if (value === undefined || value === '') return null;
    return STYLE.dateColumns.includes(column) ? toDateCellValue(value) : value;
  });
}

/**
 * A divider is the label in the first column and nothing else — which is also
 * how `parseQueueSheet` recognizes one. Writing the full width of nulls keeps
 * the divider and its patient rows in a single range write, so the block cannot
 * be seen half-built.
 */
function dividerRow(label: string, columns: readonly QueueColumn[]): unknown[] {
  const row: unknown[] = new Array(columns.length).fill(null);
  row[0] = label;
  return row;
}

/**
 * Which spelling of a camp name goes on a new divider.
 *
 * The live workbook has `elky` and `Elky`, `bnos naale` and `Bnos Naale`.
 * Grouping is case-insensitive so they are already one camp; this only decides
 * what the header says. The most common spelling wins, ties broken
 * alphabetically so the same input always produces the same header.
 */
function displayCamp(intents: readonly AppendQueueRowIntent[]): string {
  const counts = new Map<string, number>();
  for (const intent of intents) {
    if (!intent.camp) continue;
    counts.set(intent.camp, (counts.get(intent.camp) ?? 0) + 1);
  }
  if (counts.size === 0) return NO_CAMP_LABEL;

  return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0]![0];
}
