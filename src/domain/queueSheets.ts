import {
  LAYOUT,
  QUEUE_COLUMNS,
  type QueueColumn,
  type QueueSheetName,
  type WorkbookLayout,
} from '../config';
import { cellFromGrid, columnToIndex, isBlank, offsetColumn, parseAddress } from './cells';
import { normalizeCamp } from './dailySheets';
import { normalizeSyncId } from './syncId';
import type { RangeData } from '../graph/workbook';

export interface QueueRow {
  sheet: QueueSheetName;
  row: number;
  syncId: string | undefined;
  values: Partial<Record<QueueColumn, unknown>>;
  /** Camp group this row currently sits under, from the nearest header above it. */
  camp: string | undefined;
}

export interface CampGroup {
  camp: string;
  /** 1-based row of the `CampName - N patients` header. */
  headerRow: number;
  declaredCount: number;
  /** Rows belonging to this group, in sheet order. */
  rows: QueueRow[];
}

export interface ParsedQueueSheet {
  sheet: QueueSheetName;
  rows: QueueRow[];
  groups: CampGroup[];
  /** Last used row on the sheet, for append positioning. */
  lastRow: number;
}

/** `Camp Ramah - 12 patients` */
const HEADER_PATTERN = /^\s*(.+?)\s*-\s*(\d+)\s+patients?\s*$/i;

/** `TOTAL - 374 patients`, written once at the foot of each queue sheet. */
const TOTAL_LABEL = 'TOTAL';
const TOTAL_PATTERN = /^\s*TOTAL\s*-\s*(\d+)\s+patients?\s*$/i;

/**
 * The camp divider row: one cell, camp name and a count.
 *
 * The count is a snapshot taken when the row was written, not a live figure —
 * it says what the queue held at that moment. Keeping it in a single cell (as
 * opposed to a formula, or a second column) is deliberate: this workbook is
 * being rescued from a formula load, and a static string costs nothing to
 * recalculate.
 */
export function formatGroupHeader(camp: string, count: number): string {
  return `${camp} - ${count} patient${count === 1 ? '' : 's'}`;
}

export function formatGrandTotal(count: number): string {
  return `${TOTAL_LABEL} - ${count} patient${count === 1 ? '' : 's'}`;
}

export function parseGrandTotal(value: unknown): number | undefined {
  if (isBlank(value)) return undefined;
  const match = TOTAL_PATTERN.exec(String(value));
  return match ? Number(match[1]) : undefined;
}

/**
 * True for the grand total row, which must never be mistaken for a camp called
 * "TOTAL" — otherwise every rebuild would nest a group inside the total.
 */
export function isGrandTotalRow(value: unknown): boolean {
  return parseGrandTotal(value) !== undefined;
}

export function parseGroupHeader(value: unknown): { camp: string; count: number } | undefined {
  if (isBlank(value)) return undefined;
  if (isGrandTotalRow(value)) return undefined;
  const match = HEADER_PATTERN.exec(String(value));
  if (!match) return undefined;
  return { camp: match[1]!.trim(), count: Number(match[2]) };
}

/**
 * Parses one queue sheet's usedRange into rows and camp groups.
 *
 * A row is a data row when it carries a SyncID or any patient field; a row whose
 * first column parses as `Camp - N patients` is a group header. Blank spacer
 * rows are ignored. Instruction blocks above `firstDataRow` are never examined.
 */
export function parseQueueSheet(
  sheet: QueueSheetName,
  used: RangeData,
  layout: WorkbookLayout = LAYOUT,
): ParsedQueueSheet {
  const { startRow, startColumn } = parseAddress(used.address);
  const first = layout.queue.firstColumn;

  const rows: QueueRow[] = [];
  const groups: CampGroup[] = [];
  let currentGroup: CampGroup | undefined;

  const firstRow = Math.max(startRow, layout.queue.firstDataRow);
  const lastRow = startRow + used.values.length - 1;

  for (let row = firstRow; row <= lastRow; row++) {
    const cell = (column: string): unknown =>
      cellFromGrid(used.values, startRow, startColumn, row, column);

    const header = parseGroupHeader(cell(first));
    if (header) {
      currentGroup = {
        camp: header.camp,
        headerRow: row,
        declaredCount: header.count,
        rows: [],
      };
      groups.push(currentGroup);
      continue;
    }

    const values: Partial<Record<QueueColumn, unknown>> = {};
    let hasValue = false;
    QUEUE_COLUMNS.forEach((column, index) => {
      const value = cell(offsetColumn(first, index));
      values[column] = value;
      if (!isBlank(value)) hasValue = true;
    });

    const syncId = normalizeSyncId(cell(layout.queue.syncIdColumn));
    if (!hasValue && !syncId) continue;

    const queueRow: QueueRow = {
      sheet,
      row,
      syncId,
      values,
      camp: currentGroup?.camp,
    };
    rows.push(queueRow);
    currentGroup?.rows.push(queueRow);
  }

  return { sheet, rows, groups, lastRow };
}

/**
 * Matches a configured sheet name against the workbook's actual tab names,
 * forgiving the differences a human would ignore: leading and trailing spaces,
 * doubled internal spaces, and casing.
 *
 * The live workbook has a tab literally named `Not Accepted ` with a trailing
 * space. An exact-match lookup silently reports it as "not present", and a queue
 * with no sheet is a queue that never gets populated — a failure that looks like
 * nothing happening at all.
 */
export function normalizeSheetName(name: string): string {
  return name.replace(/\s+/g, ' ').trim().toLowerCase();
}

export function resolveSheetName(
  wanted: string,
  actualNames: readonly string[],
): string | undefined {
  const exact = actualNames.find((name) => name === wanted);
  if (exact) return exact;

  const target = normalizeSheetName(wanted);
  const matches = actualNames.filter((name) => normalizeSheetName(name) === target);

  // More than one tab differing only by whitespace or case is ambiguous, and
  // picking one silently could populate a stale tab. Better to report nothing.
  return matches.length === 1 ? matches[0] : undefined;
}

/** Column letter on a queue sheet for a given queue column. */
export function queueColumnLetter(
  column: QueueColumn,
  layout: WorkbookLayout = LAYOUT,
): string {
  const index = QUEUE_COLUMNS.indexOf(column);
  if (index === -1) throw new Error(`Not a queue column: ${column}`);
  return offsetColumn(layout.queue.firstColumn, index);
}

export function lastQueueColumnLetter(layout: WorkbookLayout = LAYOUT): string {
  return offsetColumn(layout.queue.firstColumn, QUEUE_COLUMNS.length - 1);
}

/**
 * Sanity check for the SyncID column placement: it must sit clear of the data
 * columns on every sheet, not just look far away.
 */
export function assertSyncIdColumnIsClear(layout: WorkbookLayout = LAYOUT): void {
  const lastData = columnToIndex(lastQueueColumnLetter(layout));
  if (columnToIndex(layout.queue.syncIdColumn) <= lastData) {
    throw new Error(
      `Queue SyncID column ${layout.queue.syncIdColumn} overlaps the data columns ` +
        `(which end at ${lastQueueColumnLetter(layout)}).`,
    );
  }
}
