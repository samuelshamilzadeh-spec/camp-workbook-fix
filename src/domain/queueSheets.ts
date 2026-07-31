import {
  LAYOUT,
  QUEUE_COLUMNS,
  RESOLVED_VALUES,
  queueColumnsFor,
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
  /** What the `Resolved` cell says. */
  resolved: ResolvedSignal;
}

/**
 * The state of a row's `Resolved` cell.
 *
 * `unrecognized` is a deliberate third state rather than a synonym for `none`.
 * Somebody typed something in that column meaning to say something, and quietly
 * ignoring it would leave them believing the row had been dealt with. It is
 * reported so a human can look, and never acted on.
 */
export type ResolvedSignal =
  | { kind: 'none' }
  | { kind: 'resolved'; raw: string }
  | { kind: 'unrecognized'; raw: string };

export function readResolvedSignal(value: unknown): ResolvedSignal {
  if (isBlank(value)) return { kind: 'none' };
  const raw = String(value).trim();
  return RESOLVED_VALUES.includes(raw.toLowerCase())
    ? { kind: 'resolved', raw }
    : { kind: 'unrecognized', raw };
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
  /** 1-based row holding the column headers, as found on this sheet. */
  headerRow: number;
  /** 1-based first row that can hold data. */
  firstDataRow: number;
  /** Row of the `TOTAL - N patients` line, when the sheet has one. */
  totalRow: number | undefined;
  /**
   * The count that line declares. A snapshot from when it was written, so it can
   * legitimately disagree with `rows.length` — comparing the two after a write
   * is how you check the write landed.
   */
  totalDeclared: number | undefined;
  /**
   * False when the header row could not be found and the configured LAYOUT was
   * used instead. Worth surfacing before a write: it means the shape is a guess.
   */
  shapeDetected: boolean;
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
 * Column headings a queue tab might carry, normalized to letters and digits.
 *
 * Two vocabularies are in play. The four mirror-layout tabs were written by
 * hand years ago, and the two tabs this project created copied their header row
 * from `Missing Info (New)`, which uses the daily sheet's abbreviations
 * (`LAST Nm`, `INS CARRIER`). Both are listed, because the point of detection is
 * to work on a tab nobody has described to us.
 */
const HEADER_MARKERS = new Set([
  'dateofvisit',
  'sourcerow',
  'resolved',
  'lastname',
  'lastnm',
  'firstname',
  'firstnm',
  'dateofbirth',
  'dob',
  'billingaddress',
  'phonenumber',
  'inscarrier',
  'insurancecarrier',
  'insid',
  'insuranceid',
  'medicaid',
  'medicalhistory',
  'allergies',
]);

/** Two markers is enough to be sure, and few enough that a sparse header still counts. */
const HEADER_MARKERS_REQUIRED = 2;

/** How far down to look. A header below this is not a header, it is data. */
const HEADER_SEARCH_ROWS = 20;

function normalizeLabel(value: unknown): string {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

export interface QueueShape {
  headerRow: number;
  firstDataRow: number;
  /** False when nothing header-shaped was found and LAYOUT was used instead. */
  detected: boolean;
  /** Which markers were recognized, for reporting before a write. */
  markers: string[];
}

/**
 * Finds the header row by reading it off the sheet rather than trusting config.
 *
 * `LAYOUT.queue` describes the four mirror tabs — instructions in rows 1-7,
 * headers on row 9, data from row 10 — and that is right for them. It is wrong
 * for `Verify Insurance` and `United Refuah`, which this project created with
 * `create-queue-sheet.ts` and which carry their header on row 1 with no
 * instruction block above it.
 *
 * Trusting the config on those two would skip rows 2-9 when reading and leave
 * eight blank rows above the first patient when writing — on a tab that is
 * supposed to be built correctly from scratch. Since the answer is sitting in
 * the sheet, read it from there and keep the config as the fallback.
 */
export function detectQueueShape(
  used: RangeData,
  layout: WorkbookLayout = LAYOUT,
): QueueShape {
  const { startRow } = parseAddress(used.address);
  const limit = Math.min(used.values.length, HEADER_SEARCH_ROWS);

  for (let i = 0; i < limit; i++) {
    const found = new Set<string>();
    for (const cell of used.values[i] ?? []) {
      const label = normalizeLabel(cell);
      if (HEADER_MARKERS.has(label)) found.add(label);
    }
    if (found.size >= HEADER_MARKERS_REQUIRED) {
      const headerRow = startRow + i;
      return {
        headerRow,
        firstDataRow: headerRow + 1,
        detected: true,
        markers: [...found].sort(),
      };
    }
  }

  return {
    headerRow: layout.queue.headerRow,
    firstDataRow: layout.queue.firstDataRow,
    detected: false,
    markers: [],
  };
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
  const shape = detectQueueShape(used, layout);
  const columns = queueColumnsFor(sheet);

  const rows: QueueRow[] = [];
  const groups: CampGroup[] = [];
  let currentGroup: CampGroup | undefined;
  let totalRow: number | undefined;
  let totalDeclared: number | undefined;

  const firstRow = Math.max(startRow, shape.firstDataRow);
  const lastRow = startRow + used.values.length - 1;

  for (let row = firstRow; row <= lastRow; row++) {
    const cell = (column: string): unknown =>
      cellFromGrid(used.values, startRow, startColumn, row, column);

    // A camp divider is a row with something in the first column and nothing
    // anywhere else. The live tabs write a bare name (`Achim`); this code
    // writes a name plus a count (`Achim - 12 patients`). Both are dividers,
    // and treating a bare one as a patient row is what made 74 of them fail to
    // link to a source row that was never going to exist.
    const firstCell = cell(first);
    const restBlank =
      !isBlank(firstCell) &&
      columns.slice(1).every((_, index) => isBlank(cell(offsetColumn(first, index + 1))));

    const header = parseGroupHeader(firstCell) ?? (restBlank && !isGrandTotalRow(firstCell)
      ? { camp: String(firstCell).trim(), count: 0 }
      : undefined);

    if (restBlank && isGrandTotalRow(firstCell)) {
      // Remembered, not parsed as a group: the appender has to rewrite this
      // line with the new count, and it can only do that if it knows the row.
      totalRow = row;
      totalDeclared = parseGrandTotal(firstCell);
      continue;
    }

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
    columns.forEach((column, index) => {
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
      resolved: readResolvedSignal(values['Resolved']),
    };
    rows.push(queueRow);
    currentGroup?.rows.push(queueRow);
  }

  return {
    sheet,
    rows,
    groups,
    lastRow,
    headerRow: shape.headerRow,
    firstDataRow: shape.firstDataRow,
    totalRow,
    totalDeclared,
    shapeDetected: shape.detected,
  };
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

/**
 * Column letter of the LAST queue column for a destination.
 *
 * Destination-scoped, because United Refuah has no `Resolved` column and so
 * every column after `Source Row` sits one place to its left there.
 */
export function lastQueueColumnLetter(
  destination: QueueSheetName,
  layout: WorkbookLayout = LAYOUT,
): string {
  return offsetColumn(layout.queue.firstColumn, queueColumnsFor(destination).length - 1);
}

/**
 * Sanity check for the SyncID column placement: it must sit clear of the data
 * columns on every sheet, not just look far away.
 */
export function assertSyncIdColumnIsClear(layout: WorkbookLayout = LAYOUT): void {
  // Checked against the WIDEST layout, so adding a column cannot quietly grow
  // the data past the SyncID column on one tab while the check passes on another.
  const lastData = columnToIndex(offsetColumn(layout.queue.firstColumn, QUEUE_COLUMNS.length - 1));
  if (columnToIndex(layout.queue.syncIdColumn) <= lastData) {
    throw new Error(
      `Queue SyncID column ${layout.queue.syncIdColumn} overlaps the data columns ` +
        `(which end at ${offsetColumn(layout.queue.firstColumn, QUEUE_COLUMNS.length - 1)}).`,
    );
  }
}
