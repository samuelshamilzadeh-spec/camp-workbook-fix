import { LAYOUT, type QueueColumn, type WorkbookLayout } from '../config';
import { cellFromGrid, isBlank, parseAddress } from './cells';
import { classifyStatus, isAmbiguous, type StatusOutcome } from './status';
import { normalizeSyncId } from './syncId';
import type { RangeData } from '../graph/workbook';

export interface DailyRow {
  sheet: string;
  /** 1-based sheet row number. Not unique across sheets. */
  row: number;
  syncId: string | undefined;
  statusRaw: unknown;
  outcome: StatusOutcome;
  ambiguous: boolean;
  camp: string | undefined;
  /** Queue column -> daily cell value. Only columns present in the field map. */
  fields: Partial<Record<QueueColumn, unknown>>;
}

export interface ParsedDailySheet {
  sheet: string;
  rows: DailyRow[];
  /** Rows whose column B held something we do not recognize. Count only, never values. */
  unrecognizedRows: number[];
}

/**
 * Turns one sheet's usedRange into rows.
 *
 * Filtering happens in memory because reads are cheap and writes are expensive:
 * one call for the whole sheet beats a targeted read per row, every time.
 */
export function parseDailySheet(
  sheetName: string,
  used: RangeData,
  layout: WorkbookLayout = LAYOUT,
): ParsedDailySheet {
  const { startRow, startColumn } = parseAddress(used.address);
  const rows: DailyRow[] = [];
  const unrecognizedRows: number[] = [];

  const firstRow = Math.max(startRow, layout.daily.firstDataRow);
  const lastRow = startRow + used.values.length - 1;

  for (let row = firstRow; row <= lastRow; row++) {
    const cell = (column: string): unknown =>
      cellFromGrid(used.values, startRow, startColumn, row, column);

    const statusRaw = cell(layout.daily.statusColumn);
    const outcome = classifyStatus(statusRaw);

    if (outcome.kind === 'blank') continue;
    if (outcome.kind === 'unrecognized') {
      unrecognizedRows.push(row);
      continue;
    }

    const fields: Partial<Record<QueueColumn, unknown>> = {};
    for (const [queueColumn, dailyColumn] of Object.entries(layout.daily.fieldColumns)) {
      if (dailyColumn) fields[queueColumn as QueueColumn] = cell(dailyColumn);
    }

    rows.push({
      sheet: sheetName,
      row,
      syncId: normalizeSyncId(cell(layout.daily.syncIdColumn)),
      statusRaw,
      outcome,
      ambiguous: isAmbiguous(outcome),
      camp: normalizeCamp(cell(layout.daily.campColumn)),
      fields,
    });
  }

  return { sheet: sheetName, rows, unrecognizedRows };
}

/**
 * Camp names are a grouping key, so they need to be stable across casing and
 * stray whitespace: `Camp Ramah `, `camp ramah` and `Camp  Ramah` must land in
 * one group rather than three.
 *
 * What is deliberately NOT done here is abbreviation expansion. `CR` and
 * `Camp Ramah` may or may not be the same camp, and guessing would silently
 * merge two camps' patients. `npm run inspect` prints the distinct camp values
 * so a human can decide whether an alias table is needed.
 */
export function normalizeCamp(raw: unknown): string | undefined {
  if (isBlank(raw)) return undefined;
  return String(raw).replace(/\s+/g, ' ').trim();
}

/** Case-insensitive grouping key. The display name keeps whatever staff typed. */
export function campKey(camp: string | undefined): string {
  return (camp ?? '(no camp)').toLowerCase();
}

export function isDailySheetName(name: string, layout: WorkbookLayout = LAYOUT): boolean {
  if (layout.knownNonDailySheets.includes(name)) return false;
  return layout.dailySheetPattern.test(name);
}

/**
 * Bounded scan: the most recent N daily sheets, plus every sheet an existing
 * queue row points at, so a stale queue row on an old date still reconciles.
 * Never every sheet in the workbook.
 */
export function selectSheetsToScan(
  allSheetNames: string[],
  referencedSheets: Iterable<string>,
  recentCount: number,
  layout: WorkbookLayout = LAYOUT,
): string[] {
  const daily = allSheetNames.filter((name) => isDailySheetName(name, layout));

  const dated = daily
    .map((name) => ({ name, date: layout.parseDailySheetDate(name) }))
    .filter((entry): entry is { name: string; date: Date } => entry.date !== null)
    .sort((a, b) => b.date.getTime() - a.date.getTime());

  const selected = new Set(dated.slice(0, Math.max(recentCount, 0)).map((e) => e.name));

  // Sheets whose name we could not parse as a date still match the daily
  // pattern, so they are real sheets we cannot order. Including the tail of them
  // is safer than dropping them, and `inspect` surfaces them for a naming fix.
  for (const name of daily) {
    if (layout.parseDailySheetDate(name) === null) selected.add(name);
  }

  for (const name of referencedSheets) {
    if (allSheetNames.includes(name)) selected.add(name);
  }

  return [...selected];
}
