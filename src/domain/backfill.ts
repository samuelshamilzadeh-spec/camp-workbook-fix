import { LAYOUT, type WorkbookLayout } from '../config';
import { cellFromGrid, parseAddress } from './cells';
import { classifyStatus } from './status';
import { newSyncId, normalizeSyncId } from './syncId';

export interface SheetStampPlan {
  sheet: string;
  /** Contiguous block of rows the write will cover, from first stamp to last. */
  firstRow: number;
  lastRow: number;
  /** Every row in the block, mapped to the ID it should hold afterwards. */
  stamps: Map<number, string>;
  /** How many of those already had an ID and are being preserved, not minted. */
  alreadyStamped: number;
  /** Rows receiving a newly minted ID. */
  newStamps: number;
  headerNeeded: boolean;
}

/**
 * Phase 0 planning for one sheet.
 *
 * Stamps only rows carrying a queued keyword in column B. Not a bulk walk over
 * blanks: a patient who never receives a keyword never needs an ID, which also
 * excludes group headers and instruction blocks for free.
 *
 * The write is a single range from the first stamp to the last, so rows in
 * between that already carry an ID must be written back unchanged. Writing
 * `null` into one would blank an ID a queue row depends on — the one
 * unrecoverable mistake this script could make.
 */
export function planSheetStamps(
  sheet: string,
  used: { address: string; values: unknown[][] },
  layout: WorkbookLayout = LAYOUT,
  mintId: () => string = newSyncId,
): SheetStampPlan {
  const { startRow, startColumn } = parseAddress(used.address);
  const minted = new Map<number, string>();
  const preserved = new Map<number, string>();

  const firstDataRow = Math.max(startRow, layout.daily.firstDataRow);
  const lastUsedRow = startRow + used.values.length - 1;

  for (let row = firstDataRow; row <= lastUsedRow; row++) {
    const status = cellFromGrid(
      used.values,
      startRow,
      startColumn,
      row,
      layout.daily.statusColumn,
    );
    if (classifyStatus(status).kind !== 'queued') continue;

    const existing = normalizeSyncId(
      cellFromGrid(used.values, startRow, startColumn, row, layout.daily.syncIdColumn),
    );
    if (existing) {
      preserved.set(row, existing);
      continue;
    }
    minted.set(row, mintId());
  }

  const mintedRows = [...minted.keys()];
  const firstRow = mintedRows.length > 0 ? Math.min(...mintedRows) : 0;
  const lastRow = mintedRows.length > 0 ? Math.max(...mintedRows) : 0;

  // Preserve every existing ID inside the block, not just those on queued rows.
  // A terminal or unrecognized row sandwiched between two queued ones may still
  // carry an ID from an earlier run, and the single range write covers it.
  const stamps = new Map(minted);
  if (mintedRows.length > 0) {
    for (let row = firstRow; row <= lastRow; row++) {
      if (stamps.has(row)) continue;
      const existing = normalizeSyncId(
        cellFromGrid(used.values, startRow, startColumn, row, layout.daily.syncIdColumn),
      );
      if (existing) stamps.set(row, existing);
    }
  }

  const headerValue = cellFromGrid(
    used.values,
    startRow,
    startColumn,
    layout.daily.headerRow,
    layout.daily.syncIdColumn,
  );

  return {
    sheet,
    firstRow,
    lastRow,
    stamps,
    alreadyStamped: preserved.size,
    newStamps: minted.size,
    headerNeeded: String(headerValue ?? '').trim() !== 'SyncID',
  };
}

/**
 * The values for the single range write, first row to last. Empty when there is
 * nothing to mint — `firstRow` is 0 in that case, and a caller that wrote the
 * result anyway would blank a cell.
 */
export function stampWriteValues(plan: SheetStampPlan): unknown[][] {
  if (plan.newStamps === 0) return [];
  const values: unknown[][] = [];
  for (let row = plan.firstRow; row <= plan.lastRow; row++) {
    values.push([plan.stamps.get(row) ?? null]);
  }
  return values;
}
