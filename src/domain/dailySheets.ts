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

const DAY_MS = 24 * 60 * 60 * 1000;

export interface ScanOptions {
  allSheetNames: string[];
  /** Sheets an existing queue row points at. Always scanned. */
  referencedSheets: Iterable<string>;
  today: Date;
  hotDaysBack: number;
  hotDaysForward: number;
  coldBatchSize: number;
  /** Rotation position carried in the checkpoint. */
  cursor: number;
  layout?: WorkbookLayout;
}

export interface ScanSelection {
  /** Scanned every cycle. */
  hot: string[];
  /** This cycle's slice of the rotation over everything else. */
  cold: string[];
  /** Everything to read this cycle: hot plus cold. */
  sheets: string[];
  /** Persist this as the next cursor. */
  nextCursor: number;
  /** How many cycles a full sweep of the cold pool takes. */
  sweepCycles: number;
  totalDaily: number;
  /** Daily-pattern sheets whose name did not parse as a date. A naming problem. */
  unparseable: string[];
}

/**
 * Tiered scan.
 *
 * The office creates daily sheets **in advance** — on 3 August there may already
 * be sheets for the 4th through the 8th, empty until the day arrives. That rules
 * out "the N sheets with the latest dates": the latest-dated sheets are the empty
 * future ones, and today's would be crowded out.
 *
 * It also has to cover the whole year, because a staff member can put a keyword
 * into column B on a sheet from three weeks ago at any time. But reading every
 * sheet every cycle is a few hundred Graph calls at a five-second cadence, which
 * is a throttling incident rather than a design.
 *
 * So:
 *   - **hot** — a date window around today, in both directions so tomorrow's
 *     pre-created sheet is picked up the moment staff start filling it, plus
 *     every sheet an existing queue row references. Read every cycle.
 *   - **cold** — everything else, a fixed-size slice per cycle, rotating through
 *     the checkpointed cursor. Full-year coverage at bounded cost.
 *
 * Nothing is ever excluded, only deferred: the guarantee is that every daily
 * sheet is read within `sweepCycles` changed cycles, and anything recent or
 * already queued is read immediately.
 */
export function planScan(options: ScanOptions): ScanSelection {
  const layout = options.layout ?? LAYOUT;
  const daily = options.allSheetNames.filter((name) => isDailySheetName(name, layout));

  const dated: { name: string; date: Date }[] = [];
  const unparseable: string[] = [];
  for (const name of daily) {
    const date = layout.parseDailySheetDate(name);
    if (date) dated.push({ name, date });
    else unparseable.push(name);
  }

  const todayUtc = Date.UTC(
    options.today.getUTCFullYear(),
    options.today.getUTCMonth(),
    options.today.getUTCDate(),
  );
  const lower = todayUtc - Math.max(options.hotDaysBack, 0) * DAY_MS;
  const upper = todayUtc + Math.max(options.hotDaysForward, 0) * DAY_MS;

  const hot = new Set<string>();
  for (const entry of dated) {
    const time = entry.date.getTime();
    if (time >= lower && time <= upper) hot.add(entry.name);
  }

  // Off-season, or a workbook whose sheets are all in the past: the window
  // catches nothing and only the rotation would run. Fall back to the most
  // recent sheets that are not in the future, so the job stays responsive to
  // whatever staff are actually working on.
  if (hot.size === 0 && dated.length > 0) {
    const notFuture = dated
      .filter((entry) => entry.date.getTime() <= todayUtc)
      .sort((a, b) => b.date.getTime() - a.date.getTime());
    for (const entry of (notFuture.length > 0 ? notFuture : dated).slice(0, 3)) {
      hot.add(entry.name);
    }
  }

  // A queue row's source sheet is scanned however old it is, so a row sitting on
  // a June date still reconciles every cycle rather than waiting for its turn.
  for (const name of options.referencedSheets) {
    if (daily.includes(name)) hot.add(name);
  }

  // Recent first, so a cold start surfaces the likeliest edits soonest. The
  // unparseable ones go last: they still get scanned, but they are a naming
  // problem to fix rather than a population to prioritise.
  const coldPool = [
    ...dated
      .filter((entry) => !hot.has(entry.name))
      .sort((a, b) => b.date.getTime() - a.date.getTime())
      .map((entry) => entry.name),
    ...unparseable.filter((name) => !hot.has(name)),
  ];

  const cold: string[] = [];
  let nextCursor = 0;

  if (coldPool.length > 0) {
    const batch = Math.min(Math.max(options.coldBatchSize, 0), coldPool.length);
    const start =
      Number.isFinite(options.cursor) && options.cursor >= 0
        ? Math.floor(options.cursor) % coldPool.length
        : 0;
    for (let i = 0; i < batch; i++) {
      cold.push(coldPool[(start + i) % coldPool.length]!);
    }
    nextCursor = batch > 0 ? (start + batch) % coldPool.length : start;
  }

  return {
    hot: [...hot],
    cold,
    sheets: [...new Set([...hot, ...cold])],
    nextCursor,
    sweepCycles:
      coldPool.length === 0 || options.coldBatchSize <= 0
        ? 0
        : Math.ceil(coldPool.length / options.coldBatchSize),
    totalDaily: daily.length,
    unparseable,
  };
}
