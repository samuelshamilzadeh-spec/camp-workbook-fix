import { describe, expect, it } from 'vitest';
import {
  MONTH_SEGMENTS,
  QUEUE_FAMILIES,
  QUEUE_SHEET_NAMES,
  QUEUE_SHEET_TABS,
  REQUIRED_FIELDS,
  SPLIT_FAMILIES,
  familyOf,
  isAppendOnly,
  isSplitFamily,
  queueColumnsFor,
  queueTabFor,
  segmentOf,
  type QueueSheetName,
} from '../src/config';
import {
  MAX_SHEET_NAME_LENGTH,
  assertQueueTabNamesFit,
  parseQueueSheet,
} from '../src/domain/queueSheets';
import { queueSheetFor, segmentForVisit, visitDate } from '../src/domain/segments';
import { parseDailySheet } from '../src/domain/dailySheets';
import { planQueueAppend } from '../src/domain/append';
import { reconcile, type AppendQueueRowIntent } from '../src/domain/reconcile';
import { planBlankRequiredShading } from '../src/domain/style';
import type { RangeData } from '../src/graph/workbook';

describe('a visit date resolves to a month', () => {
  it('reads the daily sheet name the office actually uses', () => {
    expect(segmentForVisit('July 30, 2026')).toBe('July');
    expect(segmentForVisit('June 1, 2026')).toBe('June');
    expect(segmentForVisit('August 31, 2026')).toBe('August');
    expect(segmentForVisit('Sept. 3, 2026')).toBe('Other');
  });

  it('reads the Excel serial the queue tabs hold', () => {
    // Both representations are in play: a row this project appends carries its
    // source sheet's NAME, and a row already on a queue tab carries a serial,
    // because `npm run migrate` converted the column to real dates. Reading only
    // one of them would put every row from the other into `Other`.
    const serial = (iso: string) =>
      Math.round((Date.parse(`${iso}T00:00:00Z`) - Date.UTC(1899, 11, 30)) / 86400000);

    expect(segmentForVisit(serial('2026-06-15'))).toBe('June');
    expect(segmentForVisit(serial('2026-07-04'))).toBe('July');
    expect(segmentForVisit(serial('2026-08-20'))).toBe('August');
    expect(segmentForVisit(serial('2026-05-31'))).toBe('Other');
  });

  it('reads the ISO and US numeric forms too', () => {
    expect(segmentForVisit('2026-07-30')).toBe('July');
    expect(segmentForVisit('7/30/2026')).toBe('July');
    expect(segmentForVisit('06/01/2026')).toBe('June');
  });

  it('never fails to place a row: anything unreadable is Other, not nothing', () => {
    // A row that cannot be placed must still land somewhere a person can see it.
    // Returning undefined here is how a patient falls out of the system quietly.
    for (const value of [undefined, null, '', '   ', 'not a date', 'TBC', {}, true]) {
      expect(segmentForVisit(value)).toBe('Other');
    }
  });

  it('does not roll an impossible date forward into a real month', () => {
    // `14/24/2014` is on the live workbook. JavaScript's Date rolls it to
    // February 2015 without a word; placing a row by that would be a guess.
    expect(visitDate('14/24/2014')).toBeUndefined();
    expect(segmentForVisit('14/24/2014')).toBe('Other');
  });

  it('ignores the year, which is right for one season and wrong for two', () => {
    // Documented in MONTH_SEGMENTS: a second season needs the year in the tab
    // names. This test exists so that changing it is a deliberate act.
    expect(segmentForVisit('July 30, 2027')).toBe('July');
  });
});

describe('a keyword picks a queue, a date picks its tab', () => {
  it('splits the two families the office asked for', () => {
    expect(queueSheetFor('Missing Info', 'July 30, 2026')).toBe('Missing Info - July');
    expect(queueSheetFor('Ineligible & Inactive', 'June 2, 2026')).toBe(
      'Ineligible & Inactive - June',
    );
    expect(queueSheetFor('Missing Info', 'nonsense')).toBe('Missing Info - Other');
  });

  it('leaves the other three alone, whatever the date', () => {
    for (const family of QUEUE_FAMILIES.filter((f) => !isSplitFamily(f))) {
      expect(queueSheetFor(family, 'July 30, 2026')).toBe(family);
      expect(queueSheetFor(family, 'nonsense')).toBe(family);
    }
  });

  it('round-trips a tab name back to its family and segment', () => {
    for (const name of QUEUE_SHEET_NAMES) {
      const family = familyOf(name);
      expect(QUEUE_FAMILIES).toContain(family);
      expect(queueTabFor(family, segmentOf(name))).toBe(name);
    }
  });

  it('names a bad tab rather than returning a plausible-looking string', () => {
    // familyOf keys into REQUIRED_FIELDS. A wrong answer there is `undefined`,
    // and the next line asks it for `.filter` — a TypeError several frames from
    // the mistake.
    expect(() => familyOf('Missing Infe - July' as QueueSheetName)).toThrow(/does not name a queue/);
  });
});

describe('the tab list', () => {
  it('is three unsplit queues plus four tabs for each split one', () => {
    expect(QUEUE_SHEET_NAMES).toHaveLength(3 + SPLIT_FAMILIES.length * (MONTH_SEGMENTS.length + 1));
    expect(QUEUE_SHEET_NAMES).toContain('Missing Info - June');
    expect(QUEUE_SHEET_NAMES).toContain('Ineligible & Inactive - Other');
    expect(QUEUE_SHEET_NAMES).toContain('United Refuah');
    // The pre-split names are no longer tabs.
    expect(QUEUE_SHEET_NAMES).not.toContain('Missing Info' as QueueSheetName);
  });

  it('has a tab mapping for every name, and no name Excel would refuse', () => {
    for (const name of QUEUE_SHEET_NAMES) {
      expect(QUEUE_SHEET_TABS[name]).toBeTruthy();
    }
    expect(() => assertQueueTabNamesFit()).not.toThrow();
    // The one that would break first if a month were added.
    expect('Ineligible & Inactive - August'.length).toBeLessThanOrEqual(MAX_SHEET_NAME_LENGTH);
    expect('Ineligible & Inactive - September'.length).toBeGreaterThan(MAX_SHEET_NAME_LENGTH);
  });

  it('keeps every per-queue rule keyed by family, so the months agree', () => {
    // REQUIRED_FIELDS is keyed by family and must stay that way: keyed by tab it
    // would need eight more entries, and a missing one is not an error but an
    // `undefined` that the next `.filter` turns into a TypeError.
    expect(Object.keys(REQUIRED_FIELDS).sort()).toEqual([...QUEUE_FAMILIES].sort());

    for (const name of QUEUE_SHEET_NAMES) {
      const family = familyOf(name);
      expect(REQUIRED_FIELDS[family]).toBeDefined();
      // A month tab carries the same columns and the same append-only rule as
      // its family: cutting a queue into months changes where rows sit, nothing
      // about what they are.
      expect(queueColumnsFor(name)).toEqual(queueColumnsFor(queueTabFor(family, undefined)));
      expect(isAppendOnly(name)).toBe(family === 'United Refuah');
    }
  });
});

// --- The same rules, exercised through a real reconcile ---------------------

const WIDTH = 53;

function dailyRange(
  sheet: string,
  rows: { status?: string; camp?: string; last?: string; first?: string; phone?: string; syncId?: string }[],
): RangeData {
  const grid: unknown[][] = [new Array(WIDTH).fill(null)];
  for (const row of rows) {
    const cells = new Array(WIDTH).fill(null);
    cells[1] = row.status ?? null;
    cells[2] = row.camp ?? null;
    cells[3] = row.last ?? null;
    cells[4] = row.first ?? null;
    cells[11] = row.phone ?? null;
    cells[52] = row.syncId ?? null;
    grid.push(cells);
  }
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: WIDTH, values: grid };
}

/** A queue tab holding one already-linked row, with its Date of Visit filled in. */
function queueRange(sheet: QueueSheetName, dateOfVisit: string, syncId: string): RangeData {
  const grid: unknown[][] = [new Array(WIDTH).fill(null)];
  const cells = new Array(WIDTH).fill(null);
  cells[0] = dateOfVisit;
  cells[3] = 'A';
  cells[4] = 'B';
  cells[11] = '555';
  cells[52] = syncId;
  grid.push(cells);
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: WIDTH, values: grid };
}

const ids = () => {
  let n = 0;
  return () => `S${String(++n).padStart(12, '0')}`;
};

describe('reconciling across the monthly tabs', () => {
  it('sends a new row to the tab for the month it was seen in', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('June 12, 2026', dailyRange('June 12, 2026', [
          { status: 'missing info', camp: 'Achim', last: 'A', first: 'B' },
        ])),
        parseDailySheet('August 3, 2026', dailyRange('August 3, 2026', [
          { status: 'ineligible', camp: 'Achim', last: 'C', first: 'D' },
        ])),
      ],
      queues: [],
      newSyncId: ids(),
    });

    const appends = plan.intents.filter(
      (i): i is AppendQueueRowIntent => i.kind === 'append-queue-row',
    );
    expect(appends.map((a) => a.destination)).toEqual([
      'Missing Info - June',
      'Ineligible & Inactive - August',
    ]);
  });

  it('moves a row sitting on the wrong month, and never deletes before it appends', () => {
    // The same path a change of keyword takes. The month is part of the
    // destination, so nothing needs a second notion of "wrong tab".
    const plan = reconcile({
      daily: [
        parseDailySheet('July 30, 2026', dailyRange('July 30, 2026', [
          { status: 'missing info', camp: 'Achim', last: 'A', first: 'B', phone: '555', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info - June',
          queueRange('Missing Info - June', 'July 30, 2026', 'S000000000001'),
        ),
        parseQueueSheet('Missing Info - July', {
          address: 'Missing Info - July!A1:BA1',
          rowCount: 1,
          columnCount: WIDTH,
          values: [new Array(WIDTH).fill(null)],
        }),
      ],
      newSyncId: ids(),
    });

    const removals = plan.intents.filter((i) => i.kind === 'remove-queue-row');
    expect(removals).toHaveLength(1);
    expect(removals[0]).toMatchObject({
      queueSheet: 'Missing Info - June',
      reason: 'wrong-queue',
    });

    const appends = plan.intents.filter(
      (i): i is AppendQueueRowIntent => i.kind === 'append-queue-row',
    );
    expect(appends).toHaveLength(1);
    expect(appends[0]!.destination).toBe('Missing Info - July');
    expect(appends[0]!.syncId).toBe('S000000000001');
  });

  it('leaves a row alone once it is on the right month', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('July 30, 2026', dailyRange('July 30, 2026', [
          { status: 'missing info', camp: 'Achim', last: 'A', first: 'B', phone: '555', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info - July',
          queueRange('Missing Info - July', 'July 30, 2026', 'S000000000001'),
        ),
      ],
      newSyncId: ids(),
    });

    expect(plan.counts['append-queue-row']).toBe(0);
    expect(plan.counts['remove-queue-row']).toBe(0);
    expect(plan.counts['write-back']).toBe(0);
  });

  it('routes a daily sheet whose name is not a date to the Other tab', () => {
    // The scan already warns about these (`scan.unparseable_sheet_names`). The
    // point here is that their patients still reach a queue.
    const plan = reconcile({
      daily: [
        parseDailySheet('Overflow', dailyRange('Overflow', [
          { status: 'missing info', camp: 'Achim', last: 'A', first: 'B' },
        ])),
      ],
      queues: [],
      newSyncId: ids(),
    });

    const appends = plan.intents.filter(
      (i): i is AppendQueueRowIntent => i.kind === 'append-queue-row',
    );
    expect(appends[0]!.destination).toBe('Missing Info - Other');
  });

  it('shades a blank required field the same on every month of a queue', () => {
    const shaded = (sheet: QueueSheetName) =>
      planBlankRequiredShading(
        parseQueueSheet(sheet, {
          address: `${sheet}!A1:BA2`,
          rowCount: 2,
          columnCount: WIDTH,
          // Header row, then a row with a name and no phone number.
          values: [
            new Array(WIDTH).fill(null).map((_, i) => (i === 0 ? 'Date of Visit' : i === 1 ? 'Source Row' : null)),
            new Array(WIDTH).fill(null).map((_, i) => (i === 3 ? 'A' : null)),
          ],
        }),
      ).map((op) => op.what);

    expect(shaded('Missing Info - June')).toEqual(shaded('Missing Info - August'));
    expect(shaded('Missing Info - June')).toContain('blank required: Phone Number');
  });
});

describe('the migration keeps the office camp order', () => {
  const intent = (camp: string): AppendQueueRowIntent => ({
    kind: 'append-queue-row',
    destination: 'Missing Info - July',
    syncId: `S${camp.padEnd(12, '0').slice(0, 12).toUpperCase()}`,
    sourceSheet: 'July 30, 2026',
    sourceRow: 2,
    camp,
    values: { 'Last Name': camp },
    blankRequired: [],
  });

  const emptySheet = () =>
    parseQueueSheet('Missing Info - July', {
      address: 'Missing Info - July!A1:BA1',
      rowCount: 1,
      columnCount: WIDTH,
      values: [['Date of Visit', 'Source Row', 'Resolved', 'Last Name']],
    });

  const order = (newBlockOrder: 'alphabetical' | 'as-given') =>
    planQueueAppend({
      sheet: emptySheet(),
      appends: [intent('Zed'), intent('Achim'), intent('Melech')],
      newBlockOrder,
    }).placements.map((placement) => placement.camp);

  it('alphabetises by default, which is what a cycle wants', () => {
    expect(order('alphabetical')).toEqual(['Achim', 'Melech', 'Zed']);
  });

  it('keeps the given order for a migration, which is moving an ordered tab', () => {
    // Every camp is "new" to a freshly created monthly tab, so alphabetising
    // would silently reorder the whole sheet on a job whose premise is that
    // nothing changes but the month.
    expect(order('as-given')).toEqual(['Zed', 'Achim', 'Melech']);
  });
});
