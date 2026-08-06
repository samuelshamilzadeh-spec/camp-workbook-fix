import { describe, expect, it } from 'vitest';
import {
  LAYOUT,
  RESOLVED_DROPDOWN_VALUE,
  RESOLVED_VALUES,
  queueColumnsFor,
  type QueueColumn,
  type QueueSheetName,
} from '../src/config';
import { sameDay, toDateCellValue, toExcelSerial } from '../src/domain/dates';
import {
  assertSyncIdColumnIsClear,
  parseQueueSheet,
  readResolvedSignal,
} from '../src/domain/queueSheets';
import { parseDailySheet } from '../src/domain/dailySheets';
import { reconcile, sameFieldValue } from '../src/domain/reconcile';
import { planDividerStyle, planQueueStyle } from '../src/domain/style';
import type { RangeData } from '../src/graph/workbook';

const WIDTH = 53;

function dailyRange(
  sheet: string,
  rows: { status?: string; last?: string; first?: string; dob?: string; phone?: string; syncId?: string }[],
): RangeData {
  const grid: unknown[][] = [new Array(WIDTH).fill(null)];
  for (const row of rows) {
    const cells = new Array(WIDTH).fill(null);
    cells[1] = row.status ?? null;
    cells[3] = row.last ?? null;
    cells[4] = row.first ?? null;
    cells[5] = row.dob ?? null;
    cells[11] = row.phone ?? null;
    cells[52] = row.syncId ?? null;
    grid.push(cells);
  }
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: WIDTH, values: grid };
}

function queueRange(
  sheet: QueueSheetName,
  rows: { resolved?: string; last?: string; first?: string; dob?: unknown; phone?: string; syncId?: string }[],
): RangeData {
  const columns = queueColumnsFor(sheet);
  const header = new Array(WIDTH).fill(null);
  columns.forEach((column, index) => {
    header[index] = column;
  });
  const grid: unknown[][] = [header];

  for (const row of rows) {
    const cells = new Array(WIDTH).fill(null);
    const at = (column: QueueColumn, value: unknown) => {
      const index = columns.indexOf(column);
      if (index !== -1 && value !== undefined) cells[index] = value;
    };
    at('Resolved', row.resolved);
    at('Last Name', row.last);
    at('First Name', row.first);
    at('Date of Birth', row.dob);
    at('Phone Number', row.phone);
    cells[52] = row.syncId ?? null;
    grid.push(cells);
  }

  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: WIDTH, values: grid };
}

describe('the Resolved column', () => {
  it('is present on the work queues and absent from the append-only record', () => {
    expect(queueColumnsFor('Missing Info - July')).toContain('Resolved');
    expect(queueColumnsFor('United Refuah')).not.toContain('Resolved');
    // And it sits between Source Row and Last Name, which is what shifts the
    // rest of the sheet right.
    expect(queueColumnsFor('Missing Info - July').slice(0, 4)).toEqual([
      'Date of Visit',
      'Source Row',
      'Resolved',
      'Last Name',
    ]);
  });

  it('accepts the dropdown value and the hand-typed variants', () => {
    expect(RESOLVED_VALUES).toContain(RESOLVED_DROPDOWN_VALUE.toLowerCase());
    for (const value of ['Done', 'DONE', ' done ', 'yes', 'x', '✓', 'Fixed']) {
      expect(readResolvedSignal(value).kind).toBe('resolved');
    }
  });

  it('reports anything else rather than acting on it', () => {
    for (const value of ['maybe', 'called mom', '?', 'waiting on card']) {
      expect(readResolvedSignal(value)).toMatchObject({ kind: 'unrecognized', raw: value });
    }
    expect(readResolvedSignal('').kind).toBe('none');
    expect(readResolvedSignal(null).kind).toBe('none');
  });
});

describe('a row marked resolved', () => {
  const daily = () =>
    parseDailySheet(
      '2026-07-30',
      dailyRange('2026-07-30', [
        { status: 'missing info', last: 'Smith', first: 'A', dob: '11/30/2009', phone: '', syncId: 'S000000000001' },
      ]),
    );

  it('is removed from the queue and carries the pointer needed to clear column B', () => {
    const plan = reconcile({
      daily: [daily()],
      queues: [
        parseQueueSheet(
          'Missing Info - July',
          queueRange('Missing Info - July', [
            { resolved: 'Done', last: 'Smith', first: 'A', dob: '11/30/2009', phone: '555-1234', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: () => 'S000000000999',
    });

    const removals = plan.intents.filter((i) => i.kind === 'remove-queue-row');
    expect(removals).toHaveLength(1);
    expect(removals[0]).toMatchObject({
      reason: 'resolved-on-queue',
      queueSheet: 'Missing Info - July',
      // Without these the applier cannot clear the status at source, and the
      // patient is re-appended on the next cycle.
      sourceSheet: '2026-07-30',
      sourceRow: 2,
    });

    // And the phone the staff member typed goes back to the daily sheet. It only
    // exists on the queue row until this is applied, which is why the applier
    // writes back before it deletes.
    const writeBacks = plan.intents.filter((i) => i.kind === 'write-back');
    expect(writeBacks).toHaveLength(1);
    expect(writeBacks[0]).toMatchObject({ field: 'Phone Number', sourceSheet: '2026-07-30' });
  });

  it('is left alone when the marker is not a word we accept', () => {
    const plan = reconcile({
      daily: [daily()],
      queues: [
        parseQueueSheet(
          'Missing Info - July',
          queueRange('Missing Info - July', [
            { resolved: 'called mom', last: 'Smith', first: 'A', dob: '11/30/2009', phone: '555', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: () => 'S000000000999',
    });

    expect(plan.counts['remove-queue-row']).toBe(0);
    expect(plan.unrecognizedResolved).toEqual([
      {
        queueSheet: 'Missing Info - July',
        queueRow: 2,
        syncId: 'S000000000001',
        raw: 'called mom',
      },
    ]);
  });

  it('never writes the marker back to the daily sheet', () => {
    const plan = reconcile({
      daily: [daily()],
      queues: [
        parseQueueSheet(
          'Missing Info - July',
          queueRange('Missing Info - July', [
            { resolved: 'Done', last: 'Smith', first: 'A', dob: '11/30/2009', phone: '', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: () => 'S000000000999',
    });

    expect(plan.intents.some((i) => i.kind === 'write-back' && i.field === 'Resolved')).toBe(false);
  });
});

describe('dates', () => {
  it('converts a date to an Excel serial and back to the same day', () => {
    // 40147 is 30 November 2009, the value the live queue tabs already hold.
    expect(toExcelSerial('11/30/2009')).toBe(40147);
    expect(toExcelSerial('2009-11-30')).toBe(40147);
    expect(toExcelSerial(40147)).toBe(40147);
  });

  it('leaves a value it does not understand exactly as it found it', () => {
    // Dates of birth on the daily sheets are free text and some are not dates.
    // Writing a guess over one destroys the only record of what staff typed.
    for (const value of ['unknown', '', 'ask mom', 'n/a']) {
      expect(toDateCellValue(value)).toBe(value);
    }
  });

  it('treats a serial and the text it came from as the same day', () => {
    expect(sameDay(40147, '11/30/2009')).toBe(true);
    expect(sameDay(40147, '11/29/2009')).toBe(false);
  });

  it('does not plan a write-back just because the two sides store a date differently', () => {
    // The whole point: the queue tabs now hold serials so the column can be
    // sorted, and the daily sheets hold text. Compared as strings this would be
    // a write-back on every cycle, for every dated row, forever.
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'missing info', last: 'Smith', first: 'A', dob: '11/30/2009', phone: '555', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info - July',
          queueRange('Missing Info - July', [
            { last: 'Smith', first: 'A', dob: 40147, phone: '555', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: () => 'S000000000999',
    });

    expect(plan.counts['write-back']).toBe(0);
  });

  it('still compares non-date columns as text, so two zip codes stay two values', () => {
    // Applied to every column, the date-aware path would read `10952` as a
    // serial and start calling different values equal.
    expect(sameFieldValue('Zip Code', '10952', '10953')).toBe(false);
    expect(sameFieldValue('Date of Birth', 40147, '11/30/2009')).toBe(true);
  });
});

describe('the styling pass', () => {
  const styled = (sheet: QueueSheetName) =>
    planQueueStyle({
      sheet: parseQueueSheet(
        sheet,
        queueRange(sheet, [{ last: 'Smith', syncId: 'S000000000001' }]),
      ),
    });

  it('covers the full width of the destination, and no wider', () => {
    // 18 columns on a work queue, 17 on United Refuah.
    expect(styled('Missing Info - July').operations[0]!.address).toBe('A1:R1');
    expect(styled('United Refuah').operations[0]!.address).toBe('A1:Q1');
  });

  it('plans no operation Graph will refuse', () => {
    // The Resolved dropdown and the frozen header row belong here and are
    // deliberately absent: Graph 400s on `dataValidation` (even on a GET) and on
    // every spelling of `freezePanes` against this workbook. Both are manual
    // one-offs rather than calls that fail on every run. See src/domain/style.ts.
    const kinds = new Set(styled('Missing Info - July').operations.map((op) => op.kind));
    expect([...kinds].sort()).toEqual([
      'border',
      'fill',
      'font',
      'format',
      'number-format',
    ]);
  });

  it('formats both date columns and centres the Source Row', () => {
    const operations = styled('Missing Info - July').operations;
    const formats = operations.filter((op) => op.kind === 'number-format');
    expect(formats.map((op) => op.numberFormat?.format)).toEqual(['mm/dd/yyyy', 'mm/dd/yyyy']);
    expect(
      operations.some(
        (op) => op.what === 'centre Source Row' && op.format?.horizontalAlignment === 'Center',
      ),
    ).toBe(true);
  });

  it('does not draw vertical lines', () => {
    const edges = styled('Missing Info - July')
      .operations.filter((op) => op.kind === 'border')
      .map((op) => op.border?.edge);
    expect(edges).not.toContain('InsideVertical');
    expect(edges).toContain('InsideHorizontal');
  });
});

describe('the SyncID column still clears the data', () => {
  it('passes at the current width and fails once the data would reach it', () => {
    // Checked against the WIDEST layout, so adding a column cannot quietly grow
    // the data past the SyncID column on one tab while the check passes on
    // another.
    expect(() => assertSyncIdColumnIsClear(LAYOUT)).not.toThrow();
    expect(() =>
      assertSyncIdColumnIsClear({
        ...LAYOUT,
        queue: { ...LAYOUT.queue, syncIdColumn: 'R' },
      }),
    ).toThrow(/overlaps the data columns/);
  });
});

describe('a date that is not a day on the calendar', () => {
  it('is left exactly as staff typed it', () => {
    // `14/24/2014` is a typo — there is no fourteenth month. JavaScript's Date
    // rolls it forward to 24 February 2015 without a word, so converting the
    // daily sheets' text dates to real dates turned a mistake somebody could
    // spot into a plausible, wrong, unremarkable date of birth. It reached the
    // live workbook once, on `Not Accepted ` row 377.
    expect(toExcelSerial('14/24/2014')).toBeUndefined();
    expect(toDateCellValue('14/24/2014')).toBe('14/24/2014');

    expect(toExcelSerial('13/01/2020')).toBeUndefined();
    expect(toExcelSerial('02/30/2019')).toBeUndefined();
  });

  it('still accepts the awkward dates that are real', () => {
    expect(toExcelSerial('02/29/2020')).toBe(43890); // a leap day
    expect(toExcelSerial('12/31/1999')).toBeDefined();
  });

  it('does not treat two different bad dates as the same patient', () => {
    // Both roll to a real date under the old behaviour, and could have collided.
    expect(sameDay('14/24/2014', '02/24/2015')).toBe(false);
  });
});

describe('a camp appearing for the first time', () => {
  it('gets the same band every other divider has', () => {
    // A cycle writes a new divider's TEXT but nothing was dressing it — only
    // `npm run migrate` did. So a new camp turned up as a bare label sitting
    // between banded ones, which is the moment a queue looks broken to somebody
    // reading it.
    const sheet = parseQueueSheet(
      'Missing Info - July',
      queueRange('Missing Info - July', [{ last: 'A', syncId: 'S000000000001' }]),
    );
    // Pretend row 2 is the divider for a brand new camp.
    const withGroup = { ...sheet, groups: [{ camp: 'Gevurah', headerRow: 2, declaredCount: 1, rows: [] }] };

    const ops = planDividerStyle(withGroup as any, ['Gevurah']);
    expect(ops.some((o) => o.kind === 'fill' && o.address === 'A2:R2')).toBe(true);
    expect(ops.some((o) => o.kind === 'font' && o.font?.bold === true)).toBe(true);
  });

  it('styles only the camps named, not every divider on the tab', () => {
    // Restyling all 35 camps on Not Accepted is a hundred Graph calls, which
    // does not belong in a five-second cycle.
    const sheet = parseQueueSheet(
      'Missing Info - July',
      queueRange('Missing Info - July', [{ last: 'A', syncId: 'S000000000001' }]),
    );
    const withGroups = {
      ...sheet,
      groups: [
        { camp: 'Achim', headerRow: 2, declaredCount: 1, rows: [] },
        { camp: 'Gevurah', headerRow: 5, declaredCount: 1, rows: [] },
      ],
    };

    const ops = planDividerStyle(withGroups as any, ['Gevurah']);
    expect(ops.every((o) => !o.address.startsWith('A2'))).toBe(true);
    expect(ops.some((o) => o.address === 'A5:R5')).toBe(true);
  });
});
