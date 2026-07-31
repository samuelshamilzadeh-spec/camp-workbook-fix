import { describe, expect, it } from 'vitest';
import { queueColumnsFor, type QueueColumn, type QueueSheetName } from '../src/config';
import { parseDailySheet } from '../src/domain/dailySheets';
import { parseQueueSheet } from '../src/domain/queueSheets';
import { reconcile } from '../src/domain/reconcile';
import type { RangeData } from '../src/graph/workbook';

const WIDTH = 53;

function dailyRange(
  sheet: string,
  rows: { status?: string; camp?: string; last?: string; first?: string; dob?: string; phone?: string; carrier?: string; syncId?: string }[],
): RangeData {
  const grid: unknown[][] = [new Array(WIDTH).fill(null)];
  for (const row of rows) {
    const cells = new Array(WIDTH).fill(null);
    cells[1] = row.status ?? null;
    cells[2] = row.camp ?? null;
    cells[3] = row.last ?? null;
    cells[4] = row.first ?? null;
    cells[5] = row.dob ?? null;
    cells[11] = row.phone ?? null;
    cells[12] = row.carrier ?? null;
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
  columns.forEach((c, i) => { header[i] = c; });
  const grid: unknown[][] = [header];
  for (const row of rows) {
    const cells = new Array(WIDTH).fill(null);
    const at = (column: QueueColumn, value: unknown) => {
      const i = columns.indexOf(column);
      if (i !== -1 && value !== undefined) cells[i] = value;
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

describe('probe', () => {
  it('1: fully cleared queue row ON its destination queue', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'missing info', last: 'Smith', first: 'Aron', dob: '11/30/2009', phone: '555-1234', syncId: 'S000000000001' },
      ]))],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', [{ syncId: 'S000000000001' }]))],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE1', JSON.stringify(plan.intents, null, 1));
  });

  it('2: fully cleared STALE queue row (patient moved queues)', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'verify insurance', last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
      ]))],
      queues: [
        parseQueueSheet('Verify Insurance', queueRange('Verify Insurance', [
          { last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
        ])),
        parseQueueSheet('Missing Info', queueRange('Missing Info', [{ syncId: 'S000000000001' }])),
      ],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE2', JSON.stringify(plan.intents, null, 1));
  });

  it('3: Resolved marked on the WRONG (stale) queue', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'verify insurance', last: 'Smith', first: 'Aron', phone: '', syncId: 'S000000000001' },
      ]))],
      queues: [
        parseQueueSheet('Verify Insurance', queueRange('Verify Insurance', [
          { last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
        ])),
        parseQueueSheet('Missing Info', queueRange('Missing Info', [
          { resolved: 'Done', last: 'Smith', first: 'Aron', phone: '555-9999', syncId: 'S000000000001' },
        ])),
      ],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE3', JSON.stringify(plan.intents, null, 1));
    console.log('PROBE3-unrec', JSON.stringify(plan.unrecognizedResolved));
  });

  it('4: terminal source + blanked queue row', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'ohi', last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
      ]))],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', [{ syncId: 'S000000000001' }]))],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE4', JSON.stringify(plan.intents, null, 1));
  });

  it('5: two daily rows share a SyncID', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'missing info', last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
        { status: 'missing info', last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
      ]))],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', []))],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE5', JSON.stringify(plan.intents, null, 1), JSON.stringify(plan.orphans));
  });

  it('6: resolved row on an append-only destination (United Refuah has no Resolved col)', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'united refuah', last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
      ]))],
      queues: [parseQueueSheet('United Refuah', queueRange('United Refuah', [
        { last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
      ]))],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE6', JSON.stringify(plan.intents), JSON.stringify(plan.orphans));
  });

  it('7: duplicate on destination, second one marked Resolved', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'missing info', last: 'Smith', first: 'Aron', phone: '', syncId: 'S000000000001' },
      ]))],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', [
        { last: 'Smith', first: 'Aron', syncId: 'S000000000001' },
        { resolved: 'Done', last: 'Smith', first: 'Aron', phone: '555-1111', syncId: 'S000000000001' },
      ]))],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE7', JSON.stringify(plan.intents), JSON.stringify(plan.orphans));
  });

  it('8: blank-vs-blank / partial edits', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [
        { status: 'missing info', last: 'Smith', first: 'Aron', dob: '11/30/2009', phone: '555-0000', syncId: 'S000000000001' },
      ]))],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', [
        { last: 'Smith', first: 'Aron', dob: 40147, phone: '', syncId: 'S000000000001' },
      ]))],
      newSyncId: () => 'NEW',
    });
    console.log('PROBE8', JSON.stringify(plan.intents));
  });
});
