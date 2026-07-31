import { describe, expect, it } from 'vitest';
import { queueColumnsFor, type QueueColumn, type QueueSheetName } from '../src/config';
import { parseDailySheet } from '../src/domain/dailySheets';
import { parseQueueSheet } from '../src/domain/queueSheets';
import { reconcile } from '../src/domain/reconcile';
import type { RangeData } from '../src/graph/workbook';

/**
 * Each of these was a real defect found by reviewing the plan the reconciler
 * produces rather than the code that produces it. Every one of them ends in
 * patient data being lost or duplicated, and none of them raised an error.
 */

const W = 53;

function dailyRange(sheet: string, rows: any[]): RangeData {
  const grid: unknown[][] = [new Array(W).fill(null)];
  for (const r of rows) {
    const c = new Array(W).fill(null);
    c[1] = r.status ?? null;
    c[3] = r.last ?? null;
    c[4] = r.first ?? null;
    c[5] = r.dob ?? null;
    c[11] = r.phone ?? null;
    c[12] = r.carrier ?? null;
    c[52] = r.syncId ?? null;
    grid.push(c);
  }
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: W, values: grid };
}

function queueRange(sheet: QueueSheetName, rows: any[]): RangeData {
  const columns = queueColumnsFor(sheet);
  const header = new Array(W).fill(null);
  columns.forEach((c, i) => (header[i] = c));
  const grid: unknown[][] = [header];
  for (const r of rows) {
    const c = new Array(W).fill(null);
    const at = (col: QueueColumn, v: unknown) => {
      const i = columns.indexOf(col);
      if (i !== -1 && v !== undefined) c[i] = v;
    };
    at('Resolved', r.resolved);
    at('Last Name', r.last);
    at('First Name', r.first);
    at('Date of Birth', r.dob);
    at('Phone Number', r.phone);
    at('Insurance Carrier', r.carrier);
    c[52] = r.syncId ?? null;
    grid.push(c);
  }
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: W, values: grid };
}

const ids = () => {
  let n = 0;
  return () => `S${String(++n).padStart(12, '0')}`;
};

describe('a blank on the queue is not an instruction to erase the source', () => {
  it('reports the blank field instead of writing it back', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          'S',
          dailyRange('S', [
            { status: 'missing info', last: 'Smith', first: 'A', dob: 'x', phone: '', carrier: 'Fidelis', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          // The old mirror never copied Insurance Carrier onto this legacy row.
          queueRange('Missing Info', [
            { last: 'Smith', first: 'A', dob: 'x', phone: '555-1234', carrier: '', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: ids(),
    });

    const writeBacks = plan.intents.filter((i) => i.kind === 'write-back');
    // The phone the staff member supplied goes home.
    expect(writeBacks.map((i: any) => i.field)).toEqual(['Phone Number']);
    // The carrier the mirror never carried does NOT erase Fidelis. An empty
    // string clears a cell even though null does not, and the daily sheet is
    // the copy that gets billed from.
    expect(plan.blankedOnQueue).toEqual([
      { queueSheet: 'Missing Info', queueRow: 2, syncId: 'S000000000001', field: 'Insurance Carrier' },
    ]);
  });
});

describe('one physical row, one removal', () => {
  it('does not remove a row twice when it both moved queue and was blanked', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('S', dailyRange('S', [
          { status: 'verify insurance', last: 'A', first: 'B', dob: 'x', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Missing Info', queueRange('Missing Info', [{ syncId: 'S000000000001' }])),
      ],
      newSyncId: ids(),
    });

    const removals = plan.intents.filter((i) => i.kind === 'remove-queue-row');
    expect(removals).toHaveLength(1);
    expect(removals[0]).toMatchObject({ reason: 'wrong-queue' });
    // The second removal carried a source pointer that would have cleared column
    // B for a patient still legitimately queued on Verify Insurance.
    expect(removals.some((r: any) => r.reason === 'cleared-on-queue')).toBe(false);
  });
});

describe('work typed onto a row that is about to be removed', () => {
  it('carries the edit home before removing a wrong-queue row marked Done', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('S', dailyRange('S', [
          { status: 'verify insurance', last: 'A', first: 'B', dob: 'x', phone: '', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Missing Info', queueRange('Missing Info', [
          { resolved: 'Done', last: 'A', first: 'B', dob: 'x', phone: '555-9999', syncId: 'S000000000001' },
        ])),
      ],
      newSyncId: ids(),
    });

    const writeBacks = plan.intents.filter((i) => i.kind === 'write-back');
    expect(writeBacks).toHaveLength(1);
    expect(writeBacks[0]).toMatchObject({ field: 'Phone Number', value: '555-9999' });

    // And it is planned BEFORE the removal, because the queue row is the only
    // place that number exists until the write-back lands.
    const removalAt = plan.intents.findIndex((i) => i.kind === 'remove-queue-row');
    const writeBackAt = plan.intents.findIndex((i) => i.kind === 'write-back');
    expect(writeBackAt).toBeLessThan(removalAt);
  });
});

describe('a queue row with no SyncID is still a row', () => {
  it('does not append beside a patient who is already on that tab unstamped', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('S', dailyRange('S', [
          { status: 'missing info', last: 'A', first: 'B', dob: 'x', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        // Adoption could not link this one, so it carries no ID.
        parseQueueSheet('Missing Info', queueRange('Missing Info', [{ last: 'A', first: 'B', dob: 'x' }])),
      ],
      newSyncId: ids(),
    });

    expect(plan.counts['append-queue-row']).toBe(0);
    expect(plan.wouldDuplicate).toEqual([
      {
        queueSheet: 'Missing Info',
        queueRow: 2,
        syncId: 'S000000000001',
        reason: 'unstamped-row-for-this-patient',
      },
    ]);
  });

  it('still appends when the unstamped row is a different patient', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('S', dailyRange('S', [
          { status: 'missing info', last: 'A', first: 'B', dob: 'x', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Missing Info', queueRange('Missing Info', [{ last: 'Z', first: 'Y', dob: 'q' }])),
      ],
      newSyncId: ids(),
    });
    expect(plan.counts['append-queue-row']).toBe(1);
  });
});

describe('two daily rows carrying one SyncID', () => {
  it('reconciles neither and reports both, rather than letting scan order decide', () => {
    // Staff enter a repeat visit by copying the patient's row to the next day,
    // and column BA travels with the copy. 148 patients here have repeat visits.
    const plan = reconcile({
      daily: [
        parseDailySheet('July 30, 2026', dailyRange('July 30, 2026', [
          { status: 'missing info', last: 'A', first: 'B', dob: 'x', syncId: 'S000000000001' },
        ])),
        parseDailySheet('July 31, 2026', dailyRange('July 31, 2026', [
          { status: 'missing info', last: 'A', first: 'B', dob: 'x', syncId: 'S000000000001' },
        ])),
      ],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', []))],
      newSyncId: ids(),
    });

    expect(plan.counts['append-queue-row']).toBe(0);
    expect(plan.duplicateSources.map((d) => `${d.sheet}!${d.row}`).sort()).toEqual([
      'July 30, 2026!2',
      'July 31, 2026!2',
    ]);
  });
});
