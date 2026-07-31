import { describe, expect, it } from 'vitest';
import { queueColumnsFor, type QueueColumn, type QueueSheetName } from '../src/config';
import { parseDailySheet } from '../src/domain/dailySheets';
import { parseQueueSheet } from '../src/domain/queueSheets';
import { reconcile } from '../src/domain/reconcile';
import type { RangeData } from '../src/graph/workbook';

const W = 53;

function dailyRange(sheet: string, rows: any[]): RangeData {
  const grid: unknown[][] = [new Array(W).fill(null)];
  for (const r of rows) {
    const c = new Array(W).fill(null);
    c[1] = r.status ?? null;
    c[3] = r.last ?? null;
    c[4] = r.first ?? null;
    c[5] = r.dob ?? null;
    c[10] = r.zip ?? null;
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
    at('Zip Code', r.zip);
    at('Phone Number', r.phone);
    at('Insurance Carrier', r.carrier);
    c[52] = r.syncId ?? null;
    grid.push(c);
  }
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: W, values: grid };
}

const ids = () => () => 'S000000000999';

describe('ANALYSIS', () => {
  it('A: terminal source — stale mirror value is written over the completed daily row, then the row is deleted', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('July 20, 2026', dailyRange('July 20, 2026', [
          // The patient was processed: column B now says lasante, and the office
          // corrected the phone number on the daily sheet after the mirror froze.
          { status: 'lasante', last: 'Klein', first: 'M', dob: '11/30/2009', phone: '845-222-2222', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Ineligible & Inactive', queueRange('Ineligible & Inactive', [
          // Adopted mirror row, untouched by anybody since the formulas were frozen.
          { last: 'Klein', first: 'M', dob: 40147, phone: '845-111-1111', syncId: 'S000000000001' },
        ])),
      ],
      newSyncId: ids(),
    });

    console.log('A intents:', JSON.stringify(plan.intents, null, 1));
  });

  it('B: on-destination row nobody marked — write-back planned anyway', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('July 20, 2026', dailyRange('July 20, 2026', [
          { status: 'missing info', last: 'Klein', first: 'M', dob: '11/30/2009', carrier: 'Fidelis', phone: '845-222-2222', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Missing Info', queueRange('Missing Info', [
          { last: 'Klein', first: 'M', dob: 40147, carrier: 'Healthfirst', phone: '845-111-1111', syncId: 'S000000000001' },
        ])),
      ],
      newSyncId: ids(),
    });

    console.log('B intents:', JSON.stringify(plan.intents, null, 1));
    expect(plan.counts['write-back']).toBe(2);
  });

  it('C: a genuinely different zip is written back in its leading-zero-stripped form', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('July 20, 2026', dailyRange('July 20, 2026', [
          { status: 'missing info', last: 'Klein', first: 'M', dob: '11/30/2009', zip: '08701', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Missing Info', queueRange('Missing Info', [
          // Excel stored 08527 in a numeric cell, so it reads back as 8527.
          { last: 'Klein', first: 'M', dob: 40147, zip: 8527, syncId: 'S000000000001' },
        ])),
      ],
      newSyncId: ids(),
    });

    console.log('C intents:', JSON.stringify(plan.intents.filter((i) => i.kind === 'write-back'), null, 1));
  });

  it('D: wrong-queue row — write-back planned with no marker at all', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet('July 20, 2026', dailyRange('July 20, 2026', [
          { status: 'verify insurance', last: 'Klein', first: 'M', dob: '11/30/2009', phone: '845-222-2222', syncId: 'S000000000001' },
        ])),
      ],
      queues: [
        parseQueueSheet('Missing Info', queueRange('Missing Info', [
          { last: 'Klein', first: 'M', dob: 40147, phone: '845-111-1111', syncId: 'S000000000001' },
        ])),
      ],
      newSyncId: ids(),
    });

    console.log('D intents:', JSON.stringify(plan.intents, null, 1));
  });
});
