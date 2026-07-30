import { describe, expect, it } from 'vitest';
import { planSheetStamps, stampWriteValues } from '../src/domain/backfill';
import type { RangeData } from '../src/graph/workbook';

/** A daily sheet grid: header row 1, data from row 2, B status, BA SyncID. */
function sheet(rows: { status?: string; syncId?: string }[], header = ''): RangeData {
  const width = 53;
  const grid: unknown[][] = [];

  const head = new Array(width).fill(null);
  head[1] = 'Status';
  head[52] = header || null;
  grid.push(head);

  for (const row of rows) {
    const cells = new Array(width).fill(null);
    cells[1] = row.status ?? null;
    cells[52] = row.syncId ?? null;
    grid.push(cells);
  }

  return {
    address: `'2026-07-30'!A1:BA${grid.length}`,
    rowCount: grid.length,
    columnCount: width,
    values: grid,
  };
}

let counter = 0;
const mint = () => `S${String(++counter).padStart(12, '0')}`;

describe('planSheetStamps', () => {
  it('stamps only rows carrying a queued keyword', () => {
    const plan = planSheetStamps(
      '2026-07-30',
      sheet([
        { status: 'missing info' },
        { status: 'ohi - esti' },
        { status: 'call back tuesday' },
        { status: '' },
        { status: 'not accepted' },
      ]),
      undefined,
      mint,
    );

    expect(plan.newStamps).toBe(2);
    expect([...plan.stamps.keys()].sort((a, b) => a - b)).toEqual([2, 6]);
  });

  it('does not backfill the workbook', () => {
    // A sheet full of real patients with no keyword gets nothing, however
    // complete their identifying fields are.
    const plan = planSheetStamps(
      '2026-07-30',
      sheet([{ status: '' }, { status: 'ohi' }, { status: 'lasante' }]),
      undefined,
      mint,
    );

    expect(plan.newStamps).toBe(0);
    expect(stampWriteValues(plan)).toEqual([]);
  });

  it('never re-stamps a row that already has an ID', () => {
    const plan = planSheetStamps(
      '2026-07-30',
      sheet([
        { status: 'missing info', syncId: 'S000000000042' },
        { status: 'not accepted' },
      ]),
      undefined,
      mint,
    );

    expect(plan.alreadyStamped).toBe(1);
    expect(plan.newStamps).toBe(1);
    // The existing row is outside the minted block, so it is not written at all.
    expect(plan.firstRow).toBe(3);
    expect(plan.lastRow).toBe(3);
  });

  it('preserves an existing ID that falls inside the write block', () => {
    // This is the unrecoverable mistake the planner exists to prevent: a single
    // range write from the first stamp to the last must not blank an ID in
    // between, because a queue row's note is anchored to it.
    const plan = planSheetStamps(
      '2026-07-30',
      sheet([
        { status: 'missing info' }, // row 2, minted
        { status: 'not accepted', syncId: 'S000000000042' }, // row 3, preserved
        { status: 'ohi', syncId: 'S000000000043' }, // row 4, terminal but stamped
        { status: 'verify insurance' }, // row 5, minted
      ]),
      undefined,
      mint,
    );

    const values = stampWriteValues(plan);
    expect(values).toHaveLength(4);
    expect(values[1]).toEqual(['S000000000042']);
    expect(values[2]).toEqual(['S000000000043']);
    expect(values[0]?.[0]).toMatch(/^S/);
    expect(values[3]?.[0]).toMatch(/^S/);
    expect(values.some((row) => row[0] === null)).toBe(false);
  });

  it('writes null only for rows in the block that genuinely have no ID', () => {
    const plan = planSheetStamps(
      '2026-07-30',
      sheet([
        { status: 'missing info' }, // row 2, minted
        { status: 'call back later' }, // row 3, unrecognized and unstamped
        { status: 'not accepted' }, // row 4, minted
      ]),
      undefined,
      mint,
    );

    expect(stampWriteValues(plan)[1]).toEqual([null]);
  });

  it('flags a missing SyncID header and clears the flag once present', () => {
    expect(planSheetStamps('x', sheet([{ status: 'missing info' }]), undefined, mint).headerNeeded).toBe(true);
    expect(
      planSheetStamps('x', sheet([{ status: 'missing info' }], 'SyncID'), undefined, mint)
        .headerNeeded,
    ).toBe(false);
  });

  it('is idempotent: a second pass over the stamped result mints nothing', () => {
    const first = planSheetStamps(
      '2026-07-30',
      sheet([{ status: 'missing info' }, { status: 'not accepted' }]),
      undefined,
      mint,
    );

    const stamped = sheet(
      [
        { status: 'missing info', syncId: first.stamps.get(2) },
        { status: 'not accepted', syncId: first.stamps.get(3) },
      ],
      'SyncID',
    );

    const second = planSheetStamps('2026-07-30', stamped, undefined, mint);
    expect(second.newStamps).toBe(0);
    expect(second.headerNeeded).toBe(false);
    expect(stampWriteValues(second)).toEqual([]);
  });
});
