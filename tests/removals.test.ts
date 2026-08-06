import { describe, expect, it, vi } from 'vitest';
import { applyRemovals, type RemovalWorkbook } from '../src/sync/removals';
import type { RemoveQueueRowIntent, WriteBackIntent } from '../src/domain/reconcile';

const log = { debug() {}, info() {}, warn: vi.fn(), error() {} } as any;
const tabFor = new Map([['Missing Info - July', 'Missing Info - July']]);

/**
 * A sheet that behaves like the workbook: deleting a row shifts everything below
 * it up, so a stale row number silently refers to a different patient.
 */
function fakeWorkbook(rowsBySyncId: Record<number, string>) {
  let rows = new Map<number, string>(
    Object.entries(rowsBySyncId).map(([r, id]) => [Number(r), id]),
  );
  const calls: string[] = [];

  const workbook: RemovalWorkbook = {
    async getRange(_sheet, address) {
      const row = Number(address.replace(/[^0-9]/g, ''));
      return { values: [[rows.get(row) ?? '']] };
    },
    async writeRange(sheet, address) {
      calls.push(`write ${sheet}!${address}`);
    },
    async deleteRows(sheet, firstRow, count) {
      calls.push(`delete ${sheet}!${firstRow}`);
      const next = new Map<number, string>();
      for (const [row, id] of rows) {
        if (row >= firstRow && row < firstRow + count) continue;
        next.set(row > firstRow ? row - count : row, id);
      }
      rows = next;
    },
    async clearRange(sheet, address) {
      calls.push(`clear ${sheet}!${address}`);
    },
  };
  return { workbook, calls, remaining: () => rows };
}

const removal = (row: number, syncId: string, over: Partial<RemoveQueueRowIntent> = {}): RemoveQueueRowIntent => ({
  kind: 'remove-queue-row',
  syncId,
  queueSheet: 'Missing Info - July',
  queueRow: row,
  sourceSheet: 'July 30, 2026',
  sourceRow: 5,
  reason: 'resolved-on-queue',
  ...over,
});

describe('applying removals', () => {
  it('deletes deepest first, so no delete moves another row out from under it', async () => {
    const { workbook, calls, remaining } = fakeWorkbook({
      10: 'S000000000001',
      11: 'S000000000002',
      12: 'S000000000003',
    });

    const result = await applyRemovals({
      workbook,
      removals: [removal(10, 'S000000000001'), removal(11, 'S000000000002'), removal(12, 'S000000000003')],
      writeBacks: [],
      tabFor,
      log,
    });

    expect(result.deleted).toBe(3);
    expect(result.skipped).toBe(0);
    expect(calls.filter((c) => c.startsWith('delete'))).toEqual([
      'delete Missing Info - July!12',
      'delete Missing Info - July!11',
      'delete Missing Info - July!10',
    ]);
    expect(remaining().size).toBe(0);
  });

  it('refuses a row whose SyncID has changed under it', async () => {
    // Somebody inserted a row above 10 after the plan was made, so row 10 is now
    // a different patient. Deleting it would remove the wrong person.
    const { workbook, calls, remaining } = fakeWorkbook({ 10: 'SOMEONEELSE1' });

    const result = await applyRemovals({
      workbook,
      removals: [removal(10, 'S000000000001')],
      writeBacks: [],
      tabFor,
      log,
    });

    expect(result).toMatchObject({ deleted: 0, skipped: 1, cleared: 0, wrote: 0 });
    expect(calls).toEqual([]);
    expect(remaining().get(10)).toBe('SOMEONEELSE1');
    expect(log.warn).toHaveBeenCalledWith('resolve.row_moved', expect.objectContaining({ row: 10 }));
  });

  it('writes the fix home before deleting the row that carries it', async () => {
    const { workbook, calls } = fakeWorkbook({ 10: 'S000000000001' });
    const writeBack: WriteBackIntent = {
      kind: 'write-back',
      syncId: 'S000000000001',
      sourceSheet: 'July 30, 2026',
      sourceRow: 5,
      queueSheet: 'Missing Info - July',
      queueRow: 10,
      field: 'Phone Number',
      value: '845-555-1234',
    };

    await applyRemovals({ workbook, removals: [removal(10, 'S000000000001')], writeBacks: [writeBack], tabFor, log });

    // The queue row is the only place that number exists until the write lands.
    expect(calls).toEqual([
      'write July 30, 2026!L5',
      'delete Missing Info - July!10',
      'clear July 30, 2026!B5',
    ]);
  });

  it('deletes before clearing, because only that order fails safely', async () => {
    const { workbook, calls } = fakeWorkbook({ 10: 'S000000000001' });
    await applyRemovals({ workbook, removals: [removal(10, 'S000000000001')], writeBacks: [], tabFor, log });

    // Clear-then-delete would strand the row: a blank column B makes the source
    // invisible to the reconciler, so nothing could ever remove the queue row.
    expect(calls.indexOf('delete Missing Info - July!10')).toBeLessThan(
      calls.indexOf('clear July 30, 2026!B5'),
    );
  });

  it('does not clear the status for a row that merely moved queue', async () => {
    const { workbook, calls } = fakeWorkbook({ 10: 'S000000000001' });
    const result = await applyRemovals({
      workbook,
      removals: [removal(10, 'S000000000001', { reason: 'wrong-queue' })],
      writeBacks: [],
      tabFor,
      log,
    });

    // Its column B already says the right thing. Blanking it would pull the
    // patient out of the queue they were just moved into.
    expect(result.cleared).toBe(0);
    expect(calls).toEqual(['delete Missing Info - July!10']);
  });

  it('carries on past a row that moved, and still gets the others right', async () => {
    const { workbook, remaining } = fakeWorkbook({
      10: 'S000000000001',
      11: 'MOVEDAWAY111',
      12: 'S000000000003',
    });

    const result = await applyRemovals({
      workbook,
      removals: [removal(10, 'S000000000001'), removal(11, 'S000000000002'), removal(12, 'S000000000003')],
      writeBacks: [],
      tabFor,
      log,
    });

    expect(result).toMatchObject({ deleted: 2, skipped: 1 });
    expect([...remaining().values()]).toEqual(['MOVEDAWAY111']);
  });
});
