import { describe, expect, it } from 'vitest';
import { LAYOUT, type QueueSheetName } from '../src/config';
import { parseDailySheet } from '../src/domain/dailySheets';
import { parseQueueSheet } from '../src/domain/queueSheets';
import { reconcile, type AppendQueueRowIntent } from '../src/domain/reconcile';
import type { RangeData } from '../src/graph/workbook';

/**
 * Builds a daily sheet usedRange starting at A1. Columns follow the configured
 * layout: B status, C camp, D-M patient fields, BA SyncID.
 */
function dailyRange(
  sheet: string,
  rows: { status?: string; camp?: string; last?: string; first?: string; dob?: string; phone?: string; carrier?: string; syncId?: string }[],
): RangeData {
  const width = 53; // through BA
  const grid: unknown[][] = [];

  const header = new Array(width).fill(null);
  header[0] = 'Row';
  header[1] = 'Status';
  header[2] = 'Camp';
  header[52] = 'SyncID';
  grid.push(header);

  for (const row of rows) {
    const cells = new Array(width).fill(null);
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

  return {
    address: `${sheet}!A1:BA${grid.length}`,
    rowCount: grid.length,
    columnCount: width,
    values: grid,
  };
}

/** Builds a queue sheet usedRange: instructions rows 1-8, headers row 9, data row 10+. */
function queueRange(
  sheet: QueueSheetName,
  entries: ({ kind: 'header'; camp: string; count: number } | {
    kind: 'row';
    dateOfVisit?: string;
    notes?: string;
    last?: string;
    first?: string;
    dob?: string;
    phone?: string;
    carrier?: string;
    syncId?: string;
  })[],
): RangeData {
  const width = 53;
  const grid: unknown[][] = [];

  for (let i = 0; i < LAYOUT.queue.headerRow; i++) {
    grid.push(new Array(width).fill(null));
  }

  for (const entry of entries) {
    const cells = new Array(width).fill(null);
    if (entry.kind === 'header') {
      cells[0] = `${entry.camp} - ${entry.count} patients`;
    } else {
      cells[0] = entry.dateOfVisit ?? null;
      cells[2] = entry.notes ?? null;
      cells[3] = entry.last ?? null;
      cells[4] = entry.first ?? null;
      cells[5] = entry.dob ?? null;
      cells[11] = entry.phone ?? null;
      cells[12] = entry.carrier ?? null;
      cells[52] = entry.syncId ?? null;
    }
    grid.push(cells);
  }

  return {
    address: `${sheet}!A1:BA${grid.length}`,
    rowCount: grid.length,
    columnCount: width,
    values: grid,
  };
}

/** Deterministic ID minting, so a plan can be compared across runs. */
function sequentialIds(): () => string {
  let n = 0;
  return () => `S${String(++n).padStart(12, '0')}`;
}

describe('reconcile', () => {
  it('appends a queued row and stamps it, skipping terminal and unrecognized rows', () => {
    const daily = parseDailySheet(
      '2026-07-30',
      dailyRange('2026-07-30', [
        { status: 'missing info', camp: 'Camp Ramah', last: 'A', first: 'B', dob: '2010-01-01' },
        { status: 'ohi - esti', camp: 'Camp Ramah', last: 'C', first: 'D' },
        { status: 'call back tuesday', camp: 'Camp Ramah', last: 'E', first: 'F' },
        { status: '', camp: 'Camp Ramah' },
      ]),
    );

    const plan = reconcile({
      daily: [daily],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', []))],
      newSyncId: sequentialIds(),
    });

    expect(plan.counts['stamp-id']).toBe(1);
    expect(plan.counts['append-queue-row']).toBe(1);
    expect(plan.unrecognizedCounts['2026-07-30']).toBe(1);

    const append = plan.intents.find(
      (i): i is AppendQueueRowIntent => i.kind === 'append-queue-row',
    )!;
    expect(append.destination).toBe('Missing Info');
    expect(append.camp).toBe('Camp Ramah');
    expect(append.values['Date of Visit']).toBe('2026-07-30');
    expect(append.values['Source Row']).toBe('2026-07-30!B2');
    // Notes are never seeded from the source.
    expect(append.values['Notes']).toBe('');
    // Phone is blank and required for Missing Info, so it gets red shading.
    expect(append.blankRequired).toContain('Phone Number');
    expect(append.blankRequired).not.toContain('Last Name');
  });

  it('is idempotent: once applied, a second run plans nothing', () => {
    const stamped = { status: 'missing info', camp: 'Camp Ramah', last: 'A', first: 'B', dob: '2010-01-01', phone: '555', syncId: 'S000000000001' };

    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', [stamped]))],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            { kind: 'header', camp: 'Camp Ramah', count: 1 },
            {
              kind: 'row',
              dateOfVisit: '2026-07-30',
              last: 'A',
              first: 'B',
              dob: '2010-01-01',
              phone: '555',
              syncId: 'S000000000001',
            },
          ]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    expect(plan.intents).toHaveLength(0);
    expect(plan.orphans).toHaveLength(0);
  });

  it('does not re-append a row that is already queued, whatever the suffix', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'missing info - waiting on card', last: 'A', first: 'B', dob: 'x', phone: '1', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            { kind: 'row', dateOfVisit: '2026-07-30', last: 'A', first: 'B', dob: 'x', phone: '1', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    expect(plan.counts['append-queue-row']).toBe(0);
    expect(plan.counts['write-back']).toBe(0);
  });

  it('plans a write-back when staff edit a field on the queue sheet', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'missing info', last: 'Smith', first: 'A', dob: 'x', phone: '', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            {
              kind: 'row',
              dateOfVisit: '2026-07-30',
              last: 'Smith',
              first: 'A',
              dob: 'x',
              phone: '555-1234',
              notes: 'called mom, left voicemail',
              syncId: 'S000000000001',
            },
          ]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    const writeBacks = plan.intents.filter((i) => i.kind === 'write-back');
    expect(writeBacks).toHaveLength(1);
    expect(writeBacks[0]).toMatchObject({
      field: 'Phone Number',
      sourceSheet: '2026-07-30',
      sourceRow: 2,
    });
    // Notes stay on the queue sheet and are never written back.
    expect(plan.intents.some((i) => i.kind === 'write-back' && i.field === 'Notes')).toBe(false);
  });

  it('removes a queue row once the source reaches a terminal status', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'ohi', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            { kind: 'row', dateOfVisit: '2026-07-30', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    expect(plan.intents).toHaveLength(1);
    expect(plan.intents[0]).toMatchObject({
      kind: 'remove-queue-row',
      reason: 'no-longer-queued-at-source',
    });
  });

  it('moves a row when its keyword changes destination', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'not accepted', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            { kind: 'row', dateOfVisit: '2026-07-30', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
        parseQueueSheet('Not Accepted', queueRange('Not Accepted', [])),
      ],
      newSyncId: sequentialIds(),
    });

    expect(plan.counts['remove-queue-row']).toBe(1);
    expect(plan.counts['append-queue-row']).toBe(1);
    const append = plan.intents.find((i) => i.kind === 'append-queue-row');
    expect(append).toMatchObject({ destination: 'Not Accepted' });
    // No new ID: the row was already stamped and an ID is stamped once, never changed.
    expect(plan.counts['stamp-id']).toBe(0);
  });

  it('treats a fully blanked queue row as a clear-and-remove', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'missing info', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [{ kind: 'row', syncId: 'S000000000001' }]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    const removals = plan.intents.filter((i) => i.kind === 'remove-queue-row');
    expect(removals).toHaveLength(1);
    expect(removals[0]).toMatchObject({
      reason: 'cleared-on-queue',
      sourceSheet: '2026-07-30',
      sourceRow: 2,
    });
  });

  it('does not treat a row with one blank required field as cleared', () => {
    // A Missing Info row is missing something by definition. Deleting it would
    // drop exactly the patients this system exists to chase.
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'missing info', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            { kind: 'row', dateOfVisit: '2026-07-30', last: 'A', first: 'B', syncId: 'S000000000001' },
          ]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    expect(plan.counts['remove-queue-row']).toBe(0);
  });

  it('reports a queue row whose source is not in the scanned set rather than deleting it', () => {
    const plan = reconcile({
      daily: [parseDailySheet('2026-07-30', dailyRange('2026-07-30', []))],
      queues: [
        parseQueueSheet(
          'Missing Info',
          queueRange('Missing Info', [
            { kind: 'row', dateOfVisit: '2026-06-01', last: 'A', syncId: 'S000000000009' },
          ]),
        ),
      ],
      newSyncId: sequentialIds(),
    });

    expect(plan.intents).toHaveLength(0);
    expect(plan.orphans).toEqual([
      {
        queueSheet: 'Missing Info',
        queueRow: 10,
        syncId: 'S000000000009',
        reason: 'unknown-sync-id',
      },
    ]);
  });

  it('reports multi-match cells instead of resolving them silently', () => {
    const plan = reconcile({
      daily: [
        parseDailySheet(
          '2026-07-30',
          dailyRange('2026-07-30', [
            { status: 'missing info + verify insurance', last: 'A', syncId: 'S000000000001' },
          ]),
        ),
      ],
      queues: [parseQueueSheet('Missing Info', queueRange('Missing Info', []))],
      newSyncId: sequentialIds(),
    });

    expect(plan.ambiguous).toHaveLength(1);
    expect(plan.ambiguous[0]).toMatchObject({
      sheet: '2026-07-30',
      row: 2,
      destination: 'Missing Info',
    });
    // Flagged, but still routed — never dropped on the floor.
    expect(plan.counts['append-queue-row']).toBe(1);
  });

  it('produces a structurally identical plan when run twice on the same state', () => {
    const build = () =>
      reconcile({
        daily: [
          parseDailySheet(
            '2026-07-30',
            dailyRange('2026-07-30', [
              { status: 'missing info', camp: 'Camp A', last: 'A', first: 'B' },
              { status: 'not accepted', camp: 'Camp B', last: 'C', first: 'D' },
            ]),
          ),
        ],
        queues: [
          parseQueueSheet('Missing Info', queueRange('Missing Info', [])),
          parseQueueSheet('Not Accepted', queueRange('Not Accepted', [])),
        ],
        newSyncId: sequentialIds(),
      });

    expect(JSON.stringify(build())).toBe(JSON.stringify(build()));
  });
});
