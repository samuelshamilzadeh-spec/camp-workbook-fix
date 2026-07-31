import { describe, expect, it } from 'vitest';
import { LAYOUT, QUEUE_COLUMNS, type QueueColumn, type QueueSheetName } from '../src/config';
import {
  NO_CAMP_LABEL,
  planQueueAppend,
  type InsertRowsOperation,
  type QueueOperation,
  type WriteCellsOperation,
} from '../src/domain/append';
import { columnToIndex, parseAddress } from '../src/domain/cells';
import { detectQueueShape, parseQueueSheet } from '../src/domain/queueSheets';
import type { AppendQueueRowIntent } from '../src/domain/reconcile';
import type { RangeData } from '../src/graph/workbook';

const WIDTH = 53; // through BA
const SYNC_ID_INDEX = 52;

type Entry =
  | { kind: 'blank' }
  | { kind: 'header' }
  | { kind: 'divider'; camp: string; count: number }
  | { kind: 'total'; count: number }
  | { kind: 'row'; last?: string; syncId?: string };

/**
 * Builds a queue tab's usedRange from row 1. A `header` entry writes real
 * column labels, which is what `detectQueueShape` looks for.
 */
function queueRange(sheet: QueueSheetName, entries: Entry[]): RangeData {
  const grid: unknown[][] = entries.map((entry) => {
    const cells: unknown[] = new Array(WIDTH).fill(null);
    switch (entry.kind) {
      case 'header':
        QUEUE_COLUMNS.forEach((column, index) => {
          cells[index] = column;
        });
        break;
      case 'divider':
        cells[0] = `${entry.camp} - ${entry.count} patients`;
        break;
      case 'total':
        cells[0] = `TOTAL - ${entry.count} patients`;
        break;
      case 'row':
        cells[2] = entry.last ?? 'Surname';
        cells[SYNC_ID_INDEX] = entry.syncId ?? null;
        break;
      case 'blank':
        break;
    }
    return cells;
  });

  return {
    address: `${sheet}!A1:BA${grid.length}`,
    rowCount: grid.length,
    columnCount: WIDTH,
    values: grid,
  };
}

let nextId = 0;
function intent(
  destination: QueueSheetName,
  camp: string | undefined,
  overrides: Partial<AppendQueueRowIntent> = {},
): AppendQueueRowIntent {
  nextId++;
  return {
    kind: 'append-queue-row',
    destination,
    syncId: `S${String(nextId).padStart(12, '0')}`,
    sourceSheet: 'July 30, 2026',
    sourceRow: nextId + 1,
    camp,
    values: {
      'Date of Visit': 'July 30, 2026',
      'Source Row': nextId + 1,
      'Last Name': 'Surname',
      'First Name': 'Given',
    },
    blankRequired: [],
    ...overrides,
  };
}

const inserts = (operations: QueueOperation[]): InsertRowsOperation[] =>
  operations.filter((op): op is InsertRowsOperation => op.kind === 'insert-rows');

const writes = (operations: QueueOperation[], purpose: WriteCellsOperation['purpose']) =>
  operations.filter(
    (op): op is WriteCellsOperation => op.kind === 'write-cells' && op.purpose === purpose,
  );

describe('detectQueueShape', () => {
  it('finds a header on row 1, as the tabs this project created have', () => {
    const shape = detectQueueShape(queueRange('United Refuah', [{ kind: 'header' }]));
    expect(shape).toMatchObject({ headerRow: 1, firstDataRow: 2, detected: true });
  });

  it('finds a header below a block of anything else', () => {
    const shape = detectQueueShape(
      queueRange('Missing Info', [
        ...Array.from({ length: 8 }, () => ({ kind: 'blank' }) as const),
        { kind: 'header' },
        { kind: 'row' },
      ]),
    );
    expect(shape).toMatchObject({ headerRow: 9, firstDataRow: 10, detected: true });
  });

  it('falls back to the configured layout when no header is recognizable', () => {
    const shape = detectQueueShape(queueRange('Missing Info', [{ kind: 'blank' }]));
    expect(shape).toMatchObject({
      headerRow: LAYOUT.queue.headerRow,
      firstDataRow: LAYOUT.queue.firstDataRow,
      detected: false,
    });
  });
});

describe('planQueueAppend on an empty tab', () => {
  /** United Refuah as it stands: a header row and nothing else. */
  const empty = () => parseQueueSheet('United Refuah', queueRange('United Refuah', [{ kind: 'header' }]));

  it('writes straight below the header and never opens a gap', () => {
    const plan = planQueueAppend({
      sheet: empty(),
      appends: [
        intent('United Refuah', 'Achim'),
        intent('United Refuah', 'Achim'),
        intent('United Refuah', 'Bnos Naale'),
      ],
    });

    expect(plan.requiresInserts).toBe(false);
    expect(inserts(plan.operations)).toHaveLength(0);
    expect(plan.appended).toBe(3);
    // Two dividers plus three patients.
    expect(plan.rowsAdded).toBe(5);

    const body = writes(plan.operations, 'rows');
    expect(body).toHaveLength(1);
    expect(body[0]!.address).toBe('A2:Q6');
  });

  it('orders camps alphabetically with a divider carrying the count', () => {
    const plan = planQueueAppend({
      sheet: empty(),
      appends: [
        intent('United Refuah', 'Zebulun'),
        intent('United Refuah', 'Achim'),
        intent('United Refuah', 'Achim'),
      ],
    });

    // A divider carries its label in the first column and nothing else; a
    // patient row carries Date of Visit there and a Source Row beside it.
    const grid = writes(plan.operations, 'rows')[0]!.values;
    expect(grid.map((row) => (row[1] === null ? row[0] : 'patient'))).toEqual([
      'Achim - 2 patients',
      'patient',
      'patient',
      'Zebulun - 1 patient',
      'patient',
    ]);
  });

  it('puts the SyncIDs in BA, blank on the divider rows', () => {
    const appends = [intent('United Refuah', 'Achim'), intent('United Refuah', 'Achim')];
    const plan = planQueueAppend({ sheet: empty(), appends });

    const ids = writes(plan.operations, 'sync-ids')[0]!;
    expect(ids.address).toBe('BA2:BA4');
    expect(ids.values).toEqual([[null], [appends[0]!.syncId], [appends[1]!.syncId]]);
  });

  it('closes with a TOTAL row below everything', () => {
    const plan = planQueueAppend({
      sheet: empty(),
      appends: [intent('United Refuah', 'Achim'), intent('United Refuah', 'Bnos Naale')],
    });

    const total = writes(plan.operations, 'total');
    expect(total).toHaveLength(1);
    // Rows 2-5 hold two dividers and two patients, so the total lands on 6.
    expect(total[0]!.address).toBe('A6:A6');
    expect(total[0]!.values).toEqual([['TOTAL - 2 patients']]);
    // And it is written last, when nothing further can shift it.
    expect(plan.operations[plan.operations.length - 1]).toBe(total[0]);
  });

  it('groups patients with no camp under one block rather than merging them', () => {
    const plan = planQueueAppend({
      sheet: empty(),
      appends: [intent('United Refuah', undefined), intent('United Refuah', 'Achim')],
    });

    const labels = writes(plan.operations, 'rows')[0]!.values.map((row) => row[0]);
    expect(labels).toContain(`${NO_CAMP_LABEL} - 1 patient`);
  });

  it('writes a block the next run parses back into the same camps', () => {
    const first = planQueueAppend({
      sheet: empty(),
      appends: [intent('United Refuah', undefined), intent('United Refuah', 'Achim')],
    });

    // Feed the written block back in as though it were read from the workbook.
    const grid = writes(first.operations, 'rows')[0]!.values;
    const syncIds = writes(first.operations, 'sync-ids')[0]!.values;
    const entries: Entry[] = [{ kind: 'header' }];
    grid.forEach((row, index) => {
      const label = row[0];
      entries.push(
        typeof label === 'string' && label.includes(' - ')
          ? { kind: 'divider', camp: label.split(' - ')[0]!, count: 1 }
          : { kind: 'row', syncId: String(syncIds[index]![0]) },
      );
    });

    const reparsed = parseQueueSheet('United Refuah', queueRange('United Refuah', entries));
    expect(reparsed.groups.map((group) => group.camp).sort()).toEqual([
      NO_CAMP_LABEL,
      'Achim',
    ]);
    // Nothing left to do, so nothing is planned.
    expect(planQueueAppend({ sheet: reparsed, appends: [] }).operations).toEqual([]);
  });
});

describe('planQueueAppend on a populated tab', () => {
  /**
   * Two camps and a TOTAL, in the mirror-tab shape: header on row 9, data from
   * row 10.
   *
   *   10  Achim - 2 patients      13  Bnos Naale - 1 patient
   *   11  patient                 14  patient
   *   12  patient                 15  TOTAL - 3 patients
   */
  const populated = () =>
    parseQueueSheet(
      'Missing Info',
      queueRange('Missing Info', [
        ...Array.from({ length: 8 }, () => ({ kind: 'blank' }) as const),
        { kind: 'header' },
        { kind: 'divider', camp: 'Achim', count: 2 },
        { kind: 'row', syncId: 'S000000000901' },
        { kind: 'row', syncId: 'S000000000902' },
        { kind: 'divider', camp: 'Bnos Naale', count: 1 },
        { kind: 'row', syncId: 'S000000000903' },
        { kind: 'total', count: 3 },
      ]),
    );

  it('reads the existing shape, groups and total row', () => {
    const sheet = populated();
    expect(sheet.firstDataRow).toBe(10);
    expect(sheet.totalRow).toBe(15);
    expect(sheet.rows).toHaveLength(3);
    expect(sheet.groups.map((group) => [group.camp, group.headerRow])).toEqual([
      ['Achim', 10],
      ['Bnos Naale', 13],
    ]);
  });

  it('inserts into an existing camp block and updates its divider', () => {
    const plan = planQueueAppend({
      sheet: populated(),
      appends: [intent('Missing Info', 'Achim')],
    });

    expect(inserts(plan.operations)).toEqual([
      expect.objectContaining({ kind: 'insert-rows', row: 13, count: 1 }),
    ]);
    expect(writes(plan.operations, 'rows')[0]!.address).toBe('A13:Q13');
    expect(writes(plan.operations, 'divider')[0]).toMatchObject({
      address: 'A10:A10',
      values: [['Achim - 3 patients']],
    });
    expect(writes(plan.operations, 'total')[0]).toMatchObject({
      address: 'A15:A15',
      values: [['TOTAL - 4 patients']],
    });
  });

  it('emits inserts bottom-up so no row number goes stale', () => {
    const plan = planQueueAppend({
      sheet: populated(),
      appends: [
        intent('Missing Info', 'Achim'),
        intent('Missing Info', 'Bnos Naale'),
        intent('Missing Info', 'Chayeinu'),
      ],
    });

    const rows = inserts(plan.operations).map((op) => op.row);
    expect(rows).toEqual([...rows].sort((a, b) => b - a));
    // New camp at the append point, Bnos Naale above it, Achim above that.
    expect(rows).toEqual([15, 15, 13]);
  });

  it('does not let a new camp and the last existing camp collide on one row', () => {
    const plan = planQueueAppend({
      sheet: populated(),
      appends: [intent('Missing Info', 'Chayeinu'), intent('Missing Info', 'Bnos Naale')],
    });

    // Both want row 15. The new camp is written there first, then Bnos Naale
    // inserts at 15 and pushes it down — so the two blocks never overlap.
    const bodyWrites = writes(plan.operations, 'rows');
    expect(bodyWrites.map((write) => write.address)).toEqual(['A15:Q16', 'A15:Q15']);
    expect(inserts(plan.operations).map((op) => op.row)).toEqual([15, 15]);
  });

  it('rewrites the existing TOTAL before anything shifts it', () => {
    const plan = planQueueAppend({
      sheet: populated(),
      appends: [intent('Missing Info', 'Achim')],
    });
    expect(plan.operations[0]).toMatchObject({ purpose: 'total', row: 15 });
  });

  it('plans nothing when there is nothing to append', () => {
    const plan = planQueueAppend({ sheet: populated(), appends: [] });
    expect(plan.operations).toEqual([]);
    expect(plan.totalAfter).toBe(3);
  });

  it('ignores intents bound for another queue', () => {
    const plan = planQueueAppend({
      sheet: populated(),
      appends: [intent('Not Accepted', 'Achim')],
    });
    expect(plan.operations).toEqual([]);
  });
});

/**
 * Applies a plan to an in-memory sheet exactly as `scripts/append-queue.ts`
 * applies it to the workbook: operations in order, an insert shifting every row
 * below it down.
 *
 * This is the check that matters. The planner's whole correctness argument is
 * that addresses computed against the sheet as it was READ stay valid while the
 * sheet is being changed underneath them, and the only way to test that claim is
 * to actually change the sheet underneath them.
 */
function applyPlan(sheet: QueueSheetName, before: RangeData, operations: QueueOperation[]): RangeData {
  const grid = before.values.map((row) => [...row]);
  const blank = () => new Array(WIDTH).fill(null);
  const at = (row: number) => {
    while (grid.length < row) grid.push(blank());
    return grid[row - 1]!;
  };

  for (const operation of operations) {
    if (operation.kind === 'shade') continue;

    if (operation.kind === 'insert-rows') {
      grid.splice(operation.row - 1, 0, ...Array.from({ length: operation.count }, blank));
      continue;
    }

    const { startColumn, startRow } = parseAddress(operation.address);
    const firstColumn = columnToIndex(startColumn);
    operation.values.forEach((values, rowOffset) => {
      const row = at(startRow + rowOffset);
      values.forEach((value, columnOffset) => {
        row[firstColumn + columnOffset] = value;
      });
    });
  }

  return {
    address: `${sheet}!A1:BA${grid.length}`,
    rowCount: grid.length,
    columnCount: WIDTH,
    values: grid,
  };
}

describe('applying a plan to a sheet', () => {
  it('lands every row where the plan said, on a tab built from empty', () => {
    const before = queueRange('United Refuah', [{ kind: 'header' }]);
    const appends = [
      intent('United Refuah', 'Zebulun'),
      intent('United Refuah', 'Achim'),
      intent('United Refuah', 'Achim'),
    ];
    const plan = planQueueAppend({
      sheet: parseQueueSheet('United Refuah', before),
      appends,
    });

    const after = parseQueueSheet('United Refuah', applyPlan('United Refuah', before, plan.operations));

    expect(after.rows).toHaveLength(3);
    expect(after.groups.map((group) => [group.camp, group.declaredCount, group.rows.length])).toEqual([
      ['Achim', 2, 2],
      ['Zebulun', 1, 1],
    ]);
    expect(after.totalDeclared).toBe(3);
    expect(after.rows.map((row) => row.syncId).sort()).toEqual(
      appends.map((append) => append.syncId).sort(),
    );
  });

  it('keeps existing rows intact and their SyncIDs aligned when inserting', () => {
    const before = queueRange('Missing Info', [
      ...Array.from({ length: 8 }, () => ({ kind: 'blank' }) as const),
      { kind: 'header' },
      { kind: 'divider', camp: 'Achim', count: 2 },
      { kind: 'row', last: 'Existing-A1', syncId: 'S000000000901' },
      { kind: 'row', last: 'Existing-A2', syncId: 'S000000000902' },
      { kind: 'divider', camp: 'Bnos Naale', count: 1 },
      { kind: 'row', last: 'Existing-B1', syncId: 'S000000000903' },
      { kind: 'total', count: 3 },
    ]);

    const appends = [
      intent('Missing Info', 'Achim'),
      intent('Missing Info', 'Bnos Naale'),
      intent('Missing Info', 'Chayeinu'),
      intent('Missing Info', 'Chayeinu'),
    ];
    const plan = planQueueAppend({ sheet: parseQueueSheet('Missing Info', before), appends });
    const after = parseQueueSheet('Missing Info', applyPlan('Missing Info', before, plan.operations));

    expect(after.groups.map((group) => [group.camp, group.declaredCount, group.rows.length])).toEqual([
      ['Achim', 3, 3],
      ['Bnos Naale', 2, 2],
      ['Chayeinu', 2, 2],
    ]);
    expect(after.rows).toHaveLength(7);
    expect(after.totalDeclared).toBe(7);

    // No row lost its own identifier to the row above or below it — the failure
    // an insert scoped to the data columns would have caused.
    for (const row of after.rows) {
      const last = String(row.values['Last Name'] ?? '');
      if (last.startsWith('Existing-')) {
        expect(row.syncId).toBe(
          { 'Existing-A1': 'S000000000901', 'Existing-A2': 'S000000000902', 'Existing-B1': 'S000000000903' }[last],
        );
      } else {
        expect(appends.map((append) => append.syncId)).toContain(row.syncId);
      }
    }
  });

  it('plans nothing on a second pass over the sheet it just wrote', () => {
    const before = queueRange('United Refuah', [{ kind: 'header' }]);
    const appends = [intent('United Refuah', 'Achim'), intent('United Refuah', 'Bnos Naale')];
    const first = planQueueAppend({ sheet: parseQueueSheet('United Refuah', before), appends });
    const written = applyPlan('United Refuah', before, first.operations);

    // The reconciler would emit nothing second time round, because every source
    // row now has a queue row. Feeding the appends back in anyway would double
    // them, so this asserts the weaker but checkable thing: the written sheet
    // reads back as the plan described it.
    const reparsed = parseQueueSheet('United Refuah', written);
    expect(planQueueAppend({ sheet: reparsed, appends: [] }).operations).toEqual([]);
    expect(reparsed.rows).toHaveLength(first.totalAfter);
    expect(reparsed.shapeDetected).toBe(true);
  });
});

describe('blank required fields', () => {
  const blankRequired: QueueColumn[] = ['Last Name', 'First Name', 'Date of Birth', 'Phone Number'];

  it('shades contiguous columns as one range', () => {
    const plan = planQueueAppend({
      sheet: parseQueueSheet('Missing Info', queueRange('Missing Info', [{ kind: 'header' }])),
      appends: [intent('Missing Info', 'Achim', { blankRequired })],
    });

    const shades = plan.operations.filter((op) => op.kind === 'shade');
    // C-E are adjacent, K is not, so two ranges rather than four.
    expect(shades.map((op) => (op as { address: string }).address)).toEqual(['C3:E3', 'K3:K3']);
    expect(plan.shadedCells).toBe(4);
  });

  it('shades nothing for an append-only destination, which requires nothing', () => {
    const plan = planQueueAppend({
      sheet: parseQueueSheet('United Refuah', queueRange('United Refuah', [{ kind: 'header' }])),
      appends: [intent('United Refuah', 'Achim')],
    });
    expect(plan.operations.some((op) => op.kind === 'shade')).toBe(false);
  });
});
