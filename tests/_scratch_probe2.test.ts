import { describe, it } from 'vitest';
import { queueColumnsFor, type QueueSheetName } from '../src/config';
import { planCountRefresh, planQueueAppend, type QueueOperation } from '../src/domain/append';
import { columnToIndex, parseAddress } from '../src/domain/cells';
import { campKey, normalizeCamp } from '../src/domain/dailySheets';
import { parseQueueSheet } from '../src/domain/queueSheets';
import type { AppendQueueRowIntent } from '../src/domain/reconcile';
import type { RangeData } from '../src/graph/workbook';

const WIDTH = 53;
const SYNC_ID_INDEX = 52;

type Entry =
  | { kind: 'blank' }
  | { kind: 'header' }
  | { kind: 'divider'; camp: string; count: number }
  | { kind: 'bare'; camp: string }
  | { kind: 'total'; count: number }
  | { kind: 'row'; last?: string; syncId?: string };

function queueRange(sheet: QueueSheetName, entries: Entry[]): RangeData {
  const grid: unknown[][] = entries.map((entry) => {
    const cells: unknown[] = new Array(WIDTH).fill(null);
    switch (entry.kind) {
      case 'header':
        queueColumnsFor(sheet).forEach((column, index) => { cells[index] = column; });
        break;
      case 'divider': cells[0] = `${entry.camp} - ${entry.count} patients`; break;
      case 'bare': cells[0] = entry.camp; break;
      case 'total': cells[0] = `TOTAL - ${entry.count} patients`; break;
      case 'row': {
        const index = queueColumnsFor(sheet).indexOf('Last Name');
        cells[index] = entry.last ?? 'Surname';
        cells[SYNC_ID_INDEX] = entry.syncId ?? null;
        break;
      }
      case 'blank': break;
    }
    return cells;
  });
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: WIDTH, values: grid };
}

let nextId = 0;
function intent(destination: QueueSheetName, camp: string | undefined): AppendQueueRowIntent {
  nextId++;
  return {
    kind: 'append-queue-row',
    destination,
    syncId: `N${String(nextId).padStart(12, '0')}`,
    sourceSheet: 'July 30, 2026',
    sourceRow: nextId + 1,
    camp: normalizeCamp(camp),
    values: { 'Date of Visit': 'July 30, 2026', 'Source Row': nextId + 1, 'Last Name': `New-${nextId}`, 'First Name': 'G' },
    blankRequired: [],
  };
}

function applyPlan(sheet: QueueSheetName, before: RangeData, operations: QueueOperation[], graphNullSemantics: boolean): RangeData {
  const grid = before.values.map((row) => [...row]);
  const blank = () => new Array(WIDTH).fill(null);
  const at = (row: number) => { while (grid.length < row) grid.push(blank()); return grid[row - 1]!; };
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
        if (graphNullSemantics && (value === null || value === undefined)) return;
        row[firstColumn + columnOffset] = value;
      });
    });
  }
  return { address: `${sheet}!A1:BA${grid.length}`, rowCount: grid.length, columnCount: WIDTH, values: grid };
}

const dump = (r: RangeData) => r.values.map((row, i) => `${i + 1}: A=${JSON.stringify(row[0])} D=${JSON.stringify(row[3])} BA=${JSON.stringify(row[SYNC_ID_INDEX])}`).join('\n');
const ops = (o: QueueOperation[]) => o.map(x => x.kind === 'insert-rows' ? `INS ${x.row} x${x.count}` : x.kind === 'shade' ? `SHADE ${x.address}` : `W ${x.purpose} ${x.address}`);

describe('probe2', () => {
  it('J: camp name with a doubled internal space on the existing divider', () => {
    console.log('campKey("Bnos  Naale")', JSON.stringify(campKey('Bnos  Naale')));
    console.log('campKey(normalizeCamp("Bnos  Naale"))', JSON.stringify(campKey(normalizeCamp('Bnos  Naale'))));
    const before = queueRange('Missing Info', [
      { kind: 'header' },
      { kind: 'divider', camp: 'Bnos  Naale', count: 1 },
      { kind: 'row', last: 'Existing-B1', syncId: 'S0002' },
      { kind: 'total', count: 1 },
    ]);
    const parsed = parseQueueSheet('Missing Info', before);
    console.log('J groups', parsed.groups.map(g => [JSON.stringify(g.camp), g.headerRow, g.rows.length]));
    const plan = planQueueAppend({ sheet: parsed, appends: [intent('Missing Info', 'Bnos  Naale')] });
    console.log('J placements', plan.placements);
    console.log('J ops', ops(plan.operations));
    console.log('J after (graph null semantics)\n' + dump(applyPlan('Missing Info', before, plan.operations, true)));
  });

  it('K: TOTAL row displaced into the middle (as a sort would do)', () => {
    // Sorting the table by Last Name scatters the divider and TOTAL text rows.
    const before = queueRange('Missing Info', [
      { kind: 'header' },
      { kind: 'divider', camp: 'Achim', count: 2 },
      { kind: 'row', last: 'Aaronson', syncId: 'S0001' },
      { kind: 'total', count: 3 },
      { kind: 'divider', camp: 'Zebulun', count: 1 },
      { kind: 'row', last: 'Zimmer', syncId: 'S0003' },
    ]);
    const parsed = parseQueueSheet('Missing Info', before);
    console.log('K parsed', { lastRow: parsed.lastRow, totalRow: parsed.totalRow, groups: parsed.groups.map(g => [g.camp, g.headerRow, g.rows.map(r => r.row)]) });
    const plan = planQueueAppend({ sheet: parsed, appends: [intent('Missing Info', 'Zebulun')] });
    console.log('K ops', ops(plan.operations));
    const after = applyPlan('Missing Info', before, plan.operations, true);
    console.log('K after (graph null semantics)\n' + dump(after));
    const re = parseQueueSheet('Missing Info', after);
    console.log('K reparsed rows', re.rows.length, 'groups', re.groups.map(g => [g.camp, g.declaredCount, g.rows.length]));
  });

  it('L: TOTAL displaced, new camp only', () => {
    const before = queueRange('Missing Info', [
      { kind: 'header' },
      { kind: 'divider', camp: 'Achim', count: 1 },
      { kind: 'row', last: 'Aaronson', syncId: 'S0001' },
      { kind: 'total', count: 2 },
      { kind: 'divider', camp: 'Zebulun', count: 1 },
      { kind: 'row', last: 'Zimmer', syncId: 'S0003' },
    ]);
    const parsed = parseQueueSheet('Missing Info', before);
    const plan = planQueueAppend({ sheet: parsed, appends: [intent('Missing Info', 'Meor')] });
    console.log('L ops', ops(plan.operations));
    console.log('L after\n' + dump(applyPlan('Missing Info', before, plan.operations, true)));
  });

  it('M: planCountRefresh where a stray value sits on the TOTAL row', () => {
    const grid = queueRange('Missing Info', [
      { kind: 'header' },
      { kind: 'divider', camp: 'Achim', count: 1 },
      { kind: 'row', last: 'A1', syncId: 'S1' },
      { kind: 'total', count: 1 },
    ]);
    // someone typed a note in column B of the TOTAL row
    grid.values[3]![1] = 'check me';
    const parsed = parseQueueSheet('Missing Info', grid);
    console.log('M parsed', { totalRow: parsed.totalRow, rows: parsed.rows.length, lastRow: parsed.lastRow, groups: parsed.groups.map(g => [g.camp, g.declaredCount, g.rows.length]) });
    console.log('M countRefresh', planCountRefresh(parsed).map(o => [o.address, o.values[0]![0]]));
    const plan = planQueueAppend({ sheet: parsed, appends: [intent('Missing Info', 'Achim')] });
    console.log('M append ops', ops(plan.operations));
  });

  it('N: many appends into one grown block plus new camps, check totals', () => {
    const before = queueRange('Missing Info', [
      { kind: 'header' },
      { kind: 'divider', camp: 'Achim', count: 2 },
      { kind: 'row', last: 'A1', syncId: 'S1' },
      { kind: 'row', last: 'A2', syncId: 'S2' },
      { kind: 'total', count: 2 },
    ]);
    const parsed = parseQueueSheet('Missing Info', before);
    const plan = planQueueAppend({
      sheet: parsed,
      appends: [
        intent('Missing Info', 'Achim'), intent('Missing Info', 'Achim'),
        intent('Missing Info', 'Bnos'), intent('Missing Info', 'Zed'),
      ],
    });
    console.log('N ops', ops(plan.operations), 'rowsAdded', plan.rowsAdded, 'totalAfter', plan.totalAfter);
    const after = applyPlan('Missing Info', before, plan.operations, true);
    console.log('N after\n' + dump(after));
    const re = parseQueueSheet('Missing Info', after);
    console.log('N reparsed rows', re.rows.length, 'totalDeclared', re.totalDeclared, 'groups', re.groups.map(g => [g.camp, g.declaredCount, g.rows.length]));
  });
});
