import { describe, expect, it, vi } from 'vitest';
import { QUEUE_SHEET_NAMES, queueColumnsFor, type QueueColumn, type QueueSheetName } from '../src/config';
import { parseDailySheet } from '../src/domain/dailySheets';
import { parseQueueSheet } from '../src/domain/queueSheets';
import { reconcile } from '../src/domain/reconcile';
import { applyPlan } from '../src/sync/apply';
import type { RangeData } from '../src/graph/workbook';

/**
 * The cycle applier is the only write path that will ever run unattended, and
 * until now it was the only one with no tests. What matters here is ORDER and
 * what happens when a step cannot run.
 */

const W = 53;
const log = { debug() {}, info() {}, warn: vi.fn(), error() {} } as any;

function dailyRange(sheet: string, rows: any[]): RangeData {
  const g: unknown[][] = [new Array(W).fill(null)];
  for (const r of rows) {
    const c = new Array(W).fill(null);
    c[1] = r.status ?? null;
    c[2] = r.camp ?? null;
    c[3] = r.last ?? null;
    c[4] = r.first ?? null;
    c[5] = r.dob ?? null;
    c[52] = r.syncId ?? null;
    g.push(c);
  }
  return { address: `${sheet}!A1:BA${g.length}`, rowCount: g.length, columnCount: W, values: g };
}

function queueRange(sheet: QueueSheetName, rows: any[]): RangeData {
  const cols = queueColumnsFor(sheet);
  const h = new Array(W).fill(null);
  cols.forEach((c, i) => (h[i] = c));
  const g: unknown[][] = [h];
  for (const r of rows) {
    const c = new Array(W).fill(null);
    // A camp divider is the label in the first column and nothing else, which is
    // exactly how parseQueueSheet recognizes one.
    if (r.divider) {
      c[0] = r.divider;
      g.push(c);
      continue;
    }
    const at = (col: QueueColumn, v: unknown) => {
      const i = cols.indexOf(col);
      if (i !== -1 && v !== undefined) c[i] = v;
    };
    at('Last Name', r.last);
    at('First Name', r.first);
    at('Date of Birth', r.dob);
    c[52] = r.syncId ?? null;
    g.push(c);
  }
  return { address: `${sheet}!A1:BA${g.length}`, rowCount: g.length, columnCount: W, values: g };
}

function fakeWorkbook(syncIdAt: Record<string, string> = {}) {
  const calls: string[] = [];
  const workbook = {
    async getUsedRange(sheet: string) {
      calls.push(`read ${sheet}`);
      return queueRange(sheet as QueueSheetName, []);
    },
    async getRange(sheet: string, address: string) {
      // The identity check reads column BA before it deletes anything, so this
      // has to answer with what is actually at that row.
      return { values: [[syncIdAt[`${sheet}!${address}`] ?? '']] };
    },
    async writeRange(sheet: string, address: string) {
      calls.push(`write ${sheet}!${address}`);
    },
    async insertRows(sheet: string, row: number) {
      calls.push(`insert ${sheet}!${row}`);
    },
    async setFill(sheet: string, address: string) {
      calls.push(`fill ${sheet}!${address}`);
    },
    async deleteRows(sheet: string, row: number) {
      calls.push(`delete ${sheet}!${row}`);
    },
    async clearRange(sheet: string, address: string) {
      calls.push(`clear ${sheet}!${address}`);
    },
  } as any;
  return { workbook, calls };
}

const config = (phase: number) => ({ phase, dryRun: false }) as any;
const tabFor = new Map(QUEUE_SHEET_NAMES.map((n) => [n, n]));

describe('the cycle applier', () => {
  const newPatient = () => {
    const daily = [
      parseDailySheet('July 30, 2026', dailyRange('July 30, 2026', [
        { status: 'missing info', camp: 'Achim', last: 'A', first: 'B', dob: '01/02/2010' },
      ])),
    ];
    const queues = [parseQueueSheet('Missing Info', queueRange('Missing Info', []))];
    return { daily, queues, plan: reconcile({ daily, queues, newSyncId: () => 'S000000000001' }) };
  };

  it('stamps the source before it appends, or the next cycle duplicates the patient', async () => {
    const { workbook, calls } = fakeWorkbook();
    const { daily, queues, plan } = newPatient();

    const result = await applyPlan({ workbook, config: config(2), plan, daily, queues, tabFor, log });

    expect(result.stamped).toBe(1);
    expect(result.appended).toBe(1);
    const stampAt = calls.findIndex((c) => c.startsWith('write July 30, 2026!BA'));
    const appendAt = calls.findIndex((c) => c.startsWith('write Missing Info!A'));
    expect(stampAt).toBeGreaterThanOrEqual(0);
    expect(stampAt).toBeLessThan(appendAt);
  });

  it('does nothing above its phase', async () => {
    const { workbook } = fakeWorkbook();
    const { daily, queues, plan } = newPatient();
    // Phase 2 is stamps and appends only.
    const result = await applyPlan({ workbook, config: config(2), plan, daily, queues, tabFor, log });
    expect(result.removed).toBe(0);
    expect(result.wroteBack).toBe(0);
  });
});

describe('when a destination tab cannot be reached', () => {
  /**
   * The office renamed `Missing Info (New)` to `Missing Info` once already, and
   * the tab stopped resolving with nothing but a warning. If that happens while
   * a patient is moving ONTO that queue, their old row must not be deleted —
   * the replacement was never written, so it is the only row they have.
   */
  const movingPatient = () => {
    const daily = [
      parseDailySheet('July 30, 2026', dailyRange('July 30, 2026', [
        { status: 'missing info', camp: 'Achim', last: 'A', first: 'B', dob: '01/02/2010', syncId: 'S000000000001' },
      ])),
    ];
    // They currently sit on Not Accepted; their status now routes to Missing Info.
    const queues = [
      parseQueueSheet('Not Accepted', queueRange('Not Accepted', [
        { last: 'A', first: 'B', dob: '01/02/2010', syncId: 'S000000000001' },
      ])),
      // Missing Info did not resolve, so it is absent from `queues` entirely.
    ];
    return { daily, queues, plan: reconcile({ daily, queues, newSyncId: () => 'S000000000009' }) };
  };

  it('keeps the old row instead of leaving the patient on no queue at all', async () => {
    const { workbook, calls } = fakeWorkbook();
    const { daily, queues, plan } = movingPatient();

    // The plan wants both: remove from Not Accepted, append to Missing Info.
    expect(plan.intents.some((i) => i.kind === 'remove-queue-row' && i.reason === 'wrong-queue')).toBe(true);
    expect(plan.counts['append-queue-row']).toBe(1);

    const result = await applyPlan({
      workbook,
      config: config(4),
      plan,
      daily,
      queues,
      tabFor: new Map([['Not Accepted', 'Not Accepted ']]), // Missing Info unbound
      log,
    });

    expect(result.appended).toBe(0);
    expect(result.skipped).toBe(1);
    // THE point: nothing was deleted.
    expect(result.removed).toBe(0);
    expect(calls.some((c) => c.startsWith('delete'))).toBe(false);
  });

  it('still removes rows whose own replacement did land', async () => {
    const { workbook } = fakeWorkbook({ 'Not Accepted!BA2': 'S000000000001' });
    const { daily, queues, plan } = movingPatient();

    const result = await applyPlan({
      workbook,
      config: config(4),
      plan,
      daily,
      // Missing Info resolves this time.
      queues: [...queues, parseQueueSheet('Missing Info', queueRange('Missing Info', []))],
      tabFor,
      log,
    });

    expect(result.appended).toBe(1);
    expect(result.removed).toBe(1);
  });
});

/**
 * The tidy-up pass — recounting dividers and dressing a camp block this cycle
 * created — is a consequence of APPENDING, and appending starts at phase 2.
 * Written as early returns, both sat below the phase 3 and phase 4 gates and so
 * only ever ran at phase 4, which is not the phase this goes live in.
 */
describe('the pass that recounts and dresses what was just written', () => {
  // The re-read has to answer with the divider the append just created, the way
  // the live workbook would.
  function workbookThatSeesTheNewBlock() {
    const calls: string[] = [];
    const workbook = {
      async getUsedRange(sheet: string) {
        return queueRange(sheet as QueueSheetName, [
          { divider: 'Achim - 1 patient' },
          { last: 'A', first: 'B', dob: '01/02/2010', syncId: 'S000000000001' },
        ]);
      },
      async getRange() {
        return { values: [['S000000000001']] };
      },
      async writeRange(s: string, a: string) {
        calls.push(`write ${s}!${a}`);
      },
      async insertRows(s: string, r: number) {
        calls.push(`insert ${s}!${r}`);
      },
      async setFill(s: string, a: string) {
        calls.push(`fill ${s}!${a}`);
      },
      async setFont(s: string, a: string) {
        calls.push(`font ${s}!${a}`);
      },
      async setRangeFormat(s: string, a: string) {
        calls.push(`format ${s}!${a}`);
      },
      async setBorder(s: string, a: string) {
        calls.push(`border ${s}!${a}`);
      },
      async deleteRows(s: string, r: number) {
        calls.push(`delete ${s}!${r}`);
      },
      async clearRange(s: string, a: string) {
        calls.push(`clear ${s}!${a}`);
      },
    } as any;
    return { workbook, calls };
  }

  const firstEverCamp = () => {
    const daily = [
      parseDailySheet('July 30, 2026', dailyRange('July 30, 2026', [
        { status: 'missing info', camp: 'Achim', last: 'A', first: 'B', dob: '01/02/2010' },
      ])),
    ];
    const queues = [parseQueueSheet('Missing Info', queueRange('Missing Info', []))];
    return { daily, queues, plan: reconcile({ daily, queues, newSyncId: () => 'S000000000001' }) };
  };

  it('dresses a camp appearing for the first time at PHASE 2, not just at phase 4', async () => {
    const { workbook, calls } = workbookThatSeesTheNewBlock();
    const { daily, queues, plan } = firstEverCamp();

    const result = await applyPlan({
      workbook,
      config: config(2),
      plan,
      daily,
      queues,
      tabFor,
      log,
    });

    expect(result.appended).toBe(1);
    // The block was created, so the divider must not be left as bare text.
    expect(result.styled).toBeGreaterThan(0);
    expect(calls.some((c) => c.startsWith('font Missing Info'))).toBe(true);
  });

  it('recounts the dividers at phase 2 as well', async () => {
    const { workbook } = workbookThatSeesTheNewBlock();
    const { daily, queues, plan } = firstEverCamp();

    const result = await applyPlan({
      workbook,
      config: config(2),
      plan,
      daily,
      queues,
      tabFor,
      log,
    });

    expect(result.recounted).toBeGreaterThan(0);
  });

  it('still writes nothing above its phase', async () => {
    const { workbook, calls } = workbookThatSeesTheNewBlock();
    const { daily, queues, plan } = firstEverCamp();

    const result = await applyPlan({
      workbook,
      config: config(2),
      plan,
      daily,
      queues,
      tabFor,
      log,
    });

    expect(result.removed).toBe(0);
    expect(result.wroteBack).toBe(0);
    expect(calls.some((c) => c.startsWith('delete'))).toBe(false);
    expect(calls.some((c) => c.startsWith('clear'))).toBe(false);
  });
});
