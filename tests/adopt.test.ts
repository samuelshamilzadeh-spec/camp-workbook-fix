import { describe, expect, it } from 'vitest';
import {
  excelSerialToDate,
  planAdoption,
  resolveSourceSheet,
  sameDateOrValue,
  toDateKey,
} from '../src/domain/adopt';
import { sameValue } from '../src/domain/reconcile';
import type { DailyRow } from '../src/domain/dailySheets';
import type { QueueRow } from '../src/domain/queueSheets';
import type { QueueSheetName } from '../src/config';

function dailyRow(sheet: string, row: number, last: string, first: string, dob: string, syncId?: string): DailyRow {
  return {
    sheet,
    row,
    syncId,
    statusRaw: 'missing info',
    outcome: { kind: 'queued', destination: 'Missing Info', keyword: 'missing info', matched: ['missing info'] },
    ambiguous: false,
    camp: 'Achim',
    fields: { 'Last Name': last, 'First Name': first, 'Date of Birth': dob },
  };
}

function queueRow(
  row: number,
  dateOfVisit: unknown,
  sourceRow: unknown,
  last?: string,
  first?: string,
  dob?: string,
  syncId?: string,
): QueueRow {
  return {
    sheet: 'Missing Info' as QueueSheetName,
    row,
    syncId,
    camp: 'Achim',
    values: {
      'Date of Visit': dateOfVisit,
      'Source Row': sourceRow,
      'Last Name': last,
      'First Name': first,
      'Date of Birth': dob,
    },
  };
}

const SHEET = 'July 22, 2026';  // Excel serial 46225, as seen on the live queue tabs
const sheets = [{ sheet: SHEET, rows: [], unrecognizedRows: [] }];

function index(rows: DailyRow[]) {
  const inner = new Map<number, DailyRow>();
  for (const r of rows) inner.set(r.row, r);
  return new Map([[SHEET, inner]]);
}

let n = 0;
const mint = () => `S${String(++n).padStart(12, '0')}`;

describe('excelSerialToDate', () => {
  it('converts the serials the live queue tabs use', () => {
    // The queue tabs store Date of Visit as a serial while the daily tabs are
    // named `July 5, 2026`, so one has to reach the other.
    // Verified against the serials actually present on the queue tabs.
    expect(excelSerialToDate(46225)?.toISOString().slice(0, 10)).toBe('2026-07-22');
    expect(excelSerialToDate(46229)?.toISOString().slice(0, 10)).toBe('2026-07-26');
    expect(excelSerialToDate(46219)?.toISOString().slice(0, 10)).toBe('2026-07-16');
  });

  it('rejects nonsense rather than returning an epoch date', () => {
    expect(excelSerialToDate(0)).toBeNull();
    expect(excelSerialToDate(-5)).toBeNull();
    expect(excelSerialToDate(Number.NaN)).toBeNull();
  });
});

describe('resolveSourceSheet', () => {
  const names = ['July 22, 2026', 'July 26, 2026', 'June 30, 2026'];

  it('resolves an Excel serial to the sheet named for that date', () => {
    expect(resolveSourceSheet(46225, names)).toBe('July 22, 2026');
    expect(resolveSourceSheet(46229, names)).toBe('July 26, 2026');
  });

  it('accepts a sheet name directly, for rows this code writes', () => {
    expect(resolveSourceSheet('July 26, 2026', names)).toBe('July 26, 2026');
  });

  it('returns undefined when no sheet matches that date', () => {
    expect(resolveSourceSheet(1, names)).toBeUndefined();
    expect(resolveSourceSheet('', names)).toBeUndefined();
  });
});

describe('planAdoption', () => {
  it('links an existing queue row to its source row', () => {
    const daily = [dailyRow(SHEET, 12, 'Halperin', 'Menachem', '2015-03-01')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12, 'Halperin', 'Menachem', '2015-03-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });

    expect(result.problems).toHaveLength(0);
    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]).toMatchObject({
      queueRow: 3,
      sourceSheet: SHEET,
      sourceRow: 12,
    });
    expect(result.matches[0]!.confirmedBy).toEqual(['Last Name', 'First Name', 'Date of Birth']);
  });

  it('reuses the source row ID when it already has one', () => {
    // Otherwise the same patient ends up with two identities.
    const daily = [dailyRow(SHEET, 12, 'Halperin', 'Menachem', '2015-03-01', 'S000000000042')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12, 'Halperin', 'Menachem', '2015-03-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });
    expect(result.matches[0]!.syncId).toBe('S000000000042');
  });

  it('REFUSES to adopt when the pointer landed on a different patient', () => {
    // The whole reason SyncID exists: the old pointer breaks silently when rows
    // are inserted, deleted or sorted. It still resolves to a row — just the
    // wrong one. Binding here would send one patient's edits to another
    // patient's record.
    const daily = [dailyRow(SHEET, 12, 'Weinberger', 'Binyomin', '2013-08-14')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12, 'Halperin', 'Menachem', '2015-03-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });

    expect(result.matches).toHaveLength(0);
    expect(result.problems[0]).toMatchObject({
      queueRow: 3,
      reason: 'identity-mismatch',
      sourceRow: 12,
    });
    expect(result.problems[0]!.mismatchedFields).toContain('Last Name');
  });

  it('rejects a partial identity match too', () => {
    // Same surname, different person. One agreeing field is not agreement.
    const daily = [dailyRow(SHEET, 12, 'Weiss', 'Yehuda', '2015-03-01')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12, 'Weiss', 'Shimon', '2015-03-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });
    expect(result.matches).toHaveLength(0);
    expect(result.problems[0]!.mismatchedFields).toEqual(['First Name']);
  });

  it('will not adopt on no evidence at all', () => {
    const daily = [dailyRow(SHEET, 12, '', '', '')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12)],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });
    expect(result.problems[0]).toMatchObject({ reason: 'too-little-identity' });
  });

  it('never lets two queue rows claim the same source row', () => {
    // Two queue rows pointing at one source row would aim both patients' edits
    // at the same cells.
    const daily = [dailyRow(SHEET, 12, 'Cohen', 'Dovid', '2014-01-01')];
    const result = planAdoption({
      queueRows: [
        queueRow(3, 46225, 12, 'Cohen', 'Dovid', '2014-01-01'),
        queueRow(4, 46225, 12, 'Cohen', 'Dovid', '2014-01-01'),
      ],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });

    expect(result.matches).toHaveLength(1);
    expect(result.problems[0]).toMatchObject({ reason: 'source-already-adopted', queueRow: 4 });
  });

  it('reports a pointer to a row that no longer exists', () => {
    const daily = [dailyRow(SHEET, 12, 'Cohen', 'Dovid', '2014-01-01')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 999, 'Cohen', 'Dovid', '2014-01-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });
    expect(result.problems[0]).toMatchObject({ reason: 'source-row-out-of-range', sourceRow: 999 });
  });

  it('reports a Date of Visit that matches no daily sheet', () => {
    const daily = [dailyRow(SHEET, 12, 'Cohen', 'Dovid', '2014-01-01')];
    const result = planAdoption({
      queueRows: [queueRow(3, 1, 12, 'Cohen', 'Dovid', '2014-01-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });
    expect(result.problems[0]).toMatchObject({ reason: 'unknown-source-sheet' });
  });

  it('skips rows that already carry an ID', () => {
    const daily = [dailyRow(SHEET, 12, 'Cohen', 'Dovid', '2014-01-01')];
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12, 'Cohen', 'Dovid', '2014-01-01', 'S000000000009')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: mint,
    });
    expect(result.matches).toHaveLength(0);
    expect(result.problems).toHaveLength(0);
  });

  it('is idempotent: adopting twice over the same state changes nothing', () => {
    const daily = [dailyRow(SHEET, 12, 'Cohen', 'Dovid', '2014-01-01')];
    const input = {
      queueRows: [queueRow(3, 46225, 12, 'Cohen', 'Dovid', '2014-01-01')],
      daily: sheets,
      dailyRowsBySheet: index(daily),
      newSyncId: () => 'S000000000001',
    };
    expect(JSON.stringify(planAdoption(input))).toBe(JSON.stringify(planAdoption(input)));
  });
});

describe('date-aware identity comparison', () => {
  it('treats an Excel serial and a text date as the same birthday', () => {
    // The live workbook stores the same birthday two ways: 40147 on the queue
    // tabs, "11/30/2009" on the daily sheets. Comparing string forms made 13
    // real patients look like 13 different people.
    expect(toDateKey(40147)).toBe('2009-11-30');
    expect(toDateKey('11/30/2009')).toBe('2009-11-30');
    expect(toDateKey('2009-11-30')).toBe('2009-11-30');
  });

  it('matches across the two representations', () => {
    const never = () => false;
    expect(sameDateOrValue(40147, '11/30/2009', never)).toBe(true);
    expect(sameDateOrValue('11/30/2009', 40147, never)).toBe(true);
  });

  it('still separates genuinely different birthdays', () => {
    const never = () => false;
    expect(sameDateOrValue(40147, '11/30/2010', never)).toBe(false);
    expect(sameDateOrValue(40147, 40148, never)).toBe(false);
  });

  it('falls back to the ordinary comparison for things that are not dates', () => {
    expect(sameDateOrValue('Cohen', 'Cohen', sameValue)).toBe(true);
    expect(sameDateOrValue('Cohen', 'Cohn', sameValue)).toBe(false);
  });

  it('adopts a row whose only disagreement was the date format', () => {
    const daily = new Map([
      [SHEET, new Map([[12, dailyRow(SHEET, 12, 'Levin', 'Chanie', '11/30/2009')]])],
    ]);
    const result = planAdoption({
      queueRows: [queueRow(3, 46225, 12, 'Levin', 'Chanie', 40147 as unknown as string)],
      daily: sheets,
      dailyRowsBySheet: daily,
      newSyncId: mint,
    });
    expect(result.problems).toHaveLength(0);
    expect(result.matches).toHaveLength(1);
  });
});
