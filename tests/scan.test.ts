import { describe, expect, it } from 'vitest';
import { planScan, type ScanOptions } from '../src/domain/dailySheets';

const TODAY = new Date('2026-08-03T14:00:00Z');

/** A season of daily sheets, plus sheets the office created in advance. */
function season(): string[] {
  const sheets: string[] = [];
  for (let day = 1; day <= 31; day++) {
    sheets.push(`2026-06-${String(day).padStart(2, '0')}`);
    sheets.push(`2026-07-${String(day).padStart(2, '0')}`);
  }
  // August, including sheets ahead of today (the 3rd).
  for (let day = 1; day <= 8; day++) {
    sheets.push(`2026-08-0${day}`);
  }
  return sheets;
}

function plan(overrides: Partial<ScanOptions> = {}) {
  return planScan({
    allSheetNames: season(),
    referencedSheets: [],
    today: TODAY,
    hotDaysBack: 7,
    hotDaysForward: 14,
    coldBatchSize: 10,
    // Deliberately below the 70-sheet fixture, so these exercise the tiering.
    // The full-scan path (the normal one for a single season) is covered
    // separately below.
    maxSheetsPerCycle: 30,
    cursor: 0,
    ...overrides,
  });
}

describe('planScan', () => {
  it("always scans today's sheet, even though later-dated sheets exist", () => {
    // The regression this guards: the office creates sheets in advance, so
    // "the N sheets with the latest dates" picks empty future sheets and misses
    // the one staff are actually working in.
    const result = plan();
    expect(result.hot).toContain('2026-08-03');
  });

  it('scans sheets created in advance, so they are live the moment staff fill them', () => {
    const result = plan();
    expect(result.hot).toContain('2026-08-04');
    expect(result.hot).toContain('2026-08-08');
  });

  it('scans recent past days every cycle', () => {
    const result = plan();
    expect(result.hot).toContain('2026-07-31');
    expect(result.hot).toContain('2026-07-28'); // 6 days back
  });

  it('leaves older sheets out of the hot window', () => {
    const result = plan();
    expect(result.hot).not.toContain('2026-06-15');
  });

  it('keeps the per-cycle read count bounded well below the sheet total', () => {
    const result = plan();
    expect(result.totalDaily).toBe(70);
    // Hot window (~15 days) plus one cold slice, not all 70.
    expect(result.sheets.length).toBeLessThan(30);
  });

  it('covers every sheet across a full sweep of the rotation', () => {
    const seen = new Set<string>();
    let cursor = 0;

    // One extra pass, since the hot set is re-read every cycle anyway.
    for (let cycle = 0; cycle < 20; cycle++) {
      const result = plan({ cursor });
      result.sheets.forEach((sheet) => seen.add(sheet));
      cursor = result.nextCursor;
    }

    expect(seen.size).toBe(70);
    for (const sheet of season()) expect(seen.has(sheet)).toBe(true);
  });

  it('reports how many cycles a full sweep takes', () => {
    const result = plan();
    // 70 sheets, ~15 hot, so ~55 cold at 10 per cycle.
    expect(result.sweepCycles).toBeLessThanOrEqual(6);
    expect(result.sweepCycles).toBeGreaterThan(0);
  });

  it('advances the cursor so successive cycles read different cold sheets', () => {
    const first = plan({ cursor: 0 });
    const second = plan({ cursor: first.nextCursor });
    expect(second.cold).not.toEqual(first.cold);
    expect(first.cold).toHaveLength(10);
  });

  it('wraps the cursor around the end of the cold pool', () => {
    const result = plan({ cursor: 999999 });
    expect(result.cold).toHaveLength(10);
    expect(result.nextCursor).toBeGreaterThanOrEqual(0);
  });

  it('scans a queue row source sheet every cycle however old it is', () => {
    const result = plan({ referencedSheets: ['2026-06-02'] });
    expect(result.hot).toContain('2026-06-02');
  });

  it('ignores a referenced sheet that does not exist', () => {
    const result = plan({ referencedSheets: ['2020-01-01'] });
    expect(result.sheets).not.toContain('2020-01-01');
  });

  it('never scans queue sheets or known non-daily tabs', () => {
    const result = plan({
      allSheetNames: [...season(), 'Missing Info', 'United Refresh', 'Claude Log'],
    });
    expect(result.sheets).not.toContain('Missing Info');
    expect(result.sheets).not.toContain('United Refresh');
    expect(result.sheets).not.toContain('Claude Log');
  });

  it('falls back to the most recent past sheets when nothing is near today', () => {
    // Off-season: the workbook is all last year's sheets. The hot window is
    // empty, and waiting for the rotation would make the job look dead.
    const result = plan({
      allSheetNames: ['2025-06-01', '2025-06-02', '2025-06-03', '2025-06-04'],
      maxSheetsPerCycle: 0,
    });
    expect(result.hot).toContain('2025-06-04');
    expect(result.hot.length).toBeGreaterThan(0);
  });

  it('reports sheet names that match the daily pattern but do not parse', () => {
    const result = plan({ allSheetNames: [...season(), '13/45/2026'] });
    expect(result.unparseable).toContain('13/45/2026');
  });

  it('still scans unparseable sheets, on the tail of the rotation', () => {
    const seen = new Set<string>();
    let cursor = 0;
    for (let cycle = 0; cycle < 20; cycle++) {
      const result = plan({ allSheetNames: [...season(), '13/45/2026'], cursor });
      result.sheets.forEach((sheet) => seen.add(sheet));
      cursor = result.nextCursor;
    }
    expect(seen.has('13/45/2026')).toBe(true);
  });

  it('handles a workbook with no daily sheets at all', () => {
    const result = plan({ allSheetNames: ['Missing Info', 'United Refresh'] });
    expect(result.sheets).toEqual([]);
    expect(result.totalDaily).toBe(0);
    expect(result.sweepCycles).toBe(0);
  });

  it('handles every sheet being hot, leaving nothing to rotate', () => {
    const result = plan({
      allSheetNames: ['2026-08-02', '2026-08-03', '2026-08-04'],
      maxSheetsPerCycle: 0,
    });
    expect(result.cold).toEqual([]);
    expect(result.hot).toHaveLength(3);
    expect(result.sweepCycles).toBe(0);
  });

  it('respects a widened forward window when the office works further ahead', () => {
    const sheets = ['2026-08-03', '2026-09-15'];
    expect(plan({ allSheetNames: sheets, maxSheetsPerCycle: 0 }).hot).not.toContain('2026-09-15');
    expect(
      plan({ allSheetNames: sheets, maxSheetsPerCycle: 0, hotDaysForward: 60 }).hot,
    ).toContain('2026-09-15');
  });
});

describe('planScan, full-scan mode', () => {
  // The normal case: one camp season, about 60 sheets, all read every cycle.
  it('reads every daily sheet when the workbook is one season', () => {
    const result = plan({ maxSheetsPerCycle: 90 });

    expect(result.full).toBe(true);
    expect(result.sheets).toHaveLength(70);
    expect(result.cold).toEqual([]);
    expect(result.sweepCycles).toBe(1);
    for (const sheet of season()) expect(result.sheets).toContain(sheet);
  });

  it('reads a June sheet in the same cycle as an August one, with no rotation wait', () => {
    // The whole point: a keyword typed onto any sheet is picked up next cycle,
    // not on that sheet's turn.
    const result = plan({ maxSheetsPerCycle: 90 });
    expect(result.sheets).toContain('2026-06-01');
    expect(result.sheets).toContain('2026-08-08');
  });

  it('still excludes queue sheets and non-daily tabs in full-scan mode', () => {
    const result = plan({
      allSheetNames: [...season(), 'Missing Info', 'United Refresh'],
      maxSheetsPerCycle: 90,
    });
    expect(result.full).toBe(true);
    expect(result.sheets).not.toContain('Missing Info');
    expect(result.sheets).not.toContain('United Refresh');
  });

  it('degrades to tiering once the workbook outgrows the threshold', () => {
    // Adding the 2027 season is what eventually trips this.
    const twoSeasons = [...season(), ...season().map((s) => s.replace('2026', '2027'))];
    const result = plan({ allSheetNames: twoSeasons, maxSheetsPerCycle: 90 });

    expect(result.totalDaily).toBe(140);
    expect(result.full).toBe(false);
    expect(result.cold.length).toBeGreaterThan(0);
    expect(result.sheets.length).toBeLessThan(140);
  });

  it('does not lose a sheet at the threshold boundary', () => {
    const exactly = season().slice(0, 40);
    expect(plan({ allSheetNames: exactly, maxSheetsPerCycle: 40 }).sheets).toHaveLength(40);
    expect(plan({ allSheetNames: exactly, maxSheetsPerCycle: 39 }).full).toBe(false);
  });
});
