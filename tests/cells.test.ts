import { describe, expect, it } from 'vitest';
import {
  cellFromGrid,
  columnToIndex,
  indexToColumn,
  offsetColumn,
  parseAddress,
  rangeAddress,
} from '../src/domain/cells';
import { assertNoFormulas, encodeSheet, worksheetPath } from '../src/graph/workbook';
import { isSyncId, newSyncId, normalizeSyncId } from '../src/domain/syncId';
import { formatGroupHeader, parseGroupHeader, resolveSheetName } from '../src/domain/queueSheets';

describe('column arithmetic', () => {
  it('round-trips single and multi-letter columns', () => {
    for (const column of ['A', 'B', 'Z', 'AA', 'AZ', 'BA', 'ZZ', 'AAA']) {
      expect(indexToColumn(columnToIndex(column))).toBe(column);
    }
  });

  it('offsets across the Z boundary', () => {
    expect(offsetColumn('Y', 2)).toBe('AA');
    expect(offsetColumn('A', 12)).toBe('M');
  });

  it('builds range addresses', () => {
    expect(rangeAddress('BA', 2, 'BA', 40)).toBe('BA2:BA40');
  });
});

describe('parseAddress', () => {
  it('parses a sheet-qualified range', () => {
    expect(parseAddress("'2026-07-30'!A1:BA40")).toEqual({
      sheet: '2026-07-30',
      startColumn: 'A',
      startRow: 1,
      endColumn: 'BA',
      endRow: 40,
    });
  });

  it('parses a bare single cell', () => {
    expect(parseAddress('C7')).toMatchObject({ startColumn: 'C', startRow: 7, endRow: 7 });
  });

  it('tolerates absolute references', () => {
    expect(parseAddress('$B$2:$D$9')).toMatchObject({ startColumn: 'B', endColumn: 'D' });
  });
});

describe('cellFromGrid', () => {
  const grid = [
    ['a1', 'b1', 'c1'],
    ['a2', 'b2', 'c2'],
  ];

  it('reads by absolute sheet coordinates', () => {
    expect(cellFromGrid(grid, 1, 'A', 2, 'B')).toBe('b2');
  });

  it('returns undefined outside the grid', () => {
    expect(cellFromGrid(grid, 1, 'A', 9, 'B')).toBeUndefined();
    expect(cellFromGrid(grid, 5, 'A', 1, 'B')).toBeUndefined();
  });

  it('handles a used range that does not start at A1', () => {
    expect(cellFromGrid(grid, 10, 'C', 11, 'D')).toBe('b2');
  });
});

describe('syncId', () => {
  it('mints IDs that validate and are not sequential', () => {
    const ids = new Set(Array.from({ length: 500 }, () => newSyncId()));
    expect(ids.size).toBe(500);
    for (const id of ids) expect(isSyncId(id)).toBe(true);
  });

  it('rejects anything that is not an ID', () => {
    expect(normalizeSyncId('')).toBeUndefined();
    expect(normalizeSyncId(42)).toBeUndefined();
    expect(normalizeSyncId('Smith')).toBeUndefined();
    expect(normalizeSyncId(null)).toBeUndefined();
  });

  it('tolerates casing and surrounding whitespace from a cell read', () => {
    const id = newSyncId();
    expect(normalizeSyncId(`  ${id.toLowerCase()} `)).toBe(id);
  });
});

describe('group headers', () => {
  it('round-trips', () => {
    expect(parseGroupHeader(formatGroupHeader('Camp Ramah', 12))).toEqual({
      camp: 'Camp Ramah',
      count: 12,
    });
  });

  it('singularizes one patient', () => {
    expect(formatGroupHeader('Camp Ramah', 1)).toBe('Camp Ramah - 1 patient');
    expect(parseGroupHeader('Camp Ramah - 1 patient')).toEqual({ camp: 'Camp Ramah', count: 1 });
  });

  it('handles a camp name containing a dash', () => {
    expect(parseGroupHeader('Camp Bnos - Girls - 4 patients')).toEqual({
      camp: 'Camp Bnos - Girls',
      count: 4,
    });
  });

  it('is not fooled by a patient row', () => {
    expect(parseGroupHeader('Smith')).toBeUndefined();
    expect(parseGroupHeader('')).toBeUndefined();
  });
});

describe('write safety', () => {
  it('refuses to write a formula', () => {
    expect(() => assertNoFormulas([['ok', '=SUM(A1:A2)']], 'Missing Info', 'A1:B1')).toThrow(
      /formula/i,
    );
  });

  it('allows a phone number starting with a plus', () => {
    expect(() => assertNoFormulas([['+1 555 123 4567']], 'Missing Info', 'L10')).not.toThrow();
  });

  it('allows ordinary values', () => {
    expect(() => assertNoFormulas([['Smith', 42, null, true]], 'Missing Info', 'A1:D1')).not.toThrow();
  });
});

describe('encodeSheet', () => {
  it('quotes the name and doubles inner apostrophes', () => {
    expect(encodeSheet("Bob's sheet")).toBe("('Bob''s%20sheet')");
  });
});

describe('worksheetPath', () => {
  const WB = '/drives/abc/items/xyz/workbook';

  // The regression: call sites wrote `/worksheets/${encodeSheet(name)}`, giving
  // `/worksheets/('Name')`. Graph rejects the empty segment with a 400 whose
  // message points at the URL, not the sheet — easy to misread as a workbook
  // problem. Every read and write went through it, so nothing worked.
  it('never produces an empty path segment', () => {
    for (const name of [
      'Ineligible & Inactive',
      'United Refuah',
      'Not Accepted ',
      'Missing Info (New)',
      'July 30, 2026',
      "Bob's sheet",
    ]) {
      const path = worksheetPath(WB, name);
      expect(path).not.toContain('/(');
      expect(path).not.toContain('//');
      expect(path).toContain("/worksheets('");
    }
  });

  it('builds the documented OData shape', () => {
    expect(worksheetPath(WB, 'July 30, 2026')).toBe(
      "/drives/abc/items/xyz/workbook/worksheets('July%2030%2C%202026')",
    );
  });

  it('encodes an ampersand in the sheet name', () => {
    expect(worksheetPath(WB, 'Ineligible & Inactive')).toBe(
      "/drives/abc/items/xyz/workbook/worksheets('Ineligible%20%26%20Inactive')",
    );
  });
});

describe('resolveSheetName', () => {
  // The live workbook has a tab named `Not Accepted ` — trailing space. An
  // exact match reports "not present", and a queue with no sheet silently never
  // gets populated.
  const actual = ['Not Accepted ', 'Missing Info (New)', 'Ineligible & Inactive', 'July 30, 2026'];

  it('finds a tab whose name has a trailing space', () => {
    expect(resolveSheetName('Not Accepted', actual)).toBe('Not Accepted ');
  });

  it('prefers an exact match when one exists', () => {
    expect(resolveSheetName('Ineligible & Inactive', actual)).toBe('Ineligible & Inactive');
  });

  it('forgives casing and doubled internal spaces', () => {
    expect(resolveSheetName('not  accepted', actual)).toBe('Not Accepted ');
  });

  it('returns undefined rather than guessing between two similar tabs', () => {
    // Populating a stale tab is worse than populating none.
    expect(resolveSheetName('Missing Info', ['Missing Info ', 'missing info'])).toBeUndefined();
  });

  it('returns undefined when nothing matches', () => {
    expect(resolveSheetName('Verify Insurance', actual)).toBeUndefined();
  });
});
