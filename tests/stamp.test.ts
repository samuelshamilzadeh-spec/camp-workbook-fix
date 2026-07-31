import { describe, expect, it } from 'vitest';
import { planColumnWrites, totalCells } from '../src/domain/stamp';

const plan = (stamps: [number, string][], existing: [number, string][] = [], maxGap?: number) =>
  planColumnWrites({
    column: 'BA',
    stamps: new Map(stamps),
    existing: new Map(existing),
    ...(maxGap === undefined ? {} : { maxGap }),
  });

describe('planColumnWrites', () => {
  it('writes a contiguous run as one block', () => {
    const writes = plan([
      [2, 'A'],
      [3, 'B'],
      [4, 'C'],
    ]);
    expect(writes).toHaveLength(1);
    expect(writes[0]).toMatchObject({ address: 'BA2:BA4', newValues: 3 });
    expect(writes[0]!.values).toEqual([['A'], ['B'], ['C']]);
  });

  it('PRESERVES a value sitting inside the block that is not being stamped', () => {
    // The unrecoverable mistake: a range write replaces every cell in it, so an
    // untouched row in the middle would be blanked. That ID is the only link
    // between a queue row and its patient's daily row.
    const writes = plan(
      [
        [2, 'A'],
        [5, 'B'],
      ],
      [[3, 'KEEP-ME']],
    );
    expect(writes[0]!.values).toEqual([['A'], ['KEEP-ME'], [null], ['B']]);
  });

  it('writes null only where the cell is genuinely empty', () => {
    const writes = plan([
      [2, 'A'],
      [4, 'B'],
    ]);
    expect(writes[0]!.values).toEqual([['A'], [null], ['B']]);
  });

  it('splits distant rows into separate writes rather than one giant range', () => {
    // Otherwise stamping rows 3 and 900 rewrites 898 cells in between.
    const writes = plan(
      [
        [3, 'A'],
        [900, 'B'],
      ],
      [],
      50,
    );
    expect(writes).toHaveLength(2);
    expect(writes[0]!.address).toBe('BA3:BA3');
    expect(writes[1]!.address).toBe('BA900:BA900');
    expect(totalCells(writes)).toBe(2);
  });

  it('keeps rows within the gap threshold together', () => {
    const writes = plan(
      [
        [3, 'A'],
        [40, 'B'],
      ],
      [],
      50,
    );
    expect(writes).toHaveLength(1);
    expect(writes[0]!.address).toBe('BA3:BA40');
  });

  it('does not count an unchanged stamp as new work', () => {
    // Re-running must be a no-op in effect, so the report is honest about how
    // much is actually changing.
    const writes = plan([[2, 'SAME']], [[2, 'SAME']]);
    expect(writes[0]!.newValues).toBe(0);
  });

  it('returns nothing when there is nothing to stamp', () => {
    expect(plan([])).toEqual([]);
    expect(totalCells([])).toBe(0);
  });

  it('ignores row 0, which is not a real Excel row', () => {
    expect(plan([[0, 'A']])).toEqual([]);
  });

  it('is idempotent: planning twice gives the same writes', () => {
    const args: [number, string][] = [
      [2, 'A'],
      [7, 'B'],
    ];
    expect(JSON.stringify(plan(args))).toBe(JSON.stringify(plan(args)));
  });
});
