import { rangeAddress } from './cells';

/**
 * Plans writes into a single column, given a sparse set of row -> value.
 *
 * Two properties matter, and both are about not destroying what is already
 * there.
 *
 * A range write replaces every cell in the range, so writing rows 5 and 40 as
 * one block would blank rows 6..39. Any row inside a block that is not being
 * stamped must therefore be written back with the value it already holds — that
 * is what `existing` is for. Getting this wrong would erase IDs that queue rows
 * depend on, which is unrecoverable: the link between a queue row and its
 * patient's daily row would be gone with nothing left to rebuild it from.
 *
 * And a run of scattered rows should not become one enormous range. A gap wider
 * than `maxGap` starts a new block, so stamping rows 3 and 900 is two small
 * writes rather than one 898-row write that touches everything between.
 */

export interface ColumnWrite {
  firstRow: number;
  lastRow: number;
  address: string;
  values: unknown[][];
  /** How many cells in this block are new stamps rather than preserved. */
  newValues: number;
}

export interface PlanColumnWritesInput {
  column: string;
  /** Row -> the value to stamp. */
  stamps: ReadonlyMap<number, string>;
  /** Row -> the value already in that cell, for every row that has one. */
  existing: ReadonlyMap<number, string>;
  /** Rows further apart than this start a new write. */
  maxGap?: number;
}

export function planColumnWrites(input: PlanColumnWritesInput): ColumnWrite[] {
  const maxGap = input.maxGap ?? 50;
  const rows = [...input.stamps.keys()].filter((row) => row > 0).sort((a, b) => a - b);
  if (rows.length === 0) return [];

  // Group into blocks separated by gaps wider than maxGap.
  const blocks: number[][] = [];
  let current: number[] = [rows[0]!];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]!;
    if (row - current[current.length - 1]! > maxGap) {
      blocks.push(current);
      current = [row];
    } else {
      current.push(row);
    }
  }
  blocks.push(current);

  return blocks.map((block) => {
    const firstRow = block[0]!;
    const lastRow = block[block.length - 1]!;
    const values: unknown[][] = [];
    let newValues = 0;

    for (let row = firstRow; row <= lastRow; row++) {
      const stamp = input.stamps.get(row);
      if (stamp !== undefined) {
        values.push([stamp]);
        if (input.existing.get(row) !== stamp) newValues++;
        continue;
      }
      // Not being stamped: write back whatever is there so the range write
      // leaves it untouched. `null` would blank it.
      const held = input.existing.get(row);
      values.push([held ?? null]);
    }

    return {
      firstRow,
      lastRow,
      address: rangeAddress(input.column, firstRow, input.column, lastRow),
      values,
      newValues,
    };
  });
}

/** Total cells a plan will write, for cost reporting. */
export function totalCells(writes: readonly ColumnWrite[]): number {
  return writes.reduce((sum, write) => sum + write.values.length, 0);
}
