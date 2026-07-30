/** Excel A1 address arithmetic. Pure, so it is cheap to test. */

export function columnToIndex(column: string): number {
  const letters = column.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(letters)) {
    throw new Error(`Not a column letter: "${column}"`);
  }
  let index = 0;
  for (const char of letters) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1; // zero-based
}

export function indexToColumn(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`Not a column index: ${index}`);
  }
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function offsetColumn(column: string, offset: number): string {
  return indexToColumn(columnToIndex(column) + offset);
}

export function cellAddress(column: string, row: number): string {
  return `${column.toUpperCase()}${row}`;
}

export function rangeAddress(
  startColumn: string,
  startRow: number,
  endColumn: string,
  endRow: number,
): string {
  return `${cellAddress(startColumn, startRow)}:${cellAddress(endColumn, endRow)}`;
}

/** Reads a cell out of a usedRange grid by absolute sheet coordinates. */
export function cellFromGrid(
  values: unknown[][],
  gridFirstRow: number,
  gridFirstColumn: string,
  row: number,
  column: string,
): unknown {
  const rowIndex = row - gridFirstRow;
  const colIndex = columnToIndex(column) - columnToIndex(gridFirstColumn);
  if (rowIndex < 0 || colIndex < 0) return undefined;
  return values[rowIndex]?.[colIndex];
}

/** `Sheet1!B2:D40` or `B2:D40` -> its parts. usedRange returns the qualified form. */
export function parseAddress(address: string): {
  sheet: string | undefined;
  startColumn: string;
  startRow: number;
  endColumn: string;
  endRow: number;
} {
  const bang = address.lastIndexOf('!');
  const sheet = bang === -1 ? undefined : unquoteSheet(address.slice(0, bang));
  const body = bang === -1 ? address : address.slice(bang + 1);

  const match = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(body.trim());
  if (!match) throw new Error(`Unparseable range address: "${address}"`);

  const startColumn = match[1]!.toUpperCase();
  const startRow = Number(match[2]);
  return {
    sheet,
    startColumn,
    startRow,
    endColumn: (match[3] ?? match[1]!).toUpperCase(),
    endRow: match[4] ? Number(match[4]) : startRow,
  };
}

function unquoteSheet(raw: string): string {
  const trimmed = raw.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

export function isBlank(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}
