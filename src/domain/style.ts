import {
  BLANK_REQUIRED_FILL,
  LAYOUT,
  REQUIRED_FIELDS,
  RESOLVED_DROPDOWN_VALUE,
  STYLE,
  queueColumnsFor,
  type QueueColumn,
  type WorkbookLayout,
} from '../config';
import { isBlank, offsetColumn, rangeAddress } from './cells';
import { lastQueueColumnLetter, type ParsedQueueSheet } from './queueSheets';

/**
 * How a queue sheet is dressed.
 *
 * The office asked for one consistent treatment across all five tabs, so the
 * rules live in `STYLE` and the drawing lives here. Pure: it produces a list of
 * operations and touches nothing, so the whole appearance of a tab can be
 * printed and read before a cell changes.
 *
 * The design is subtractive. A spreadsheet's default look — gridlines
 * everywhere, every column the same width, dates as text, numbers left-aligned
 * against text — carries no information, and the reader has to filter all of it
 * to find the row they want. So: one dark header bar, horizontal rules only,
 * camp dividers as quiet bands rather than shouty labels, numeric columns
 * centred, dates formatted as dates, and every column sized to what it holds.
 *
 * Colour is spent only where it means something. The one loud thing stays loud:
 * `BLANK_REQUIRED_FILL` is the office's red and marks a field somebody has to
 * chase. Everything else is muted so that red is the only thing that catches the
 * eye.
 */

export interface StyleOperation {
  kind:
    | 'fill'
    | 'font'
    | 'format'
    | 'border'
    | 'number-format'
    | 'validation'
    | 'freeze';
  address: string;
  /** Human-readable purpose, for the dry-run report. */
  what: string;
  /** Fields carried to the Graph call. Only one shape is set per kind. */
  fill?: string;
  font?: { bold?: boolean; color?: string; size?: number; name?: string };
  format?: {
    horizontalAlignment?: 'General' | 'Left' | 'Center' | 'Right';
    verticalAlignment?: 'Top' | 'Center' | 'Bottom';
    rowHeight?: number;
    columnWidth?: number;
  };
  border?: {
    edge: 'EdgeTop' | 'EdgeBottom' | 'InsideHorizontal';
    style: string;
    color?: string;
    weight?: 'Hairline' | 'Thin' | 'Medium' | 'Thick';
  };
  numberFormat?: { format: string; rows: number; columns: number };
  validation?: { values: string[] };
  freeze?: { rows: number };
}

export interface StylePlan {
  operations: StyleOperation[];
  /** Rows the plan restyles. Divider and total rows included. */
  rows: number;
}

export interface PlanQueueStyleInput {
  sheet: ParsedQueueSheet;
  layout?: WorkbookLayout;
  /**
   * Re-shade blank required fields. Off by default: the appender already shades
   * a row as it writes it, and re-deriving it here would need the full grid.
   */
  shadeBlankRequired?: boolean;
}

export function planQueueStyle(input: PlanQueueStyleInput): StylePlan {
  const layout = input.layout ?? LAYOUT;
  const { sheet } = input;
  const destination = sheet.sheet;
  const columns = queueColumnsFor(destination);

  const first = layout.queue.firstColumn;
  const last = lastQueueColumnLetter(destination, layout);
  const header = sheet.headerRow;
  const bodyFirst = sheet.firstDataRow;
  const bodyLast = Math.max(sheet.lastRow, bodyFirst);

  const operations: StyleOperation[] = [];
  const add = (operation: StyleOperation): void => {
    operations.push(operation);
  };

  const headerRange = rangeAddress(first, header, last, header);
  const bodyRange = rangeAddress(first, bodyFirst, last, bodyLast);

  // --- The header bar -------------------------------------------------------
  add({ kind: 'fill', address: headerRange, what: 'header bar', fill: STYLE.headerFill });
  add({
    kind: 'font',
    address: headerRange,
    what: 'header text',
    font: {
      bold: true,
      color: STYLE.headerFont,
      size: STYLE.headerFontSize,
      name: STYLE.fontName,
    },
  });
  add({
    kind: 'format',
    address: headerRange,
    what: 'header alignment and height',
    format: {
      horizontalAlignment: 'Left',
      verticalAlignment: 'Center',
      rowHeight: STYLE.headerRowHeight,
    },
  });
  // Keeps the header visible against the first camp divider, which is also a
  // filled band.
  add({
    kind: 'border',
    address: headerRange,
    what: 'rule under the header',
    border: { edge: 'EdgeBottom', style: 'Continuous', color: STYLE.headerFill, weight: 'Medium' },
  });
  add({ kind: 'freeze', address: headerRange, what: 'freeze the header', freeze: { rows: header } });

  // --- The body -------------------------------------------------------------
  if (bodyLast >= bodyFirst) {
    add({
      kind: 'font',
      address: bodyRange,
      what: 'body text',
      font: { bold: false, color: '#000000', size: STYLE.fontSize, name: STYLE.fontName },
    });
    add({
      kind: 'format',
      address: bodyRange,
      what: 'body alignment',
      format: { horizontalAlignment: 'Left', verticalAlignment: 'Center' },
    });
    // Horizontal rules only. Vertical lines are what make a sheet look like a
    // form rather than a table.
    add({
      kind: 'border',
      address: bodyRange,
      what: 'hairline rules between rows',
      border: {
        edge: 'InsideHorizontal',
        style: 'Continuous',
        color: STYLE.ruleColor,
        weight: 'Hairline',
      },
    });

    for (const column of STYLE.centeredColumns) {
      const index = columns.indexOf(column);
      if (index === -1) continue;
      const letter = offsetColumn(first, index);
      add({
        kind: 'format',
        address: rangeAddress(letter, bodyFirst, letter, bodyLast),
        what: `centre ${column}`,
        format: { horizontalAlignment: 'Center' },
      });
    }

    // Dates as dates. Written as serials by the appender, they show up as
    // `46233` until this runs — worse than the text they replaced.
    for (const column of STYLE.dateColumns) {
      const index = columns.indexOf(column);
      if (index === -1) continue;
      const letter = offsetColumn(first, index);
      add({
        kind: 'number-format',
        address: rangeAddress(letter, bodyFirst, letter, bodyLast),
        what: `${column} as ${STYLE.dateFormat}`,
        numberFormat: {
          format: STYLE.dateFormat,
          rows: bodyLast - bodyFirst + 1,
          columns: 1,
        },
      });
      add({
        kind: 'format',
        address: rangeAddress(letter, bodyFirst, letter, bodyLast),
        what: `centre ${column}`,
        format: { horizontalAlignment: 'Center' },
      });
    }
  }

  // --- Column widths --------------------------------------------------------
  columns.forEach((column, index) => {
    const width = STYLE.columnWidths[column];
    if (width === undefined) return;
    const letter = offsetColumn(first, index);
    add({
      kind: 'format',
      address: rangeAddress(letter, header, letter, Math.max(bodyLast, header)),
      what: `width of ${column}`,
      format: { columnWidth: width },
    });
  });

  // --- Camp dividers --------------------------------------------------------
  //
  // A band the width of the table, bold, no rule of its own: the fill is enough
  // to say "new section" without competing with the header bar.
  for (const group of sheet.groups) {
    const address = rangeAddress(first, group.headerRow, last, group.headerRow);
    add({ kind: 'fill', address, what: `divider band: ${group.camp}`, fill: STYLE.dividerFill });
    add({
      kind: 'font',
      address,
      what: `divider text: ${group.camp}`,
      font: {
        bold: true,
        color: STYLE.dividerFont,
        size: STYLE.fontSize,
        name: STYLE.fontName,
      },
    });
    add({
      kind: 'format',
      address,
      what: `divider height: ${group.camp}`,
      format: {
        horizontalAlignment: 'Left',
        verticalAlignment: 'Center',
        rowHeight: STYLE.dividerRowHeight,
      },
    });
  }

  // --- The grand total ------------------------------------------------------
  if (sheet.totalRow !== undefined) {
    const address = rangeAddress(first, sheet.totalRow, last, sheet.totalRow);
    add({ kind: 'fill', address, what: 'total row', fill: STYLE.totalFill });
    add({
      kind: 'font',
      address,
      what: 'total text',
      font: { bold: true, color: STYLE.totalFont, size: STYLE.fontSize, name: STYLE.fontName },
    });
    add({
      kind: 'border',
      address,
      what: 'rule above the total',
      border: { edge: 'EdgeTop', style: 'Continuous', color: STYLE.headerFill, weight: 'Medium' },
    });
  }

  // --- The Resolved dropdown ------------------------------------------------
  const resolvedIndex = columns.indexOf('Resolved');
  if (resolvedIndex !== -1 && bodyLast >= bodyFirst) {
    const letter = offsetColumn(first, resolvedIndex);
    add({
      kind: 'validation',
      address: rangeAddress(letter, bodyFirst, letter, bodyLast),
      what: 'Resolved dropdown',
      validation: { values: [RESOLVED_DROPDOWN_VALUE] },
    });
  }

  return { operations, rows: bodyLast - bodyFirst + 1 };
}

/**
 * Which cells on an already-written row still need the red blank-required
 * shading, derived from the values on the sheet rather than from an intent.
 *
 * The appender shades as it writes. This exists for the tabs written before
 * `Resolved` and the date formats landed, and for a row a staff member has since
 * emptied.
 */
export function planBlankRequiredShading(
  sheet: ParsedQueueSheet,
  layout: WorkbookLayout = LAYOUT,
): StyleOperation[] {
  const required = REQUIRED_FIELDS[sheet.sheet];
  if (required.length === 0) return [];

  const columns = queueColumnsFor(sheet.sheet);
  const first = layout.queue.firstColumn;
  const operations: StyleOperation[] = [];

  for (const row of sheet.rows) {
    for (const field of required) {
      if (!isBlank(row.values[field])) continue;
      const index = columns.indexOf(field as QueueColumn);
      if (index === -1) continue;
      const letter = offsetColumn(first, index);
      operations.push({
        kind: 'fill',
        address: rangeAddress(letter, row.row, letter, row.row),
        what: `blank required: ${field}`,
        fill: BLANK_REQUIRED_FILL,
      });
    }
  }

  return operations;
}
