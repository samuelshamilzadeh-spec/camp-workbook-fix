import {
  LAYOUT,
  STYLE,
  queueColumnsFor,
  type WorkbookLayout,
} from '../config';
import { offsetColumn, rangeAddress } from './cells';
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
    | 'number-format';
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
}

export interface StylePlan {
  operations: StyleOperation[];
  /** Rows the plan restyles. Divider and total rows included. */
  rows: number;
}

export interface PlanQueueStyleInput {
  sheet: ParsedQueueSheet;
  layout?: WorkbookLayout;
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
  // NOT DONE HERE: freezing the header row. Graph's v1.0 workbook API rejects
  // `freezePanes/freezeRows`, `freezePanes/freezeAt` and every spelling in
  // between with a 400 against this workbook. It is a one-off that persists in
  // the file, so it is set by hand once per tab — View > Freeze Panes > Freeze
  // Top Row — rather than left here as a call that fails on every run.

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

  // NOT DONE HERE: the `Resolved` dropdown, and freezing the header row.
  //
  // Graph's workbook API refuses both against this workbook — `dataValidation`
  // 400s even on a GET, and every spelling of `freezePanes` 400s too. They are
  // one-off settings that live in the file once set, so they are done by hand per
  // tab rather than left here as calls that fail on every run:
  //
  //   Data > Data Validation > Allow: List > Source: Done   (on the Resolved column)
  //   View > Freeze Panes > Freeze Top Row
  //
  // Neither is load-bearing. `RESOLVED_VALUES` accepts the words staff type, so
  // the marker works with or without a dropdown to click.

  return { operations, rows: bodyLast - bodyFirst + 1 };
}

/**
 * The house style for specific divider rows, and the TOTAL line.
 *
 * A cycle writes a camp divider's TEXT when it creates a block, but nothing was
 * dressing it — only `npm run migrate` did, by hand. So a camp appearing for the
 * first time got a bare label sitting in the middle of banded ones, which is
 * exactly the moment a queue looks broken to somebody reading it.
 *
 * Scoped to the rows named rather than the whole sheet: restyling every divider
 * on a 35-camp tab is a hundred Graph calls, which does not belong in a
 * five-second cycle. A new camp is rare and costs three.
 */
export function planDividerStyle(
  sheet: ParsedQueueSheet,
  camps: readonly string[],
  layout: WorkbookLayout = LAYOUT,
): StyleOperation[] {
  const first = layout.queue.firstColumn;
  const last = lastQueueColumnLetter(sheet.sheet, layout);
  const wanted = new Set(camps.map((camp) => camp.replace(/\s+/g, ' ').trim().toLowerCase()));

  const operations: StyleOperation[] = [];
  const band = (row: number, fill: string, color: string, height: number, what: string) => {
    const address = rangeAddress(first, row, last, row);
    operations.push({ kind: 'fill', address, what, fill });
    operations.push({
      kind: 'font',
      address,
      what,
      font: { bold: true, color, size: STYLE.fontSize, name: STYLE.fontName },
    });
    operations.push({
      kind: 'format',
      address,
      what,
      format: { horizontalAlignment: 'Left', verticalAlignment: 'Center', rowHeight: height },
    });
  };

  for (const group of sheet.groups) {
    if (!wanted.has(group.camp.replace(/\s+/g, ' ').trim().toLowerCase())) continue;
    band(group.headerRow, STYLE.dividerFill, STYLE.dividerFont, STYLE.dividerRowHeight, `divider: ${group.camp}`);
  }

  if (sheet.totalRow !== undefined) {
    band(sheet.totalRow, STYLE.totalFill, STYLE.totalFont, STYLE.dividerRowHeight, 'total');
    operations.push({
      kind: 'border',
      address: rangeAddress(first, sheet.totalRow, last, sheet.totalRow),
      what: 'rule above the total',
      border: { edge: 'EdgeTop', style: 'Continuous', color: STYLE.headerFill, weight: 'Medium' },
    });
  }

  return operations;
}
