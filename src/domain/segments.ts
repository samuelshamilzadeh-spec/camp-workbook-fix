import {
  LAYOUT,
  MONTH_SEGMENTS,
  OTHER_SEGMENT,
  isSplitFamily,
  queueTabFor,
  type QueueFamily,
  type QueueSheetName,
  type Segment,
  type WorkbookLayout,
} from '../config';
import { toDateKey } from './adopt';

/**
 * Which monthly tab a row belongs on.
 *
 * The office asked to see `Missing Info` and `Ineligible & Inactive` one month
 * at a time. The month is the month of the VISIT — the daily sheet the row came
 * from — and that is the only date a queue row carries. It is also fixed: a
 * patient's visit date never changes, so a row never has to hop between monthly
 * tabs once it has landed, and the split adds no new way for a row to move.
 *
 * The keyword in column B still decides the FAMILY, exactly as before. This only
 * decides which of that family's tabs the row sits on, and for the three unsplit
 * families it decides nothing at all.
 */

/**
 * The visit date behind a value, from either representation the workbook uses.
 *
 * Two of them are in play and both have to work here. A row being appended
 * carries its source sheet's NAME (`July 30, 2026`), because that is what
 * `buildAppend` writes into `Date of Visit`. A row already sitting on a queue tab
 * carries an Excel SERIAL (46233), because `npm run migrate` converted the column
 * to real dates. Reading only one of them would put every row from the other
 * source into `Other`.
 */
export function visitDate(value: unknown, layout: WorkbookLayout = LAYOUT): Date | undefined {
  if (typeof value === 'string') {
    const named = layout.parseDailySheetDate(value);
    if (named) return named;
  }

  const key = toDateKey(value);
  if (!key) return undefined;

  const [year, month, day] = key.split('-').map(Number);
  if (!year || !month || !day) return undefined;
  return new Date(Date.UTC(year, month - 1, day));
}

/**
 * The segment a visit date falls in, or `Other`.
 *
 * `Other` is never a failure to handle a row, it is where a row goes. A visit in
 * May, or a `Date of Visit` cell that will not parse at all — the live workbook
 * has one — still ends up somewhere a person can see it. Returning undefined and
 * letting the caller decide is how a patient falls out of the system quietly.
 */
export function segmentForVisit(value: unknown, layout: WorkbookLayout = LAYOUT): Segment {
  const date = visitDate(value, layout);
  if (!date) return OTHER_SEGMENT;

  const month = date.getUTCMonth() + 1;
  return MONTH_SEGMENTS.find((entry) => entry.month === month)?.segment ?? OTHER_SEGMENT;
}

/**
 * The tab a row belongs on: its family, plus its month when that family is split.
 *
 * This is the one place the two halves of a destination come together, so a
 * caller that has a keyword and a visit date can never assemble the pair wrongly.
 */
export function queueSheetFor(
  family: QueueFamily,
  visit: unknown,
  layout: WorkbookLayout = LAYOUT,
): QueueSheetName {
  return isSplitFamily(family)
    ? queueTabFor(family, segmentForVisit(visit, layout))
    : queueTabFor(family, undefined);
}
