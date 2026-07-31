import { describe, expect, it } from 'vitest';
import { comparisonKey, sameMeaning } from '../src/domain/compare';

/**
 * Every pair below was taken from the live workbook, where a plain string
 * comparison called it a staff edit. None of them is one.
 */
describe('values the mirror sheets spelled differently', () => {
  it('does not call a case difference an edit (770 of these)', () => {
    expect(sameMeaning('State', 'LAKEWOOD', 'Lakewood')).toBe(true);
    expect(sameMeaning('City', 'Southfield', 'southfield')).toBe(true);
    expect(sameMeaning('State', 'MI', 'mi')).toBe(true);
  });

  it('does not call phone formatting an edit (369 of these)', () => {
    expect(sameMeaning('Phone Number', '(248) 331-3973', '248-331-3973')).toBe(true);
    expect(sameMeaning('Phone Number', '(845) 659-2658', '845-659-2658')).toBe(true);
    expect(sameMeaning('Phone Number', '845.825.7080', '(845) 825-7080')).toBe(true);
  });

  it('ignores a US country code', () => {
    // Live pair: the daily sheet held the eleven-digit form.
    expect(sameMeaning('Phone Number', '(718) 812-5386', '17188125386')).toBe(true);
  });

  it('restores the leading zero Excel stripped from a zip code', () => {
    // THE dangerous one. 08701 is Lakewood, New Jersey; the queue cell is
    // numeric so it arrives as 8701. Writing that back would have replaced a
    // real zip code with a four-digit number on the sheet that gets billed from.
    expect(sameMeaning('Zip Code', 8701, '08701')).toBe(true);
    expect(sameMeaning('Zip Code', '8527', '08527')).toBe(true);
    expect(comparisonKey('Zip Code', 8701)).toBe('08701');
  });

  it('reads a zip+4 by its first five digits', () => {
    expect(sameMeaning('Zip Code', '08701-1234', '08701')).toBe(true);
  });

  it('treats a date as the day it denotes, however it is stored', () => {
    expect(sameMeaning('Date of Birth', 40147, '11/30/2009')).toBe(true);
    expect(sameMeaning('Date of Visit', 46227, '07/24/2026')).toBe(true);
  });

  it('collapses whitespace in a name or an address', () => {
    expect(sameMeaning('Billing Address', '12  Maple  St ', '12 Maple St')).toBe(true);
    expect(sameMeaning('Last Name', 'halperin', 'Halperin')).toBe(true);
  });
});

describe('edits that are real, and must still get through', () => {
  it('sees a genuinely different phone number', () => {
    expect(sameMeaning('Phone Number', '(845) 111-2222', '845-333-4444')).toBe(false);
  });

  it('sees a genuinely different zip code', () => {
    expect(sameMeaning('Zip Code', '08701', '08527')).toBe(false);
  });

  it('sees a genuinely different town', () => {
    expect(sameMeaning('City', 'Lakewood', 'Monsey')).toBe(false);
  });

  it('sees a value filled in where there was none', () => {
    // The whole point of the system: somebody supplied the missing phone.
    expect(sameMeaning('Phone Number', '845-555-1234', '')).toBe(false);
    expect(sameMeaning('Phone Number', '845-555-1234', null)).toBe(false);
  });

  it('counts blank on both sides as agreement', () => {
    expect(sameMeaning('Phone Number', '', null)).toBe(true);
    expect(sameMeaning('City', '   ', undefined)).toBe(true);
  });

  it('does not let two different numbers collide through zip padding', () => {
    expect(sameMeaning('Zip Code', '1', '2')).toBe(false);
  });

  it('does not read a non-date column as a date', () => {
    // 10952 is a Monsey zip. As an Excel serial it is a day in 1999, and the
    // date-aware path applied to every column would have made unequal values
    // compare equal.
    expect(sameMeaning('Zip Code', '10952', '10953')).toBe(false);
    expect(sameMeaning('Insurance ID #', '40147', '40148')).toBe(false);
  });
});
