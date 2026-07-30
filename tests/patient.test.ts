import { describe, expect, it } from 'vitest';
import { patientKey, resolveFieldAcrossVisits } from '../src/domain/patient';
import { sameValue } from '../src/domain/reconcile';

const resolve = (visits: Parameters<typeof resolveFieldAcrossVisits>[0]) =>
  resolveFieldAcrossVisits(visits, sameValue);

describe('patientKey', () => {
  const base = { 'Last Name': 'Weiss', 'First Name': 'Yehuda', 'Date of Birth': '2015-05-04' };

  it('matches the same person seen on different days', () => {
    // The real case from the live data: Weiss Yehuda, same DOB, two visits.
    expect(patientKey(base, 'aish')).toBe(patientKey(base, 'aish'));
  });

  it('ignores casing and stray whitespace', () => {
    expect(patientKey({ ...base, 'Last Name': '  WEISS ' }, 'Aish')).toBe(patientKey(base, 'aish'));
  });

  it('separates two children who share a name and a birthday but not a camp', () => {
    // The office's point: same name, same DOB, different camps is two children.
    expect(patientKey(base, 'aish')).not.toBe(patientKey(base, 'achim'));
  });

  it('separates two children who share a name but not a birthday', () => {
    // 14 name-pairs in the live workbook cover more than one date of birth.
    expect(patientKey(base, 'aish')).not.toBe(
      patientKey({ ...base, 'Date of Birth': '2011-01-01' }, 'aish'),
    );
  });

  it('refuses to build a key from incomplete identity', () => {
    // A duplicate is recoverable. Merging two children is not.
    expect(patientKey({ ...base, 'Date of Birth': '' }, 'aish')).toBeUndefined();
    expect(patientKey({ ...base, 'Last Name': '' }, 'aish')).toBeUndefined();
    expect(patientKey({ ...base, 'First Name': '   ' }, 'aish')).toBeUndefined();
  });

  it('still keys a patient whose camp is blank', () => {
    expect(patientKey(base, undefined)).toBeDefined();
  });
});

describe('resolveFieldAcrossVisits', () => {
  it('does nothing when no visit has been edited', () => {
    expect(
      resolve([
        { syncId: 'A', queueValue: '555', sourceValue: '555' },
        { syncId: 'B', queueValue: '555', sourceValue: '555' },
      ]),
    ).toEqual({ kind: 'no-change' });
  });

  it('propagates a staff edit to the patient\'s other visits', () => {
    // The workflow the office described: the info is filled in once on the
    // July 1 row, and the July 14 row should not have to be typed again.
    expect(
      resolve([
        { syncId: 'A', queueValue: '555-1234', sourceValue: '' },
        { syncId: 'B', queueValue: '', sourceValue: '' },
      ]),
    ).toEqual({ kind: 'propagate', value: '555-1234', from: 'A' });
  });

  it('propagates across three visits from a single edit', () => {
    expect(
      resolve([
        { syncId: 'A', queueValue: '', sourceValue: '' },
        { syncId: 'B', queueValue: 'aetna', sourceValue: '' },
        { syncId: 'C', queueValue: '', sourceValue: '' },
      ]),
    ).toMatchObject({ kind: 'propagate', value: 'aetna' });
  });

  it('treats the same edit made on two visits as one value, not a conflict', () => {
    expect(
      resolve([
        { syncId: 'A', queueValue: '555', sourceValue: '' },
        { syncId: 'B', queueValue: '555', sourceValue: '' },
      ]),
    ).toMatchObject({ kind: 'propagate', value: '555' });
  });

  it('REFUSES to choose when two visits were edited differently', () => {
    // 72 of the 148 repeat patients already disagree on some filled field,
    // phone number most often. Picking one silently discards the other.
    const result = resolve([
      { syncId: 'A', queueValue: '555-1111', sourceValue: '' },
      { syncId: 'B', queueValue: '555-2222', sourceValue: '' },
    ]);
    expect(result.kind).toBe('conflict');
    expect(result).toMatchObject({ editedBy: ['A', 'B'] });
  });

  it('leaves a pre-existing disagreement alone when nobody edited anything', () => {
    // Two visits recorded different phone numbers historically. That is not an
    // edit and must not trigger a write in either direction.
    expect(
      resolve([
        { syncId: 'A', queueValue: '555-1111', sourceValue: '555-1111' },
        { syncId: 'B', queueValue: '555-2222', sourceValue: '555-2222' },
      ]),
    ).toEqual({ kind: 'no-change' });
  });

  it('handles a single visit with no siblings', () => {
    expect(resolve([{ syncId: 'A', queueValue: 'x', sourceValue: '' }])).toMatchObject({
      kind: 'propagate',
      value: 'x',
    });
    expect(resolve([{ syncId: 'A', queueValue: 'x', sourceValue: 'x' }])).toEqual({
      kind: 'no-change',
    });
  });

  it('treats a cleared field as an edit worth propagating', () => {
    expect(
      resolve([
        { syncId: 'A', queueValue: '', sourceValue: '555' },
        { syncId: 'B', queueValue: '555', sourceValue: '555' },
      ]),
    ).toMatchObject({ kind: 'propagate', from: 'A' });
  });

  it('is idempotent: once propagated, the next pass sees no change', () => {
    const after = resolve([
      { syncId: 'A', queueValue: '555', sourceValue: '555' },
      { syncId: 'B', queueValue: '555', sourceValue: '555' },
    ]);
    expect(after).toEqual({ kind: 'no-change' });
  });
});

describe('filling blanks from a sibling visit', () => {
  const fill = (visits: Parameters<typeof resolveFieldAcrossVisits>[0]) =>
    resolveFieldAcrossVisits(visits, sameValue, { fillBlanks: true });

  it('fills a blank visit from one that knows the answer', () => {
    // Same patient: if July 14 has the phone and July 1 does not, July 1 gets
    // it. 67 cells in the live workbook are in this state.
    const result = fill([
      { syncId: 'A', queueValue: '', sourceValue: '' },
      { syncId: 'B', queueValue: '555-1234', sourceValue: '555-1234' },
    ]);
    expect(result).toMatchObject({ kind: 'fill-blanks', value: '555-1234', from: 'B' });
    expect(result).toMatchObject({ blankVisits: ['A'] });
  });

  it('fills several blanks at once', () => {
    expect(
      fill([
        { syncId: 'A', queueValue: '', sourceValue: '' },
        { syncId: 'B', queueValue: 'aetna', sourceValue: 'aetna' },
        { syncId: 'C', queueValue: '', sourceValue: '' },
      ]),
    ).toMatchObject({ kind: 'fill-blanks', blankVisits: ['A', 'C'] });
  });

  it('will not fill when two visits disagree about the answer', () => {
    // A historical disagreement is not an instruction, and choosing between
    // them would discard one. 72 repeat patients are in this state.
    expect(
      fill([
        { syncId: 'A', queueValue: '', sourceValue: '' },
        { syncId: 'B', queueValue: '555-1111', sourceValue: '555-1111' },
        { syncId: 'C', queueValue: '555-2222', sourceValue: '555-2222' },
      ]),
    ).toEqual({ kind: 'no-change' });
  });

  it('does nothing when there is nothing blank', () => {
    expect(
      fill([
        { syncId: 'A', queueValue: '555', sourceValue: '555' },
        { syncId: 'B', queueValue: '555', sourceValue: '555' },
      ]),
    ).toEqual({ kind: 'no-change' });
  });

  it('leaves a lone visit alone', () => {
    expect(fill([{ syncId: 'A', queueValue: '', sourceValue: '' }])).toEqual({ kind: 'no-change' });
  });

  it('lets a live staff edit win over back-filling', () => {
    // An edit is someone deciding now; a blank is just absence.
    expect(
      fill([
        { syncId: 'A', queueValue: '555-9999', sourceValue: '' },
        { syncId: 'B', queueValue: '555-1111', sourceValue: '555-1111' },
      ]),
    ).toMatchObject({ kind: 'propagate', value: '555-9999', from: 'A' });
  });

  it('is idempotent: once filled, the next pass does nothing', () => {
    expect(
      fill([
        { syncId: 'A', queueValue: '555', sourceValue: '555' },
        { syncId: 'B', queueValue: '555', sourceValue: '555' },
      ]),
    ).toEqual({ kind: 'no-change' });
  });

  it('stays off unless asked', () => {
    expect(
      resolveFieldAcrossVisits(
        [
          { syncId: 'A', queueValue: '', sourceValue: '' },
          { syncId: 'B', queueValue: '555', sourceValue: '555' },
        ],
        sameValue,
      ),
    ).toEqual({ kind: 'no-change' });
  });
});
