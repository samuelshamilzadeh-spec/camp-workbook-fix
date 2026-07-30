import { describe, expect, it } from 'vitest';
import {
  classifyStatus,
  isAmbiguous,
  normalizeStatus,
} from '../src/domain/status';

describe('normalizeStatus', () => {
  it('lowercases, trims and collapses whitespace', () => {
    expect(normalizeStatus('  Missing   Info  ')).toBe('missing info');
  });

  it('normalizes non-breaking spaces and unicode dashes', () => {
    expect(normalizeStatus('OHI – esti')).toBe('ohi - esti');
  });

  it('treats blank-ish values as empty', () => {
    expect(normalizeStatus(null)).toBe('');
    expect(normalizeStatus(undefined)).toBe('');
    expect(normalizeStatus('   ')).toBe('');
  });
});

describe('classifyStatus', () => {
  it('matches on substring, not equality', () => {
    const outcome = classifyStatus('missing info - called 7/29, no answer');
    expect(outcome).toMatchObject({ kind: 'queued', destination: 'Missing Info' });
  });

  it('is case insensitive', () => {
    expect(classifyStatus('NOT ACCEPTED')).toMatchObject({
      kind: 'queued',
      destination: 'Not Accepted',
    });
  });

  it('treats ohi with a suffix as terminal', () => {
    // The brief's example: `ohi - esti` is still an OHI.
    expect(classifyStatus('ohi - esti').kind).toBe('terminal');
    expect(classifyStatus('OHI 7/29').kind).toBe('terminal');
    expect(classifyStatus('lasante entered').kind).toBe('terminal');
  });

  it('lets a terminal keyword suppress a queued one anywhere in the value', () => {
    const outcome = classifyStatus('was missing info, now ohi');
    expect(outcome.kind).toBe('terminal');
  });

  it('routes ineligible and inactive independently to the same sheet', () => {
    expect(classifyStatus('ineligible')).toMatchObject({
      destination: 'Ineligible & Inactive',
    });
    expect(classifyStatus('inactive')).toMatchObject({
      destination: 'Ineligible & Inactive',
    });
    expect(classifyStatus('policy inactive as of 6/1')).toMatchObject({
      destination: 'Ineligible & Inactive',
    });
  });

  it('does not require the combined ineligible/inactive string', () => {
    // That string is the sheet's name, not a value staff type.
    expect(classifyStatus('ineligible/inactive')).toMatchObject({
      destination: 'Ineligible & Inactive',
    });
  });

  it('matches verify insurance', () => {
    expect(classifyStatus('verify insurance w/ mom')).toMatchObject({
      destination: 'Verify Insurance',
    });
  });

  it('never guesses at an unrecognized value', () => {
    expect(classifyStatus('call back tuesday').kind).toBe('unrecognized');
    expect(classifyStatus('???').kind).toBe('unrecognized');
  });

  it('reports blank separately from unrecognized', () => {
    expect(classifyStatus('').kind).toBe('blank');
  });
});

describe('isAmbiguous', () => {
  it('does not flag ineligible/inactive, which share a destination', () => {
    const outcome = classifyStatus('ineligible / inactive');
    expect(outcome.kind).toBe('queued');
    expect(isAmbiguous(outcome)).toBe(false);
  });

  it('flags two queued keywords with different destinations', () => {
    const outcome = classifyStatus('missing info and verify insurance');
    expect(isAmbiguous(outcome)).toBe(true);
  });

  it('flags a terminal keyword sitting alongside a queued one', () => {
    const outcome = classifyStatus('ohi but also missing info');
    expect(outcome.kind).toBe('terminal');
    expect(isAmbiguous(outcome)).toBe(true);
  });

  it('does not flag a single match', () => {
    expect(isAmbiguous(classifyStatus('not accepted'))).toBe(false);
  });
});
