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

describe('keyword forms found in the live workbook (verified 2026-07-30)', () => {
  // These are the forms staff actually type, taken from 4,102 rows of the real
  // sheet. The brief's vocabulary and the office's vocabulary are not the same.
  it('routes `dont accept`, which the brief never mentions and 214 rows use', () => {
    expect(classifyStatus('dont accept')).toMatchObject({ destination: 'Not Accepted' });
    expect(classifyStatus('dont accept- florida mcd')).toMatchObject({
      destination: 'Not Accepted',
    });
    expect(classifyStatus('dont bill')).toMatchObject({ destination: 'Not Accepted' });
  });

  it('still routes the brief\'s `not accepted`, which appears zero times', () => {
    expect(classifyStatus('not accepted')).toMatchObject({ destination: 'Not Accepted' });
  });

  // Superseded by the office 2026-07-30: united refuah is its own category with
  // its own sheet, not a terminal EMR. Covered in "office decisions" below.

  it('catches the observed misspelling of ineligible', () => {
    expect(classifyStatus('inegilible')).toMatchObject({
      destination: 'Ineligible & Inactive',
    });
  });

  // The insurance wording and `skip` were deliberately unrouted until the
  // office decided. They now have answers — see "office decisions" below.
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

describe('needs ohi / needs lasante — the negation trap', () => {
  // 1,268 rows. The brief says ohi or lasante ANYWHERE means terminal, which
  // marks every one of these as already entered into an EMR when they are
  // precisely the rows that still need entering.
  it('treats `needs ohi` as pending, not done', () => {
    expect(classifyStatus('needs ohi').kind).toBe('pending');
    expect(classifyStatus('need ohi').kind).toBe('pending');
    expect(classifyStatus('needs ohi- put in on 7/27').kind).toBe('pending');
    expect(classifyStatus('needs ohi- already an established pt').kind).toBe('pending');
  });

  it('treats `needs lasante` as pending, not done', () => {
    expect(classifyStatus('needs lasante').kind).toBe('pending');
    expect(classifyStatus('needs lasante- under esti').kind).toBe('pending');
    expect(classifyStatus('needs lasante put in on 7/22').kind).toBe('pending');
    expect(classifyStatus('needs lasante- uunder esti').kind).toBe('pending');
  });

  it('still treats a bare or suffixed ohi/lasante as done', () => {
    // The suffix rule from the brief must survive: `ohi - esti` is still an OHI.
    for (const value of ['ohi', 'ohi -pa', 'ohi on6/30', 'ohi - need a good cc',
                         'lasante-e', 'lasante-o', 'lasante-e -pa',
                         'lasante under esti', 'lasante-o- on tuesday', 'oop-lasante']) {
      expect(classifyStatus(value).kind).toBe('terminal');
    }
  });

  it('catches the observed misspelling of lasante', () => {
    expect(classifyStatus('lasanante- e').kind).toBe('terminal');
  });

  it('produces no queue row either way, but the two are distinguishable', () => {
    // Both are non-queued, which is why the bug was invisible. The kinds differ
    // so outstanding work can still be counted.
    expect(classifyStatus('needs ohi').kind).not.toBe(classifyStatus('ohi').kind);
  });
});

describe('office decisions, 2026-07-30', () => {
  it('leaves `skip` alone entirely — a non-billable follow-up visit', () => {
    expect(classifyStatus('skip').kind).toBe('ignored');
  });

  // The eleven insurance phrasings, assigned individually by the office. They
  // are NOT one group — they split four ways, and several contain one another,
  // so the table order in QUEUE_KEYWORDS is what keeps them apart. These cases
  // exist to fail loudly if that order is ever disturbed.
  it.each([
    // "we need the information"
    ['no insurance on file', 'Missing Info'],
    ['need insurance', 'Missing Info'],
    ['need ins', 'Missing Info'],
    // genuinely uninsured, cannot bill
    ['no insurance', 'Not Accepted'],
    ['wrote paper has no ins', 'Not Accepted'],
    ['doesnt have ins', 'Not Accepted'],
    ['doesnt have insurance', 'Not Accepted'],
    ['pt doesnt have insurance', 'Not Accepted'],
    // insured, but the details need checking or fixing
    ['need insurance verification', 'Verify Insurance'],
    ['invalid ins', 'Ineligible & Inactive'],
    ['incorrect insurance', 'Ineligible & Inactive'],
  ])('routes %s to %s', (value, destination) => {
    expect(classifyStatus(value)).toMatchObject({ destination });
  });

  it('keeps the overlapping insurance phrases apart', () => {
    // Each pair differs only by a substring. If the general rule were listed
    // first it would swallow the specific one, and the office's distinction
    // would vanish without any error.
    expect(classifyStatus('no insurance on file')).toMatchObject({ destination: 'Missing Info' });
    expect(classifyStatus('no insurance')).toMatchObject({ destination: 'Not Accepted' });

    expect(classifyStatus('need insurance verification')).toMatchObject({
      destination: 'Verify Insurance',
    });
    expect(classifyStatus('need insurance')).toMatchObject({ destination: 'Missing Info' });
  });

  it('routes campium/campflow wording to Missing Info', () => {
    expect(classifyStatus('not on campium')).toMatchObject({ destination: 'Missing Info' });
    expect(classifyStatus('not on campflow')).toMatchObject({ destination: 'Missing Info' });
  });

  it('sends united refuah to its own sheet rather than treating it as done', () => {
    expect(classifyStatus('united refuah')).toMatchObject({ destination: 'United Refuah' });
    expect(classifyStatus('united refuah- 2nd visit from 6/30')).toMatchObject({
      destination: 'United Refuah',
    });
  });

  it('keeps `verify insurance` ahead of the insurance wording rules', () => {
    // `need insurance verification` contains neither `verify insurance` nor a
    // no-insurance phrase outright; make sure the common case is unaffected.
    expect(classifyStatus('verify insurance')).toMatchObject({ destination: 'Verify Insurance' });
    expect(classifyStatus('verify insurance and date of birth')).toMatchObject({
      destination: 'Verify Insurance',
    });
    expect(classifyStatus('verify insurance/dob')).toMatchObject({
      destination: 'Verify Insurance',
    });
  });
});

describe('multi-match reporting stays useful', () => {
  // The overlapping insurance rules fire constantly by design. Reporting each
  // one as a data problem would emit the same warnings on every cycle forever,
  // and a channel that always cries wolf is one nobody reads — burying the real
  // multi-matches it exists to surface.
  it('stays quiet when the specific rule beat the general one by design', () => {
    expect(isAmbiguous(classifyStatus('no insurance on file'))).toBe(false);
    expect(isAmbiguous(classifyStatus('need insurance verification'))).toBe(false);
    expect(isAmbiguous(classifyStatus('pt doesnt have insurance'))).toBe(false);
  });

  it('still flags genuinely conflicting keywords', () => {
    // Two unrelated statuses in one cell is a human error worth looking at.
    expect(isAmbiguous(classifyStatus('missing info and verify insurance'))).toBe(true);
    expect(isAmbiguous(classifyStatus('dont accept but also missing info'))).toBe(true);
  });

  it('still flags a terminal keyword sitting next to a queued one', () => {
    expect(isAmbiguous(classifyStatus('ohi but also missing info'))).toBe(true);
  });
});
