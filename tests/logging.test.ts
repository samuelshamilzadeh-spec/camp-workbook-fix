import { describe, expect, it } from 'vitest';
import { createLogger, redactPayload, type LogSink } from '../src/logging';

function capture(): { sink: LogSink; lines: string[] } {
  const lines: string[] = [];
  const push = (m: string) => lines.push(m);
  return {
    lines,
    sink: { debug: push, info: push, warn: push, error: push },
  };
}

describe('redaction', () => {
  it('keeps identifiers that are safe to log', () => {
    expect(redactPayload({ syncId: 'S0001', sheet: '2026-07-30', row: 12 })).toEqual({
      syncId: 'S0001',
      sheet: '2026-07-30',
      row: 12,
    });
  });

  it('redacts patient fields by key, whatever they are called at the call site', () => {
    const out = redactPayload({
      lastName: 'Smith',
      dateOfBirth: '2010-01-01',
      phone: '555-1234',
      insuranceCarrier: 'Acme',
      billingAddress: '1 Main St',
      value: 'anything',
    });

    for (const redacted of Object.values(out)) {
      expect(String(redacted)).toMatch(/^\[redacted:/);
    }
    expect(JSON.stringify(out)).not.toContain('Smith');
    expect(JSON.stringify(out)).not.toContain('555');
  });

  it('redacts nested objects rather than trusting the top-level key', () => {
    const out = redactPayload({
      counts: { 'append-queue-row': 3 },
      intent: { syncId: 'S0001', values: { 'Last Name': 'Smith' } },
    });
    expect(JSON.stringify(out)).not.toContain('Smith');
  });

  it('does not let a cell value ride along inside a log message', () => {
    const { sink, lines } = capture();
    const log = createLogger(sink);
    log.warn('status.multi_match', {
      sheet: '2026-07-30',
      row: 4,
      matched: ['ohi', 'missing info'],
      statusRaw: 'ohi - Smith called',
    });

    expect(lines[0]).toContain('missing info');
    expect(lines[0]).not.toContain('Smith');
  });
});
