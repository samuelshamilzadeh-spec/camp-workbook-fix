import { describe, expect, it } from 'vitest';
import { mapWithConcurrency } from '../src/domain/concurrency';

const tick = () => new Promise((resolve) => setTimeout(resolve, 1));

describe('mapWithConcurrency', () => {
  it('returns results in input order, not completion order', async () => {
    const items = [30, 1, 20, 2, 10];
    const result = await mapWithConcurrency(items, 3, async (ms) => {
      await new Promise((resolve) => setTimeout(resolve, ms));
      return ms;
    });
    expect(result).toEqual(items);
  });

  it('never exceeds the concurrency limit', async () => {
    let inFlight = 0;
    let peak = 0;

    await mapWithConcurrency(Array.from({ length: 60 }, (_, i) => i), 8, async (i) => {
      inFlight++;
      peak = Math.max(peak, inFlight);
      await tick();
      inFlight--;
      return i;
    });

    expect(peak).toBeLessThanOrEqual(8);
    expect(peak).toBeGreaterThan(1);
  });

  it('processes every item', async () => {
    const seen: number[] = [];
    await mapWithConcurrency(Array.from({ length: 60 }, (_, i) => i), 8, async (i) => {
      await tick();
      seen.push(i);
      return i;
    });
    expect(seen.sort((a, b) => a - b)).toEqual(Array.from({ length: 60 }, (_, i) => i));
  });

  it('propagates a failure rather than returning a partial picture', async () => {
    // Reconciling against a half-read workbook could remove queue rows whose
    // source sheet merely failed to load, so a failed read must fail the cycle.
    await expect(
      mapWithConcurrency([1, 2, 3, 4], 2, async (i) => {
        if (i === 3) throw new Error('423 Locked');
        return i;
      }),
    ).rejects.toThrow('423 Locked');
  });

  it('handles an empty list', async () => {
    expect(await mapWithConcurrency([], 8, async () => 1)).toEqual([]);
  });

  it('clamps a nonsense limit instead of hanging', async () => {
    expect(await mapWithConcurrency([1, 2, 3], 0, async (i) => i * 2)).toEqual([2, 4, 6]);
    expect(await mapWithConcurrency([1, 2, 3], -5, async (i) => i * 2)).toEqual([2, 4, 6]);
  });
});
