/**
 * Bounded-concurrency map.
 *
 * Reading ~60 sheets one after another at a few hundred milliseconds each does
 * not fit inside a five-second cycle. Reading all 60 at once invites throttling
 * and gives Graph a burst it will happily 429. A small fixed window is the
 * middle: the whole workbook is read in a couple of seconds, and there are never
 * more than `limit` requests outstanding.
 *
 * Results come back in input order regardless of completion order, because the
 * reconciler is order-sensitive for logging and diffing.
 */
export async function mapWithConcurrency<T, R>(
  items: readonly T[],
  limit: number,
  worker: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const effective = Math.max(1, Math.min(Math.floor(limit) || 1, items.length));
  const results = new Array<R>(items.length);
  let next = 0;

  async function run(): Promise<void> {
    for (;;) {
      const index = next++;
      if (index >= items.length) return;
      results[index] = await worker(items[index]!, index);
    }
  }

  // A rejection propagates and the remaining runners stop pulling work, which is
  // what we want: a failed read means this cycle's picture is incomplete, and
  // reconciling against an incomplete picture could remove queue rows whose
  // source sheet simply failed to load.
  await Promise.all(Array.from({ length: effective }, run));
  return results;
}
