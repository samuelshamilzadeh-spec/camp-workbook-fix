import { app, type InvocationContext, type Timer } from '@azure/functions';
import { describeError, type LogSink } from '../logging';
import { getSyncContext } from '../sync/context';
import { runCycle } from '../sync/cycle';

/**
 * Timer trigger, every 5 seconds.
 *
 * Consumption plan (Y1), deliberately not Flex Consumption: Flex bills a minimum
 * of 1,000 ms per execution at a larger default instance size, which turns this
 * cheap high-frequency polling pattern into a real bill. Consumption bills a
 * minimum of 128 MB x 100 ms and this workload fits inside the free monthly
 * grant. See the README before changing the hosting plan.
 *
 * `runOnStartup` stays false: a cold start firing a cycle before the schedule
 * settles is a small extra cost with no benefit, since the next tick is 5
 * seconds away.
 */

const SCHEDULE = process.env.SYNC_TIMER_SCHEDULE ?? '*/5 * * * * *';

export async function syncTimer(timer: Timer, context: InvocationContext): Promise<void> {
  const sink: LogSink = {
    debug: (m) => context.debug(m),
    info: (m) => context.info(m),
    warn: (m) => context.warn(m),
    error: (m) => context.error(m),
  };

  if (timer.isPastDue) {
    // Expected after a scale-in or a restart. State-based reconciliation makes
    // catch-up identical to the normal path, so there is nothing special to do.
    context.debug(JSON.stringify({ event: 'timer.past_due' }));
  }

  const deps = getSyncContext(sink);

  try {
    const result = await runCycle(deps);
    if (result.status === 'planned') {
      deps.log.info('cycle.complete', {
        count: result.plan.intents.length,
        counts: result.plan.counts,
        durationMs: result.durationMs,
        wouldWrite: result.applied,
      });
    }
  } catch (error) {
    // Throwing marks the invocation failed, which is what the health signal in
    // Application Insights should show. Exceptions are excluded from sampling in
    // host.json so none of these are lost.
    deps.log.error('cycle.error', describeError(error));
    throw error;
  }
}

app.timer('syncTimer', {
  schedule: SCHEDULE,
  runOnStartup: false,
  useMonitor: true,
  handler: syncTimer,
});
