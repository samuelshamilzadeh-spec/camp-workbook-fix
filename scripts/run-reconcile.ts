/**
 * Runs one Phase 1 reconciliation cycle from the command line and prints the
 * plan. Same code path as the timer trigger, so what it prints is what the
 * Function would do.
 *
 *   node --env-file=.env node_modules/.bin/tsx scripts/run-reconcile.ts
 *   node --env-file=.env node_modules/.bin/tsx scripts/run-reconcile.ts --force
 *
 * `--force` clears the cached checkpoint first, so the run does real work even
 * when the file has not changed since the last cycle.
 */

import { consoleSink, describeError } from '../src/logging';
import { MemoryStateStore } from '../src/state/store';
import { createSyncContext } from '../src/sync/context';
import { runCycle } from '../src/sync/cycle';

async function main(): Promise<void> {
  const force = process.argv.includes('--force');
  const context = createSyncContext(consoleSink, { phase: 1, dryRun: true });

  const result = await runCycle({
    config: context.config,
    workbook: context.workbook,
    log: context.log,
    state: force ? new MemoryStateStore() : context.state,
  });

  if (result.status === 'skipped') {
    process.stdout.write(`skipped: ${result.reason} (${result.durationMs} ms)\n`);
    return;
  }

  const { plan } = result;
  process.stdout.write('\n--- Phase 1 plan ---\n');
  process.stdout.write(`sheets scanned: ${result.scannedSheets.length}\n`);
  for (const [kind, count] of Object.entries(plan.counts)) {
    process.stdout.write(`  ${kind}: ${count}\n`);
  }
  process.stdout.write(`  multi-match cells: ${plan.ambiguous.length}\n`);
  process.stdout.write(`  orphan queue rows: ${plan.orphans.length}\n`);
  process.stdout.write(`  unrecognized column B values: ${JSON.stringify(plan.unrecognizedCounts)}\n`);
  process.stdout.write(`applied: ${result.applied}\n`);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'reconcile.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
