/**
 * Concurrency test. Run this BEFORE Phase 2, because the result decides how
 * aggressive the write strategy can be.
 *
 * The question it answers: while two people have the workbook open — one in
 * Excel desktop with autosave on, one in Excel Online — what actually happens
 * when a third party writes to it through Graph? The documentation says 423 and
 * 409; what matters is the real rate, the real latency, and above all whether a
 * human's in-flight edit is ever lost.
 *
 * RUN THIS AGAINST A COPY. It writes to a scratch cell repeatedly.
 *
 *   node --env-file=.env node_modules/.bin/tsx scripts/concurrency-test.ts \
 *     --sheet="2026-07-30" --cell=BZ1 --iterations=60 --interval=5000
 *
 * Procedure:
 *   1. Take a copy of the workbook into the same SharePoint site. Note its
 *      deletion date now, not later — it holds patient data.
 *   2. Person A opens the copy in Excel desktop, autosave on, and types into a
 *      few cells throughout the run, including a cell adjacent to --cell.
 *   3. Person B opens the copy in Excel Online and does the same.
 *   4. Run this script.
 *   5. Afterwards, check the workbook: did every edit A and B made survive?
 *      That answer is not visible from this script's output and is the single
 *      most important result.
 */

import { writeFile } from 'node:fs/promises';
import { GraphError } from '../src/graph/client';
import { consoleSink, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';

interface Attempt {
  iteration: number;
  startedAt: string;
  durationMs: number;
  ok: boolean;
  statusCode?: number;
  graphCode?: string;
  readBackMatched?: boolean;
}

function arg(name: string, fallback: string): string {
  const found = process.argv.find((a) => a.startsWith(`--${name}=`));
  return found ? found.slice(name.length + 3) : fallback;
}

async function main(): Promise<void> {
  const sheet = arg('sheet', '');
  const cell = arg('cell', 'BZ1');
  const iterations = Number(arg('iterations', '60'));
  const intervalMs = Number(arg('interval', '5000'));
  const outputPath = arg('out', 'concurrency-results.json');

  if (!sheet) {
    throw new Error('--sheet is required, e.g. --sheet="2026-07-30"');
  }

  const { workbook, log } = createSyncContext(consoleSink, { phase: 0, dryRun: false });

  log.warn('concurrency.start', {
    sheet,
    address: cell,
    count: iterations,
    reason: 'writes to a scratch cell; confirm this is a copy, not the live workbook',
  });

  const attempts: Attempt[] = [];

  for (let i = 1; i <= iterations; i++) {
    // A marker with no patient data in it. The iteration number and a timestamp
    // are enough to tell which write won.
    const marker = `sync-test ${i} ${new Date().toISOString()}`;
    const startedAt = Date.now();

    const attempt: Attempt = {
      iteration: i,
      startedAt: new Date(startedAt).toISOString(),
      durationMs: 0,
      ok: false,
    };

    try {
      // maxAttempts 1: the point is to observe the raw failure rate, not to let
      // the client's retry policy hide it.
      await workbook.writeRange(sheet, cell, [[marker]]);
      attempt.ok = true;

      const readBack = await workbook.getRange(sheet, cell);
      attempt.readBackMatched = String(readBack.values?.[0]?.[0] ?? '') === marker;
    } catch (error) {
      if (error instanceof GraphError) {
        attempt.statusCode = error.statusCode;
        attempt.graphCode = error.graphCode;
      } else {
        attempt.graphCode = 'non-graph-error';
      }
      log.warn('concurrency.write_failed', { attempt: i, ...describeError(error) });
    }

    attempt.durationMs = Date.now() - startedAt;
    attempts.push(attempt);

    process.stdout.write(
      `${String(i).padStart(3)}/${iterations}  ` +
        `${attempt.ok ? 'ok  ' : `FAIL ${attempt.statusCode ?? '-'}`}  ` +
        `${String(attempt.durationMs).padStart(5)} ms  ` +
        `readback=${attempt.readBackMatched ?? '-'}\n`,
    );

    if (i < iterations) await sleep(intervalMs);
  }

  const failures = attempts.filter((a) => !a.ok);
  const locked = failures.filter((a) => a.statusCode === 423 || a.statusCode === 409);
  const throttled = failures.filter((a) => a.statusCode === 429);
  const durations = attempts.map((a) => a.durationMs).sort((x, y) => x - y);

  const summary = {
    sheet,
    cell,
    iterations,
    intervalMs,
    succeeded: attempts.length - failures.length,
    failed: failures.length,
    lockedOrConflict: locked.length,
    throttled: throttled.length,
    readBackMismatches: attempts.filter((a) => a.ok && a.readBackMatched === false).length,
    latencyMs: {
      p50: percentile(durations, 0.5),
      p95: percentile(durations, 0.95),
      max: durations.at(-1) ?? 0,
    },
  };

  process.stdout.write(`\n${JSON.stringify(summary, null, 2)}\n`);
  await writeFile(outputPath, JSON.stringify({ summary, attempts }, null, 2), 'utf8');

  process.stdout.write(
    `\nWrote ${outputPath}.\n` +
      'Now check the workbook by hand: did every edit the two humans made survive? ' +
      'A clean run here with a lost human edit is a failed test.\n',
  );
}

function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  const index = Math.min(sorted.length - 1, Math.floor(sorted.length * p));
  return sorted[index]!;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'concurrency.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
