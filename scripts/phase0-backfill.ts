/**
 * Phase 0: stamp SyncIDs.
 *
 * Adds the `SyncID` header to every sheet it touches and stamps an ID on
 * *only* the rows currently carrying a queued keyword in column B — roughly the
 * population sitting on the existing mirror sheets today. It does not backfill
 * the workbook.
 *
 * That restriction is a correctness decision, not just a cost one. The tempting
 * alternative — "stamp any row whose identifying fields are populated" — fails
 * on exactly the population this system serves, because a Missing Info row is by
 * definition missing a DOB or a phone or an address. A recognized keyword in
 * column B is the reliable signal that a human considered this a real patient
 * row.
 *
 * Run against a COPY first:
 *
 *   node --env-file=.env node_modules/.bin/tsx scripts/phase0-backfill.ts --dry-run
 *   node --env-file=.env node_modules/.bin/tsx scripts/phase0-backfill.ts --apply
 *
 * `--dry-run` is the default. `--apply` must be passed explicitly and, on top of
 * that, requires SYNC_LAYOUT_VERIFIED=true.
 */

import { LAYOUT, assertLayoutVerified, loadConfig } from '../src/config';
import { planSheetStamps, stampWriteValues, type SheetStampPlan } from '../src/domain/backfill';
import { rangeAddress } from '../src/domain/cells';
import { isDailySheetName } from '../src/domain/dailySheets';
import { createLogger, consoleSink, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';

function stampAddress(plan: SheetStampPlan): string {
  return rangeAddress(
    LAYOUT.daily.syncIdColumn,
    plan.firstRow,
    LAYOUT.daily.syncIdColumn,
    plan.lastRow,
  );
}

async function main(): Promise<void> {
  const args = new Set(process.argv.slice(2));
  const apply = args.has('--apply');
  const dryRun = !apply;
  const limitToSheets = readSheetFilter(process.argv.slice(2));

  const log = createLogger(consoleSink, { phase: 0, dryRun });
  const config = loadConfig();
  const { workbook } = createSyncContext(consoleSink, { phase: 0, dryRun });

  if (apply) {
    assertLayoutVerified(config);
    log.warn('backfill.apply_mode', {
      reason: 'writes enabled; confirm this is a copy of the workbook, not the live file',
    });
  }

  await workbook.createSession(apply);

  try {
    const worksheets = await workbook.listWorksheets();
    const targets = worksheets
      .map((sheet) => sheet.name)
      .filter((name) => isDailySheetName(name))
      .filter((name) => limitToSheets === undefined || limitToSheets.includes(name));

    log.info('backfill.sheets_selected', { count: targets.length, sheets: targets });

    const plans: SheetStampPlan[] = [];
    for (const sheet of targets) {
      const used = await workbook.getUsedRange(sheet);
      plans.push(planSheetStamps(sheet, used));
    }

    const totalStamps = plans.reduce((sum, plan) => sum + plan.newStamps, 0);
    const totalExisting = plans.reduce((sum, plan) => sum + plan.alreadyStamped, 0);

    // The brief's requirement: report how many rows it would stamp before it
    // stamps any.
    for (const plan of plans) {
      if (plan.newStamps === 0 && !plan.headerNeeded) continue;
      log.info('backfill.sheet_plan', {
        sheet: plan.sheet,
        count: plan.newStamps,
        skipped: plan.alreadyStamped,
        address: plan.newStamps > 0 ? stampAddress(plan) : undefined,
      });
    }

    log.info('backfill.summary', {
      count: totalStamps,
      skipped: totalExisting,
      columns: [LAYOUT.daily.syncIdColumn],
      dryRun,
    });

    if (dryRun) {
      log.info('backfill.dry_run_complete', {
        reason: 'no writes performed; re-run with --apply to stamp',
      });
      return;
    }

    for (const plan of plans) {
      if (plan.headerNeeded) {
        await workbook.writeRange(
          plan.sheet,
          `${LAYOUT.daily.syncIdColumn}${LAYOUT.daily.headerRow}`,
          [['SyncID']],
        );
      }
      if (plan.newStamps === 0) continue;

      // One range write per sheet, never a cell at a time. Rows inside the block
      // that already have an ID are written back unchanged, so an ID is still
      // only ever assigned once.
      const address = stampAddress(plan);
      await workbook.writeRange(plan.sheet, address, stampWriteValues(plan));
      log.info('backfill.sheet_written', {
        sheet: plan.sheet,
        count: plan.newStamps,
        address,
      });
    }

    log.info('backfill.complete', { count: totalStamps });
  } catch (error) {
    log.error('backfill.failed', describeError(error));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

function readSheetFilter(argv: string[]): string[] | undefined {
  const flag = argv.find((arg) => arg.startsWith('--sheets='));
  if (!flag) return undefined;
  return flag
    .slice('--sheets='.length)
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'backfill.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
