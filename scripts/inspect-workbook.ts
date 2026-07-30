/**
 * Answers the "Before you write code" questions against the real workbook, so
 * the LAYOUT block in src/config.ts can be corrected before anything writes.
 *
 *   node --env-file=.env node_modules/.bin/tsx scripts/inspect-workbook.ts
 *
 * Read-only: it opens a non-persisted session and never calls a write method.
 *
 * It prints structure and never patient data. Where it has to characterise cell
 * contents — the distinct values in column B, for instance — it prints the
 * normalized keyword forms and counts, not the raw cells, because a status cell
 * can contain a note a staff member typed. Camp names are printed, because
 * confirming whether they need normalizing is one of the questions and a camp is
 * an organisation rather than a patient.
 */

import { LAYOUT, QUEUE_SHEET_NAMES } from '../src/config';
import { columnToIndex, indexToColumn, parseAddress } from '../src/domain/cells';
import { isDailySheetName, normalizeCamp } from '../src/domain/dailySheets';
import { normalizeStatus, TERMINAL_KEYWORDS, QUEUE_KEYWORDS } from '../src/domain/status';
import { consoleSink, describeError } from '../src/logging';
import { createSyncContext } from '../src/sync/context';

const SAMPLE_SHEETS = Number(process.env.INSPECT_SHEETS ?? '5');

function out(line: string): void {
  process.stdout.write(`${line}\n`);
}

async function main(): Promise<void> {
  const { workbook } = createSyncContext(consoleSink, { phase: 1, dryRun: true });
  await workbook.createSession(false);

  try {
    const metadata = await workbook.getFileMetadata();
    out('# Workbook inspection');
    out(`lastModifiedDateTime: ${metadata.lastModifiedDateTime}`);
    out('');

    // --- Q1: sheet naming convention ---------------------------------------
    const worksheets = await workbook.listWorksheets();
    out('## 1. Worksheets');
    out(`total: ${worksheets.length}`);
    for (const sheet of worksheets) {
      const kind = isDailySheetName(sheet.name)
        ? 'daily?'
        : QUEUE_SHEET_NAMES.includes(sheet.name as never)
          ? 'queue'
          : 'other';
      const parsed = LAYOUT.parseDailySheetDate(sheet.name);
      out(
        `  [${kind.padEnd(6)}] ${sheet.name}` +
          `  (${sheet.visibility}${parsed ? `, parsed ${parsed.toISOString().slice(0, 10)}` : ''})`,
      );
    }
    out('');
    out(
      'CHECK: does every daily sheet show as "daily?" and parse to a date? ' +
        'If not, fix LAYOUT.dailySheetPattern / parseDailySheetDate.',
    );
    out('');

    // --- Q5: named ranges ---------------------------------------------------
    const names = await workbook.listNames().catch(() => []);
    out('## 5. Named ranges (SyncID column must collide with none of these)');
    if (names.length === 0) out('  none');
    for (const name of names) {
      out(`  ${name.name} [${name.scope}] -> ${name.value}`);
    }
    out('');

    // --- Q2, Q3, Q4: per-sheet shape ---------------------------------------
    const dailySheets = worksheets
      .map((sheet) => sheet.name)
      .filter((name) => isDailySheetName(name))
      .slice(-SAMPLE_SHEETS);

    const statusForms = new Map<string, number>();
    const camps = new Map<string, number>();
    let widestColumn = 0;

    out(`## 2-4. Daily sheet shape (last ${dailySheets.length} sheets)`);
    for (const sheetName of dailySheets) {
      const used = await workbook.getUsedRange(sheetName);
      const { startRow, startColumn, endColumn, endRow } = parseAddress(used.address);
      widestColumn = Math.max(widestColumn, columnToIndex(endColumn));

      out(`  ${sheetName}: usedRange ${used.address} (${used.rowCount} rows)`);

      // Header row detection: the first row in the used range with three or more
      // short non-numeric cells looks like a header.
      const headerGuess = guessHeaderRow(used.values, startRow);
      out(
        `    header row guess: ${headerGuess ?? 'none'}` +
          ` (config says ${LAYOUT.daily.headerRow}, data from ${LAYOUT.daily.firstDataRow})`,
      );
      if (headerGuess !== null) {
        const headerCells = used.values[headerGuess - startRow] ?? [];
        headerCells.forEach((cell, index) => {
          const text = String(cell ?? '').trim();
          if (text) out(`      ${indexToColumn(columnToIndex(startColumn) + index)}: ${text}`);
        });
      }

      const statusIndex = columnToIndex(LAYOUT.daily.statusColumn) - columnToIndex(startColumn);
      const campIndex = columnToIndex(LAYOUT.daily.campColumn) - columnToIndex(startColumn);

      for (let i = 0; i < used.values.length; i++) {
        const row = used.values[i] ?? [];
        const normalized = normalizeStatus(row[statusIndex]);
        if (normalized) {
          const keyword = summarizeStatus(normalized);
          statusForms.set(keyword, (statusForms.get(keyword) ?? 0) + 1);
        }
        const camp = normalizeCamp(row[campIndex]);
        if (camp) camps.set(camp, (camps.get(camp) ?? 0) + 1);
      }
      out(`    last used column: ${endColumn} (row ${endRow})`);
    }
    out('');

    out('## 3. Column B keyword forms found (normalized; raw values not printed)');
    for (const [form, count] of [...statusForms].sort((a, b) => b[1] - a[1])) {
      out(`  ${String(count).padStart(5)}  ${form}`);
    }
    out('');
    out(
      'CHECK: every "UNRECOGNIZED" line is a keyword variant the matcher misses. ' +
        'Any patient behind one of those is silently unqueued.',
    );
    out('');

    out(`## Camp values in ${LAYOUT.daily.campColumn} (confirm this is the camp column)`);
    for (const [camp, count] of [...camps].sort((a, b) => b[1] - a[1])) {
      out(`  ${String(count).padStart(5)}  ${JSON.stringify(camp)}`);
    }
    out('');
    out(
      'CHECK: are any two entries the same camp under different casing or ' +
        'abbreviation? If so an alias table is needed before grouping is trustworthy.',
    );
    out('');

    // --- Q5: where the SyncID column can go --------------------------------
    out('## 5. SyncID column placement');
    out(`  widest used column across sampled sheets: ${indexToColumn(widestColumn)}`);
    out(`  configured SyncID column: ${LAYOUT.daily.syncIdColumn}`);
    out(
      columnToIndex(LAYOUT.daily.syncIdColumn) > widestColumn
        ? '  OK: clear of the data on the sampled sheets.'
        : '  PROBLEM: the configured SyncID column is inside the used range. Move it.',
    );
    out('  CHECK this against the widest daily sheet in the whole workbook, not just the sample.');
    out('');

    // --- Queue sheets -------------------------------------------------------
    out('## Queue sheets');
    for (const name of QUEUE_SHEET_NAMES) {
      const exists = worksheets.some((sheet) => sheet.name === name);
      if (!exists) {
        out(`  ${name}: not present`);
        continue;
      }
      const used = await workbook.getUsedRange(name);
      const { startRow } = parseAddress(used.address);
      out(`  ${name}: usedRange ${used.address}`);
      const headerGuess = guessHeaderRow(used.values, startRow);
      out(`    header row guess: ${headerGuess ?? 'none'} (config says ${LAYOUT.queue.headerRow})`);
    }
  } catch (error) {
    console.error(JSON.stringify({ event: 'inspect.failed', ...describeError(error) }));
    process.exitCode = 1;
  } finally {
    await workbook.closeSession();
  }
}

/** Prints which keyword a value matches, never the value itself. */
function summarizeStatus(normalized: string): string {
  const hits: string[] = [];
  for (const keyword of TERMINAL_KEYWORDS) {
    if (normalized.includes(keyword)) hits.push(`${keyword} (terminal)`);
  }
  for (const { keyword } of QUEUE_KEYWORDS) {
    if (normalized.includes(keyword)) hits.push(keyword);
  }
  if (hits.length === 0) {
    // The one place a raw-ish form is useful, since an unmatched value is by
    // definition not a keyword we know. Length-capped and stripped of digits, so
    // a phone number or DOB accidentally typed into column B cannot leak.
    return `UNRECOGNIZED: ${normalized.replace(/\d/g, '#').slice(0, 40)}`;
  }
  return hits.join(' + ');
}

function guessHeaderRow(values: unknown[][], startRow: number): number | null {
  for (let i = 0; i < Math.min(values.length, 20); i++) {
    const row = values[i] ?? [];
    const textCells = row.filter((cell) => {
      const text = String(cell ?? '').trim();
      return text.length > 0 && text.length < 40 && !/^\d+([.,]\d+)?$/.test(text);
    });
    if (textCells.length >= 3) return startRow + i;
  }
  return null;
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'inspect.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
