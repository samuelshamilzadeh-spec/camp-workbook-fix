/**
 * Creates a queue tab that does not exist yet, matching the layout of the ones
 * already in use.
 *
 *   npm run create-sheet -- "Verify Insurance"           # dry run, prints the plan
 *   npm run create-sheet -- "Verify Insurance" --apply
 *
 * Deliberately narrow: it adds a worksheet and writes one header row. It never
 * touches an existing tab, and it refuses outright if the tab already exists,
 * so re-running it cannot clobber live data.
 *
 * The header row is copied from a tab already in use rather than from the build
 * brief, because staff have to work in this sheet and it should look like the
 * ones beside it. The brief's `Notes` column is deliberately NOT added here —
 * none of the existing queue tabs have one, and introducing it on a single tab
 * would leave the five inconsistent. Notes and SyncID get added to all of them
 * together in Phase 2.
 */

import { loadConfig } from '../src/config';
import { createCredential, createTokenProvider } from '../src/graph/auth';
import { GraphClient } from '../src/graph/client';
import { consoleSink, createLogger, describeError } from '../src/logging';
import { assertNoFormulas, worksheetPath } from '../src/graph/workbook';
import { resolveSheetName } from '../src/domain/queueSheets';

/** The tab whose header row is copied. Must be one staff already use. */
const TEMPLATE_TAB = 'Missing Info (New)';

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const name = args.find((a) => !a.startsWith('--'));
  const apply = args.includes('--apply');

  if (!name) throw new Error('Usage: npm run create-sheet -- "Verify Insurance" [--apply]');

  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 0, dryRun: !apply });
  const graph = new GraphClient(createTokenProvider(createCredential(config)), log);
  const wb = `/drives/${config.driveId}/items/${config.itemId}/workbook`;

  const session = await graph.request<{ id: string }>(`${wb}/createSession`, {
    method: 'POST',
    body: { persistChanges: apply },
  });
  const headers = { 'workbook-session-id': session.id };

  try {
    const sheets = await graph.request<{ value: { name: string }[] }>(
      `${wb}/worksheets?$select=name`,
      { headers },
    );
    const names = sheets.value.map((s) => s.name);

    const existing = resolveSheetName(name, names);
    if (existing) {
      // Creating over a live tab would be unrecoverable, so this is fatal
      // rather than a skip.
      throw new Error(
        `A tab matching "${name}" already exists as "${existing}". Refusing to touch it.`,
      );
    }

    const template = resolveSheetName(TEMPLATE_TAB, names);
    if (!template) throw new Error(`Template tab "${TEMPLATE_TAB}" not found.`);

    const headerRow = await graph.request<{ values: unknown[][] }>(
      `${worksheetPath(wb, template)}/range(address='A1:Z1')?$select=values`,
      { headers },
    );
    const row = (headerRow.values[0] ?? []).map((cell) => String(cell ?? '').trim());
    const width = row.reduce((last, cell, i) => (cell ? i + 1 : last), 0);
    const values = [row.slice(0, width)];

    process.stdout.write(`\nWould create tab: ${JSON.stringify(name)}\n`);
    process.stdout.write(`Header copied from: ${JSON.stringify(template)}\n`);
    process.stdout.write(`Columns (${width}): ${values[0]!.join(' | ')}\n\n`);

    if (!apply) {
      process.stdout.write('Dry run. Re-run with --apply to create it.\n');
      return;
    }

    assertNoFormulas(values, name, 'header');

    await graph.request(`${wb}/worksheets/add`, {
      method: 'POST',
      body: { name },
      headers,
    });

    const lastColumn = String.fromCharCode(64 + Math.min(width, 26));
    await graph.request(
      `${worksheetPath(wb, name)}/range(address='A1:${lastColumn}1')`,
      { method: 'PATCH', body: { values }, headers },
    );

    log.info('sheet.created', { sheet: name, columns: values[0]!.length });
    process.stdout.write(`\nCreated "${name}" with ${width} header columns.\n`);
  } finally {
    await graph
      .request(`${wb}/closeSession`, { method: 'POST', headers, maxAttempts: 2 })
      .catch(() => {});
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error(JSON.stringify({ event: 'create_sheet.fatal', ...describeError(error) }));
    process.exit(1);
  });
}
