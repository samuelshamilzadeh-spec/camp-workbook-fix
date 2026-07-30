/**
 * Turns a workbook URL copied from the browser into the ids `.env` needs.
 *
 *   npm run resolve -- "https://premierassist.sharepoint.com/sites/.../Workbook.xlsx"
 *
 * Beats hunting for a site name by hand, and works wherever the file actually
 * lives — a site's default document library, a second library on the same site,
 * a Teams channel's files, or someone's OneDrive. All of those resolve to a
 * driveId plus an itemId, which is what addresses the file unambiguously.
 *
 * Read-only. Uses the `/shares` endpoint, which takes the URL as an opaque
 * token, so the URL does not need to be a sharing link — the address bar URL
 * works.
 *
 * Copy a URL that points at the file itself. Any of these are fine:
 *   .../Shared Documents/Camp/Workbook.xlsx
 *   .../Doc.aspx?sourcedoc=%7B...%7D&file=Workbook.xlsx&action=default
 *   a "Copy link" sharing URL
 */

import { loadConfig } from '../src/config';
import { createCredential, createTokenProvider } from '../src/graph/auth';
import { GraphClient } from '../src/graph/client';
import { consoleSink, createLogger, describeError } from '../src/logging';

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
  file?: { mimeType?: string };
  parentReference?: {
    driveId?: string;
    siteId?: string;
    path?: string;
  };
}

/** `/shares` takes a base64url of the URL, prefixed with `u!`. */
function encodeSharingUrl(url: string): string {
  const base64 = Buffer.from(url, 'utf8')
    .toString('base64')
    .replace(/=+$/, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
  return `u!${base64}`;
}

function out(line = ''): void {
  process.stdout.write(`${line}\n`);
}

async function main(): Promise<void> {
  const url = process.argv.slice(2).find((arg) => arg.startsWith('http'));
  if (!url) {
    throw new Error(
      'Pass the workbook URL, quoted:\n' +
        '  npm run resolve -- "https://premierassist.sharepoint.com/sites/.../Workbook.xlsx"',
    );
  }

  const config = loadConfig();
  const log = createLogger(consoleSink, { phase: 1, dryRun: true });
  const graph = new GraphClient(createTokenProvider(createCredential(config)), log);

  const item = await graph.request<DriveItem>(
    `/shares/${encodeSharingUrl(url)}/driveItem?$select=id,name,size,lastModifiedDateTime,webUrl,file,parentReference`,
  );

  const driveId = item.parentReference?.driveId;
  const siteId = item.parentReference?.siteId;

  out();
  out('Resolved.');
  out(`  file:         ${item.name}`);
  if (item.size !== undefined) out(`  size:         ${(item.size / 1024 / 1024).toFixed(1)} MB`);
  if (item.lastModifiedDateTime) out(`  lastModified: ${item.lastModifiedDateTime}`);
  if (item.parentReference?.path) out(`  folder:       ${item.parentReference.path}`);
  out();

  if (!/\.xlsx?$|\.xlsm$/i.test(item.name)) {
    out('WARNING: that does not look like an Excel workbook. Check the URL.');
    out();
  }

  out('Put these in .env:');
  out();
  out(`GRAPH_DRIVE_ID=${driveId ?? '<missing — see below>'}`);
  out(`GRAPH_ITEM_ID=${item.id}`);
  if (siteId) out(`GRAPH_SITE_ID=${siteId}`);
  out();

  if (!driveId) {
    out(
      'No driveId came back, which is unusual. Fall back to GRAPH_SITE_ID, but ' +
        'note that only finds the file if it sits in the site default library.',
    );
  } else {
    out('GRAPH_DRIVE_ID is what the code uses; GRAPH_SITE_ID is informational.');
  }

  if (item.size !== undefined && item.size > 40 * 1024 * 1024) {
    out();
    out(
      'NOTE: this workbook is large. The Graph Workbook API refuses files over ' +
        'roughly 75 MB, and gets slow well before that. Worth knowing now rather ' +
        'than at Phase 2.',
    );
  }
}

if (require.main === module) {
  main().catch((error) => {
    const described = describeError(error);
    console.error(JSON.stringify({ event: 'resolve.fatal', ...described }));
    if ((error as { statusCode?: number }).statusCode === 403) {
      console.error(
        'A 403 here usually means admin consent has not been granted yet, or has ' +
          'not finished propagating. Check appRoleAssignments on the service principal.',
      );
    }
    if ((error as { statusCode?: number }).statusCode === 404) {
      console.error(
        'A 404 means the URL did not resolve to a file. Copy it from the browser ' +
          'address bar with the workbook open, and quote it in the shell.',
      );
    }
    process.exit(1);
  });
}
