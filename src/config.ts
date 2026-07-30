import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

/**
 * Every assumption about the workbook's shape lives here, in one place, so that
 * checking the build brief against the real file is a matter of editing this
 * file rather than hunting through the code.
 *
 * The brief is explicit that these are *unverified*. Anything marked UNVERIFIED
 * below is a guess taken from the brief and must be checked against the actual
 * workbook (`npm run inspect`) before any write phase is enabled. The
 * `SYNC_LAYOUT_VERIFIED` gate in `assertLayoutVerified()` exists to make that
 * check impossible to forget.
 */

export type QueueSheetName =
  | 'Verify Insurance'
  | 'Missing Info'
  | 'Not Accepted'
  | 'Ineligible & Inactive';

export const QUEUE_SHEET_NAMES: readonly QueueSheetName[] = [
  'Verify Insurance',
  'Missing Info',
  'Not Accepted',
  'Ineligible & Inactive',
] as const;

/**
 * Queue sheet column order, from the brief. Notes sits immediately after Source
 * Row, ahead of the patient fields.
 */
export const QUEUE_COLUMNS = [
  'Date of Visit',
  'Source Row',
  'Notes',
  'Last Name',
  'First Name',
  'Date of Birth',
  'Gender',
  'Billing Address',
  'City',
  'State',
  'Zip Code',
  'Phone Number',
  'Insurance Carrier',
] as const;

export type QueueColumn = (typeof QUEUE_COLUMNS)[number];

/** Columns staff may edit on a queue sheet that must NOT propagate back to the daily sheet. */
export const QUEUE_ONLY_COLUMNS: readonly QueueColumn[] = ['Notes', 'Source Row'] as const;

export interface DailySheetLayout {
  /** 1-based row holding the column headers. UNVERIFIED. */
  headerRow: number;
  /** 1-based first row of patient data. UNVERIFIED. */
  firstDataRow: number;
  /** Column letter holding Status. The brief states column B and is confident. */
  statusColumn: string;
  /** Column letter holding the camp name. The brief says "believed to be column C". UNVERIFIED. */
  campColumn: string;
  /** Column letter for the SyncID stamp. Must be identical on every sheet. UNVERIFIED. */
  syncIdColumn: string;
  /**
   * Maps a queue sheet column to the column letter holding the same field on a
   * daily sheet. Drives both populate (read) and write-back (write). Every
   * entry is UNVERIFIED.
   */
  fieldColumns: Partial<Record<QueueColumn, string>>;
}

export interface QueueSheetLayout {
  headerRow: number;
  firstDataRow: number;
  syncIdColumn: string;
  /** Column letter of the first queue column ('Date of Visit'); the rest follow in order. */
  firstColumn: string;
}

export interface WorkbookLayout {
  daily: DailySheetLayout;
  queue: QueueSheetLayout;
  /** Hidden sheet holding SyncID -> note text. Phase 5. */
  notesSheetName: string;
  /** Hidden sheet holding the loop-guard marker written after each of our own write batches. */
  controlSheetName: string;
  /**
   * Sheets that are never daily visit sheets. Anything not matching
   * `dailySheetPattern` is skipped anyway; this list is for logging clarity.
   */
  knownNonDailySheets: readonly string[];
  /**
   * A worksheet whose name matches this is treated as a daily visit sheet.
   * The brief says "the sheet name is the date" but not in which format.
   * UNVERIFIED — `npm run inspect` prints the real sheet names.
   */
  dailySheetPattern: RegExp;
  /** Parses a daily sheet name into a date, for "most recent N sheets". UNVERIFIED. */
  parseDailySheetDate: (sheetName: string) => Date | null;
}

/**
 * Required fields get dark red shading when blank.
 *
 * The brief asks whether "required" differs per status. It plausibly does — a
 * Missing Info row is by definition missing something — so this is modelled per
 * status from the start, currently with an identical set. UNVERIFIED.
 */
export const REQUIRED_FIELDS: Record<QueueSheetName, readonly QueueColumn[]> = {
  'Verify Insurance': [
    'Last Name',
    'First Name',
    'Date of Birth',
    'Insurance Carrier',
  ],
  'Missing Info': ['Last Name', 'First Name', 'Date of Birth', 'Phone Number'],
  'Not Accepted': ['Last Name', 'First Name', 'Date of Birth', 'Insurance Carrier'],
  'Ineligible & Inactive': [
    'Last Name',
    'First Name',
    'Date of Birth',
    'Phone Number',
    'Insurance Carrier',
  ],
};

/** Dark red, matching the existing mirror sheets' shading. UNVERIFIED. */
export const BLANK_REQUIRED_FILL = '#C00000';

export const LAYOUT: WorkbookLayout = {
  daily: {
    headerRow: 1,
    firstDataRow: 2,
    statusColumn: 'B',
    campColumn: 'C',
    syncIdColumn: 'BA',
    fieldColumns: {
      'Last Name': 'D',
      'First Name': 'E',
      'Date of Birth': 'F',
      Gender: 'G',
      'Billing Address': 'H',
      City: 'I',
      State: 'J',
      'Zip Code': 'K',
      'Phone Number': 'L',
      'Insurance Carrier': 'M',
    },
  },
  queue: {
    // The brief describes the existing mirror sheet: instructions in rows 1-7,
    // headers on row 9, data from row 10. The new queue sheets are built by us,
    // so we can hold that shape deliberately rather than inheriting it.
    headerRow: 9,
    firstDataRow: 10,
    syncIdColumn: 'BA',
    firstColumn: 'A',
  },
  notesSheetName: '_SyncNotes',
  controlSheetName: '_SyncControl',
  knownNonDailySheets: ['United Refresh', 'Claude Log'],
  // Matches 2026-07-30, 7-30-2026, 7/30/26 and similar. Deliberately loose;
  // narrow it once the real convention is known.
  dailySheetPattern: /^\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\s*$/,
  parseDailySheetDate(sheetName: string): Date | null {
    const raw = sheetName.trim();
    const iso = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(raw);
    if (iso) {
      return utcDate(Number(iso[1]), Number(iso[2]), Number(iso[3]));
    }
    const us = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(raw);
    if (us) {
      const year = Number(us[3]);
      return utcDate(year < 100 ? 2000 + year : year, Number(us[1]), Number(us[2]));
    }
    return null;
  },
};

function utcDate(year: number, month: number, day: number): Date | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const d = new Date(Date.UTC(year, month - 1, day));
  return Number.isNaN(d.getTime()) ? null : d;
}

// --- Runtime settings -------------------------------------------------------

export type Phase = 0 | 1 | 2 | 3 | 4 | 5;

export interface RuntimeConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string | undefined;
  useManagedIdentity: boolean;
  managedIdentityClientId: string | undefined;
  /** Preferred: addresses the file wherever it lives. */
  driveId: string | undefined;
  /** Fallback; resolves to the site's default document library only. */
  siteId: string | undefined;
  itemId: string;
  phase: Phase;
  dryRun: boolean;
  layoutVerified: boolean;
  /** Days back from today that are scanned every cycle. */
  hotDaysBack: number;
  /**
   * Days forward from today that are scanned every cycle. The office creates
   * daily sheets in advance, so this has to cover how far ahead they work.
   */
  hotDaysForward: number;
  /** Sheets pulled from the rotation each cycle, covering the rest of the year. */
  coldBatchSize: number;
  /**
   * At or below this many daily sheets, every sheet is read every cycle. Above
   * it, the hot/cold tiering kicks in. Sized so one camp season scans fully and
   * a workbook that accumulates seasons degrades instead of overrunning.
   */
  maxSheetsPerCycle: number;
  /** Concurrent Graph reads. The whole point of a full scan fitting in a cycle. */
  readConcurrency: number;
  stateConnectionString: string | undefined;
  stateContainer: string;
}

/**
 * Loads `.env` from the working directory when present, so the CLI scripts pick
 * up local settings without every invocation needing `node --env-file=`.
 *
 * Never throws and never overwrites a variable that is already set: in Azure the
 * app settings are the real configuration and there is no .env file at all.
 */
let envFileLoaded = false;
function loadEnvFileOnce(): void {
  if (envFileLoaded) return;
  envFileLoaded = true;

  const path = join(process.cwd(), '.env');
  if (!existsSync(path)) return;

  try {
    for (const line of readFileSync(path, 'utf8').split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) continue;
      const eq = trimmed.indexOf('=');
      if (eq === -1) continue;

      const key = trimmed.slice(0, eq).trim();
      let value = trimmed.slice(eq + 1).trim();
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }
      if (key && process.env[key] === undefined) process.env[key] = value;
    }
  } catch {
    // A malformed .env is a local-development problem, not a reason to refuse
    // to start. requireEnv reports whatever is actually missing.
  }
}

function env(name: string): string | undefined {
  const value = process.env[name];
  return value === undefined || value.trim() === '' ? undefined : value.trim();
}

function requireEnv(name: string): string {
  const value = env(name);
  if (!value) {
    throw new Error(`Missing required setting ${name}. See .env.example.`);
  }
  return value;
}

function boolEnv(name: string, fallback: boolean): boolean {
  const value = env(name);
  if (value === undefined) return fallback;
  return /^(1|true|yes|on)$/i.test(value);
}

export function loadConfig(): RuntimeConfig {
  loadEnvFileOnce();

  const useManagedIdentity = boolEnv('AZURE_USE_MANAGED_IDENTITY', false);
  const phaseRaw = Number(env('SYNC_PHASE') ?? '1');
  if (!Number.isInteger(phaseRaw) || phaseRaw < 0 || phaseRaw > 5) {
    throw new Error(`SYNC_PHASE must be an integer 0-5, got "${env('SYNC_PHASE')}"`);
  }

  return {
    tenantId: requireEnv('AZURE_TENANT_ID'),
    clientId: requireEnv('AZURE_CLIENT_ID'),
    clientSecret: env('AZURE_CLIENT_SECRET'),
    useManagedIdentity,
    managedIdentityClientId: env('AZURE_MANAGED_IDENTITY_CLIENT_ID'),
    driveId: env('GRAPH_DRIVE_ID'),
    siteId: env('GRAPH_SITE_ID'),
    itemId: requireEnv('GRAPH_ITEM_ID'),
    phase: phaseRaw as Phase,
    // Dry run defaults to ON. Turning writes on has to be a deliberate act.
    dryRun: boolEnv('SYNC_DRY_RUN', true),
    layoutVerified: boolEnv('SYNC_LAYOUT_VERIFIED', false),
    hotDaysBack: Number(env('SYNC_HOT_DAYS_BACK') ?? '7'),
    hotDaysForward: Number(env('SYNC_HOT_DAYS_FORWARD') ?? '14'),
    coldBatchSize: Number(env('SYNC_COLD_BATCH_SIZE') ?? '10'),
    maxSheetsPerCycle: Number(env('SYNC_MAX_SHEETS_PER_CYCLE') ?? '90'),
    readConcurrency: Number(env('SYNC_READ_CONCURRENCY') ?? '8'),
    stateConnectionString:
      env('AZURE_STORAGE_CONNECTION_STRING') ?? env('AzureWebJobsStorage'),
    stateContainer: env('SYNC_STATE_CONTAINER') ?? 'camp-sync-state',
  };
}

/**
 * Writes are refused until someone has checked LAYOUT against the real file and
 * flipped the flag. Called by every code path that can modify the workbook.
 */
export function assertLayoutVerified(config: RuntimeConfig): void {
  if (!config.layoutVerified) {
    throw new Error(
      'Refusing to write: SYNC_LAYOUT_VERIFIED is false. Check the LAYOUT block ' +
        'in src/config.ts against the real workbook (`npm run inspect`) and set ' +
        'SYNC_LAYOUT_VERIFIED=true once the header rows, camp column, SyncID ' +
        'column and field mapping are confirmed.',
    );
  }
}
