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
  | 'Ineligible & Inactive'
  | 'United Refuah';

export const QUEUE_SHEET_NAMES: readonly QueueSheetName[] = [
  'Verify Insurance',
  'Missing Info',
  'Not Accepted',
  'Ineligible & Inactive',
  'United Refuah',
] as const;

/**
 * Destinations whose rows are append-only: copied across once and then never
 * touched again.
 *
 * The office confirmed United Refuah works this way — a row lands there and no
 * change is ever sent back to the daily sheet. So write-back and clear-and-
 * remove are both suppressed for it, and it is the reason those behaviours are
 * per-destination rather than global.
 */
export const APPEND_ONLY_DESTINATIONS: readonly QueueSheetName[] = ['United Refuah'] as const;

/**
 * Where each queue lives in the live workbook, verified 2026-07-30.
 *
 * `Not Accepted ` really does carry a trailing space. Matching is whitespace-
 * and case-forgiving (see resolveSheetName), so these need only be close.
 *
 * `Missing Info (New)` was RENAMED to `Missing Info` some time before
 * 2026-07-31. Nothing failed loudly: `resolveSheetName` reported the tab as
 * absent, the cycle logged `queue.sheet_missing`, and 194 patients were computed
 * into a queue that no longer had anywhere to put them. A tab rename is a
 * silent outage in this design, which is worth remembering when one is planned.
 */
export const QUEUE_SHEET_TABS: Record<QueueSheetName, string> = {
  'Verify Insurance': 'Verify Insurance',
  'Missing Info': 'Missing Info',
  'Not Accepted': 'Not Accepted',
  'Ineligible & Inactive': 'Ineligible & Inactive',
  // The original tab was renamed `United Refuah (old)` and hidden; a fresh one
  // was created under the canonical name on 2026-07-30 and is empty. It is the
  // first target for Phase 2b: 85 rows, append-only, no inserts needed.
  'United Refuah': 'United Refuah',
};

/**
 * Superseded or historical tabs. The office confirmed these are not live and
 * must never be read or written.
 */
export const IGNORED_TABS: readonly string[] = [
  // Hidden by the office 2026-07-30 and explicitly out of scope: as far as this
  // code is concerned these tabs do not exist. Never read, never written.
  'United Refuah (old)',
  'Dont Take Ins (old)',
  'Missing Ins info',
  // Historical.
  'Missing info 25',
  '2025 Archive',
  '2024 Archive',
] as const;

/**
 * Queue sheet columns, VERIFIED 2026-07-30 against the live tabs.
 *
 * `Date of Visit` and `Source Row`, then every patient field from the daily
 * sheet, D through R, in the same order. Nothing is dropped: the office
 * confirmed the whole row transfers, insurance ids and medical detail included.
 *
 * NOTE ON `Notes`: the brief places it immediately after `Source Row`. No live
 * queue tab has it, so it is not here yet. Adding it in Phase 5 shifts every
 * column after B one to the right, which is a migration across all five tabs
 * rather than an edit to this list — do them together.
 */
export const QUEUE_COLUMNS = [
  'Date of Visit',
  'Source Row',
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
  'Insurance ID #',
  'Medicaid #',
  'Medical History',
  'Medications',
  'Allergies',
] as const;

export type QueueColumn = (typeof QUEUE_COLUMNS)[number];

/**
 * Columns that live only on the queue sheet and must NOT propagate back to the
 * daily sheet. `Notes` joins this list when it is added in Phase 5.
 */
export const QUEUE_ONLY_COLUMNS: readonly QueueColumn[] = ['Date of Visit', 'Source Row'] as const;

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
  // Append-only record rather than a work queue, so nothing is chased and
  // nothing is shaded red.
  'United Refuah': [],
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
    // VERIFIED 2026-07-30 against the live header row:
    // A LABS | B EMR | C CAMP NAME | D LAST Nm | E FIRST Nm | F DOB | G GENDER |
    // H BILLING ADDRESS | I CITY | J St | K ZIP | L PHONE NUMBER |
    // M INS CARRIER | N INS ID # | O Medicaid # | P MEDICAL HISTORY | Q MEDS |
    // R ALLERGIES
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
      'Insurance ID #': 'N',
      'Medicaid #': 'O',
      'Medical History': 'P',
      Medications: 'Q',
      Allergies: 'R',
    },
  },
  queue: {
    // VERIFIED 2026-07-31 against all five live tabs: the header is on ROW 1 and
    // data starts on row 2. Every one of them.
    //
    // The brief described the mirror sheet as instructions in rows 1-7, headers
    // on row 9, data from row 10, and this said 9/10 on that authority. It is
    // wrong, and it was wrong quietly: rows 2-9 of `Not Accepted ` and
    // `Ineligible & Inactive` were never read, so 12 real patient rows were
    // invisible to Phase 1 and were skipped by Phase 2a's adoption.
    //
    // `detectQueueShape` now reads the header row off the sheet and these two
    // numbers are only the fallback for when it finds nothing. They are set to
    // the truth anyway, because a fallback that is also wrong helps nobody.
    headerRow: 1,
    firstDataRow: 2,
    syncIdColumn: 'BA',
    firstColumn: 'A',
  },
  notesSheetName: '_SyncNotes',
  controlSheetName: '_SyncControl',
  // VERIFIED 2026-07-30 against the live workbook: every tab that is not a
  // daily visit sheet. Several are historical or superseded, kept here so the
  // scan never mistakes one for a daily sheet.
  knownNonDailySheets: [
    'United Refuah',
    'United Refuah (old)',
    'Verify Insurance',
    'Claude Log',
    'Cheat Sheet',
    '_Feed',
    '2025 Archive',
    '2024 Archive',
    'Missing info 25',
    'Dont Take Ins (old)',
    'Missing Ins info',
    'Missing Info',
    'Missing Info (New)',
    'Not Accepted ',
    'Ineligible & Inactive',
  ],
  // VERIFIED 2026-07-30: sheets are named `July 30, 2026` — long month name,
  // no leading zero on the day, comma before the year. The ISO and US numeric
  // forms are still accepted so a sheet added in another style is not silently
  // skipped.
  dailySheetPattern:
    /^\s*(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\s*$/i,
  parseDailySheetDate(sheetName: string): Date | null {
    const raw = sheetName.trim();

    // `July 30, 2026`, `Jul 30 2026`, `Sept. 3, 2026`
    const named = /^([a-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/i.exec(raw);
    if (named) {
      const month = MONTHS.indexOf(named[1]!.slice(0, 3).toLowerCase()) + 1;
      if (month > 0) return utcDate(Number(named[3]), month, Number(named[2]));
      return null;
    }

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

const MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];

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
    const shadowed: string[] = [];

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
      if (!key) continue;

      const existing = process.env[key];
      if (existing === undefined) {
        process.env[key] = value;
      } else if (existing !== value && value !== '') {
        // An exported shell variable silently beating the file is a genuinely
        // confusing failure: the file looks right, gets read, and is ignored.
        // A stale exported AZURE_CLIENT_SECRET presents as "invalid client
        // secret" against a .env that is perfectly correct.
        shadowed.push(key);
      }
    }

    if (shadowed.length > 0) {
      process.stderr.write(
        `WARNING: ${shadowed.join(', ')} ${shadowed.length === 1 ? 'is' : 'are'} set in the ` +
          `environment and override${shadowed.length === 1 ? 's' : ''} .env. ` +
          `Run \`unset ${shadowed.join(' ')}\` if you meant to use the file.\n`,
      );
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

export interface LoadConfigOptions {
  /**
   * `scripts/resolve-workbook.ts` exists precisely to discover GRAPH_DRIVE_ID
   * and GRAPH_ITEM_ID, so it cannot require them to already be set.
   */
  requireWorkbook?: boolean;
}

export function loadConfig(options: LoadConfigOptions = {}): RuntimeConfig {
  loadEnvFileOnce();

  const requireWorkbook = options.requireWorkbook ?? true;

  const useManagedIdentity = boolEnv('AZURE_USE_MANAGED_IDENTITY', false);
  if (requireWorkbook && !env('GRAPH_DRIVE_ID') && !env('GRAPH_SITE_ID')) {
    throw new Error(
      'Set GRAPH_DRIVE_ID (preferred) or GRAPH_SITE_ID. Get it by running:\n' +
        '  npm run resolve -- "<the workbook URL from your browser>"',
    );
  }

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
    itemId: requireWorkbook ? requireEnv('GRAPH_ITEM_ID') : (env('GRAPH_ITEM_ID') ?? ''),
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
