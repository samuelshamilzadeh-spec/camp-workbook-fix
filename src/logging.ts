/**
 * Structured logging with a hard rule: no patient data ever leaves this process.
 *
 * The only patient-linked identifiers permitted in a log line are a SyncID, a
 * sheet name, and a row number. Everything else about a patient — name, DOB,
 * phone, address, insurance, and the raw contents of any cell — is redacted.
 *
 * `redact()` is applied to every field of every log record rather than trusting
 * call sites to remember, because the failure mode is silent and permanent:
 * once a DOB is in Application Insights it is in a system with different
 * retention and access rules than the workbook.
 */

export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface LogSink {
  debug(message: string, ...args: unknown[]): void;
  info(message: string, ...args: unknown[]): void;
  warn(message: string, ...args: unknown[]): void;
  error(message: string, ...args: unknown[]): void;
}

/** Field names that may be logged as-is. Everything else is redacted. */
const SAFE_KEYS = new Set([
  'syncId',
  'sheet',
  'sourceSheet',
  'sheetName',
  'sheets',
  'row',
  'sourceRow',
  'rowNumber',
  'destination',
  'queueSheet',
  'phase',
  'dryRun',
  'action',
  'count',
  'counts',
  'reason',
  'status',
  'statusCode',
  'durationMs',
  'attempt',
  'field',
  'fields',
  'column',
  'columns',
  'keyword',
  'keywords',
  'matched',
  'lastModified',
  'lastModifiedBy',
  'etag',
  'cycleId',
  'wouldWrite',
  'skipped',
  'address',
  'requestId',
  'camp',
  'campKey',
]);

/**
 * `camp` is on the safe list because grouping is useless in logs without it and
 * a camp name is an organisation, not a patient. It is still a business
 * identifier, so it is never logged alongside anything else identifying.
 */

const MAX_STRING = 200;

export function redact(value: unknown, key?: string): unknown {
  if (value === null || value === undefined) return value;

  if (key !== undefined && !SAFE_KEYS.has(key)) {
    return `[redacted:${typeofLabel(value)}]`;
  }

  if (typeof value === 'string') {
    return value.length > MAX_STRING ? `${value.slice(0, MAX_STRING)}…` : value;
  }
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  if (value instanceof Date) return value.toISOString();

  if (Array.isArray(value)) {
    // Arrays reached via a safe key (e.g. `matched`, `fields`) hold column or
    // keyword names, not cell values.
    return value.slice(0, 25).map((item) => redact(item, key));
  }

  if (typeof value === 'object') {
    const out: Record<string, unknown> = {};

    // `counts` is a map of category -> how many, and its keys are labels this
    // codebase writes rather than anything read out of a cell. Checking those
    // keys against the allow-list redacted every number in it — so the one line
    // that says what a cycle actually did read
    // `{"stamped":"[redacted:number]","appended":"[redacted:number]"}`, which is
    // the whole of this job's observability reduced to noise.
    //
    // Numbers pass through with their label. Anything else in there does not,
    // so a count map that somehow carried a value still cannot leak one.
    if (key === 'counts') {
      for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
        out[k] = typeof v === 'number' ? v : `[redacted:${typeofLabel(v)}]`;
      }
      return out;
    }

    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      out[k] = redact(v, k);
    }
    return out;
  }

  return `[redacted:${typeofLabel(value)}]`;
}

function typeofLabel(value: unknown): string {
  if (Array.isArray(value)) return 'array';
  return typeof value;
}

/** Redacts a whole log payload object (top-level keys are checked too). */
export function redactPayload(payload: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(payload)) {
    out[k] = redact(v, k);
  }
  return out;
}

export interface Logger {
  event(level: LogLevel, name: string, payload?: Record<string, unknown>): void;
  debug(name: string, payload?: Record<string, unknown>): void;
  info(name: string, payload?: Record<string, unknown>): void;
  warn(name: string, payload?: Record<string, unknown>): void;
  error(name: string, payload?: Record<string, unknown>): void;
  child(context: Record<string, unknown>): Logger;
}

export function createLogger(sink: LogSink, context: Record<string, unknown> = {}): Logger {
  const base = redactPayload(context);

  const logger: Logger = {
    event(level, name, payload = {}) {
      const record = { event: name, ...base, ...redactPayload(payload) };
      sink[level](JSON.stringify(record));
    },
    debug: (name, payload) => logger.event('debug', name, payload),
    info: (name, payload) => logger.event('info', name, payload),
    warn: (name, payload) => logger.event('warn', name, payload),
    error: (name, payload) => logger.event('error', name, payload),
    child: (extra) => createLogger(sink, { ...context, ...extra }),
  };

  return logger;
}

export const consoleSink: LogSink = {
  debug: (m) => console.debug(m),
  info: (m) => console.info(m),
  warn: (m) => console.warn(m),
  error: (m) => console.error(m),
};

/**
 * Errors from Graph can carry request bodies, and a request body here is a range
 * write full of patient data. Only ever log the shape of an error.
 */
export function describeError(error: unknown): Record<string, unknown> {
  if (error instanceof Error) {
    const withStatus = error as Error & { statusCode?: number; graphCode?: string };
    return {
      name: error.name,
      // Message text is developer-authored in this codebase, but a Graph error
      // message can echo a cell value, so it is length-capped and never
      // interpolated into a broader string.
      reason: String(error.message).slice(0, MAX_STRING),
      statusCode: withStatus.statusCode,
      status: withStatus.graphCode,
    };
  }
  return { name: 'UnknownError', reason: `[redacted:${typeofLabel(error)}]` };
}
