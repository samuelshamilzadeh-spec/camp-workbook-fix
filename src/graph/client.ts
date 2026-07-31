import type { Logger } from '../logging';
import { describeError } from '../logging';
import type { TokenProvider } from './auth';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export class GraphError extends Error {
  readonly statusCode: number;
  readonly graphCode: string | undefined;
  readonly requestId: string | undefined;

  constructor(statusCode: number, graphCode: string | undefined, requestId: string | undefined) {
    // Deliberately does not include the response body. A failed range write
    // echoes back the payload, and the payload is patient data.
    super(`Graph request failed with ${statusCode}${graphCode ? ` (${graphCode})` : ''}`);
    this.name = 'GraphError';
    this.statusCode = statusCode;
    this.graphCode = graphCode;
    this.requestId = requestId;
  }
}

export interface RequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  body?: unknown;
  headers?: Record<string, string>;
  /** Overrides the default retry budget for this call. */
  maxAttempts?: number;
  /** Overrides the per-attempt timeout for this call. */
  timeoutMs?: number;
  /**
   * False for a call that changes the sheet's shape — a row insert or delete.
   * See REJECTED_BEFORE_EXECUTION.
   */
  idempotent?: boolean;
}

export interface GraphClientOptions {
  maxAttempts?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  /** Per-attempt ceiling. See TIMEOUT_MS. */
  timeoutMs?: number;
  /** Injectable for tests. */
  sleep?: (ms: number) => Promise<void>;
  fetchImpl?: typeof fetch;
}

/**
 * 423 Locked and 409 Conflict are the expected shape of "a staff member has the
 * file open in Excel desktop with autosave on". The brief is explicit that these
 * are not errors, so they are retried and logged at debug, not warn.
 *
 * 429 and 503 carry Retry-After and must be honoured — Graph throttling on the
 * workbook endpoints is per-file and easy to hit at a 5-second cadence.
 */
/**
 * Per-attempt ceiling on a Graph call.
 *
 * Observed against the live workbook: a write issued while two people were
 * editing simply never returned. Not an error, not a lock, not a throttle — the
 * request hung, and was still hanging 31 minutes later.
 *
 * That is worse than a failure. A failure retries; a hang stops the cycle dead
 * and holds the worker until the host kills it, which on a five-second timer
 * means the sync silently stops running. 30s is far above the 95th percentile
 * for these calls (a full 46-sheet scan takes under a second in-session), so a
 * request past it is stuck rather than slow.
 */
const TIMEOUT_MS = 30_000;

const RETRYABLE_STATUS = new Set([409, 423, 429, 500, 501, 502, 503, 504]);

/**
 * The subset of those that mean **the request never ran**.
 *
 * This distinction only matters for an operation that is not idempotent, and in
 * this codebase that is exactly two: `insertRows` and `deleteRows`. Everything
 * else is a PATCH of specific cells or a `/clear`, and running one of those
 * twice lands on the same result.
 *
 * A row delete is different. Deleting row 40 moves row 41 up into it, so a
 * delete that ALREADY SUCCEEDED and is then retried removes a second patient —
 * and the identity check in `applyRemovals` cannot catch it, because it runs
 * above this layer and has already passed by the time the retry fires.
 *
 * 409, 423 and 429 are refusals: the file was locked, or we were throttled, and
 * the service rejected the call before touching the workbook. Retrying those is
 * safe and is the normal operating condition here — staff are in this file all
 * day.
 *
 * 500, 502, 503, 504 and a client-side timeout are all "no answer". The
 * operation may have committed. A write issued while two people were editing
 * was observed hanging for 31 minutes against this workbook, which is precisely
 * the case where a blind retry deletes the wrong row. So for those, the shape-
 * changing calls fail and the next cycle re-plans against what the sheet
 * actually holds — which is what state-based reconciliation is for.
 */
const REJECTED_BEFORE_EXECUTION = new Set([409, 423, 429]);

/**
 * 501 is normally "not implemented" and would never be retried. The Excel
 * workbook API overloads it for `OpenWorkbookBlockedWorkbook`, which means the
 * service could not open the file right now — in practice, a desktop client is
 * holding it. Observed on 18 of 46 sheets in a single scan of the live workbook
 * while staff were editing, so it is a normal operating condition here, not a
 * defect.
 */

export class GraphClient {
  private readonly maxAttempts: number;
  private readonly baseDelayMs: number;
  private readonly maxDelayMs: number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly fetchImpl: typeof fetch;
  private readonly timeoutMs: number;

  constructor(
    private readonly tokens: TokenProvider,
    private readonly log: Logger,
    options: GraphClientOptions = {},
  ) {
    this.maxAttempts = options.maxAttempts ?? 5;
    this.baseDelayMs = options.baseDelayMs ?? 400;
    this.maxDelayMs = options.maxDelayMs ?? 8000;
    this.sleep = options.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.timeoutMs = options.timeoutMs ?? TIMEOUT_MS;
  }

  async request<T>(path: string, options: RequestOptions = {}): Promise<T> {
    const method = options.method ?? 'GET';
    const maxAttempts = options.maxAttempts ?? this.maxAttempts;
    const idempotent = options.idempotent ?? true;
    const url = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;
    const canRetry = (status: number): boolean =>
      idempotent ? RETRYABLE_STATUS.has(status) : REJECTED_BEFORE_EXECUTION.has(status);

    let lastError: GraphError | undefined;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const token = await this.tokens.getToken();
      const startedAt = Date.now();

      let response: Response;
      try {
        response = await this.fetchImpl(url, {
          method,
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json',
            ...(options.body !== undefined ? { 'Content-Type': 'application/json' } : {}),
            ...options.headers,
          },
          ...(options.body !== undefined ? { body: JSON.stringify(options.body) } : {}),
          // Without this a hung request never returns and the cycle stops
          // forever. Observed on the live workbook.
          signal: AbortSignal.timeout(options.timeoutMs ?? this.timeoutMs),
        });
      } catch (error) {
        // Network failure or timeout. Both retryable on the same budget: a hung
        // request is usually the workbook being momentarily unreachable, and
        // the next attempt commonly succeeds.
        //
        // Except for a row insert or delete, where "no answer" is not the same
        // as "did not happen" — see REJECTED_BEFORE_EXECUTION.
        const timedOut = (error as Error)?.name === 'TimeoutError';
        if (attempt >= maxAttempts || !idempotent) {
          this.log.error('graph.request_failed', {
            action: method,
            attempt,
            reason: timedOut ? 'timeout' : 'network',
            durationMs: Date.now() - startedAt,
          });
          throw error;
        }
        this.log.debug(timedOut ? 'graph.timeout_retry' : 'graph.network_retry', {
          attempt,
          durationMs: Date.now() - startedAt,
          ...describeError(error),
        });
        await this.sleep(this.backoffMs(attempt));
        continue;
      }

      const durationMs = Date.now() - startedAt;

      if (response.ok) {
        if (response.status === 204) return undefined as T;
        const text = await response.text();
        return (text ? JSON.parse(text) : undefined) as T;
      }

      const graphCode = await readErrorCode(response);
      const requestId = response.headers.get('request-id') ?? undefined;
      lastError = new GraphError(response.status, graphCode, requestId);

      if (!canRetry(response.status) || attempt >= maxAttempts) {
        this.log.error('graph.request_failed', {
          action: method,
          statusCode: response.status,
          status: graphCode,
          requestId,
          attempt,
          durationMs,
        });
        throw lastError;
      }

      const delay = this.retryDelayMs(response, attempt);
      this.log.debug('graph.retry', {
        action: method,
        statusCode: response.status,
        status: graphCode,
        attempt,
        durationMs,
        reason: response.status === 423 || response.status === 409 ? 'locked' : 'transient',
      });
      await this.sleep(delay);
    }

    throw lastError ?? new Error('Graph request exhausted retries without a response.');
  }

  private retryDelayMs(response: Response, attempt: number): number {
    const retryAfter = response.headers.get('retry-after');
    if (retryAfter) {
      const seconds = Number(retryAfter);
      // `> 0`, not `>= 0`. Graph answers a 503 with `Retry-After: 0`, and
      // honouring that literally spent the whole retry budget in under a tenth
      // of a second — five attempts at 11-30 ms each, which is not a retry
      // policy, it is five ways to fail at once. Observed against the live
      // workbook mid-session.
      //
      // A zero means "no guidance", so it falls through to the jittered
      // exponential backoff below. That is also what stops a fleet of instances
      // arriving back together the moment a service wobbles.
      if (Number.isFinite(seconds) && seconds > 0) {
        return Math.min(seconds * 1000, this.maxDelayMs);
      }
    }
    return this.backoffMs(attempt);
  }

  /** Exponential with full jitter, so concurrent retries do not synchronize. */
  private backoffMs(attempt: number): number {
    const ceiling = Math.min(this.baseDelayMs * 2 ** (attempt - 1), this.maxDelayMs);
    return Math.round(ceiling * (0.5 + Math.random() * 0.5));
  }
}

async function readErrorCode(response: Response): Promise<string | undefined> {
  try {
    const body = (await response.json()) as { error?: { code?: string } };
    return body?.error?.code;
  } catch {
    return undefined;
  }
}
