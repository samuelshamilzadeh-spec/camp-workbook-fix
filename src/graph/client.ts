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
}

export interface GraphClientOptions {
  maxAttempts?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
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
const RETRYABLE_STATUS = new Set([409, 423, 429, 500, 502, 503, 504]);

export class GraphClient {
  private readonly maxAttempts: number;
  private readonly baseDelayMs: number;
  private readonly maxDelayMs: number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly fetchImpl: typeof fetch;

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
  }

  async request<T>(path: string, options: RequestOptions = {}): Promise<T> {
    const method = options.method ?? 'GET';
    const maxAttempts = options.maxAttempts ?? this.maxAttempts;
    const url = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;

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
        });
      } catch (error) {
        // Network-level failure. Retryable on the same budget.
        if (attempt >= maxAttempts) throw error;
        this.log.debug('graph.network_retry', { attempt, ...describeError(error) });
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

      if (!RETRYABLE_STATUS.has(response.status) || attempt >= maxAttempts) {
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
      if (Number.isFinite(seconds) && seconds >= 0) {
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
