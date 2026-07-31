import { describe, expect, it, vi } from 'vitest';
import { GraphClient, GraphError } from '../src/graph/client';
import { createLogger, type LogSink } from '../src/logging';

const silent: LogSink = { debug: () => {}, info: () => {}, warn: () => {}, error: () => {} };
const log = createLogger(silent);
const tokens = { getToken: async () => 'token' };

function response(status: number, body: unknown = {}, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers });
}

describe('GraphClient retries', () => {
  it('retries a 423 from a desktop client lock and then succeeds', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(response(423, { error: { code: 'resourceLocked' } }))
      .mockResolvedValueOnce(response(423, { error: { code: 'resourceLocked' } }))
      .mockResolvedValueOnce(response(200, { ok: true }));

    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
    });

    await expect(client.request('/x')).resolves.toEqual({ ok: true });
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });

  it('retries a 409 conflict', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(response(409, { error: { code: 'conflict' } }))
      .mockResolvedValueOnce(response(200, { ok: true }));

    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
    });

    await expect(client.request('/x')).resolves.toEqual({ ok: true });
  });

  it('honours Retry-After on a 429', async () => {
    const delays: number[] = [];
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(response(429, {}, { 'retry-after': '3' }))
      .mockResolvedValueOnce(response(200, { ok: true }));

    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async (ms) => {
        delays.push(ms);
      },
    });

    await client.request('/x');
    expect(delays).toEqual([3000]);
  });

  it('retries a 501 OpenWorkbookBlockedWorkbook', async () => {
    // The Excel API returns this when a desktop client is holding the file.
    // Observed on 18 of 46 sheets in one scan of the live workbook.
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(response(501, { error: { code: 'OpenWorkbookBlockedWorkbook' } }))
      .mockResolvedValueOnce(response(200, { ok: true }));

    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
    });

    await expect(client.request('/x')).resolves.toEqual({ ok: true });
  });

  it('does not retry a 403, which means admin consent has not been granted', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(response(403, { error: { code: 'accessDenied' } }));
    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
    });

    await expect(client.request('/x')).rejects.toBeInstanceOf(GraphError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('gives up after the retry budget', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(response(423, { error: { code: 'resourceLocked' } }));
    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
      maxAttempts: 3,
    });

    await expect(client.request('/x')).rejects.toMatchObject({ statusCode: 423 });
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });

  it('never puts a response body in the error', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(response(400, { error: { code: 'invalidRequest', message: 'value Smith is bad' } }));
    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
    });

    await expect(client.request('/x')).rejects.toSatisfy(
      (error) => !String((error as Error).message).includes('Smith'),
    );
  });
});

describe('request timeout', () => {
  // Observed on the live workbook: a write issued while two people were editing
  // never returned. Not an error, not a lock — a hang, still hanging 31 minutes
  // later. A failure retries; a hang stops the cycle dead and holds the worker.
  it('gives up on a request that never responds', async () => {
    const fetchImpl = vi.fn().mockImplementation(
      (_url: string, init: RequestInit) =>
        new Promise((_resolve, reject) => {
          init.signal?.addEventListener('abort', () => {
            const error = new Error('timed out');
            error.name = 'TimeoutError';
            reject(error);
          });
        }),
    );

    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
      timeoutMs: 20,
      maxAttempts: 2,
    });

    await expect(client.request('/x')).rejects.toThrow(/timed out/);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('retries a hang and succeeds on the next attempt', async () => {
    let call = 0;
    const fetchImpl = vi.fn().mockImplementation((_url: string, init: RequestInit) => {
      call++;
      if (call === 1) {
        return new Promise((_resolve, reject) => {
          init.signal?.addEventListener('abort', () => {
            const error = new Error('timed out');
            error.name = 'TimeoutError';
            reject(error);
          });
        });
      }
      return Promise.resolve(response(200, { ok: true }));
    });

    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
      timeoutMs: 20,
    });

    await expect(client.request('/x')).resolves.toEqual({ ok: true });
  });

  it('passes an abort signal on every request', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(response(200, {}));
    const client = new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleep: async () => {},
    });
    await client.request('/x');
    expect(fetchImpl.mock.calls[0]?.[1]?.signal).toBeDefined();
  });
});

describe('Retry-After: 0', () => {
  it('backs off anyway instead of spending the whole budget at once', async () => {
    // Graph answers a 503 with `Retry-After: 0`. Taken literally that produced
    // five attempts inside a tenth of a second against the live workbook —
    // every retry fired while the service was still down.
    const slept: number[] = [];
    const client = new GraphClient(tokens, log, {
      fetchImpl: vi.fn().mockResolvedValue(response(503, {}, { 'retry-after': '0' })) as unknown as typeof fetch,
      sleep: async (ms) => {
        slept.push(ms);
      },
    });

    await client.request('/x', { maxAttempts: 5 }).catch(() => {});

    expect(slept).toHaveLength(4);
    expect(slept.every((ms) => ms > 0)).toBe(true);
    // And it grows, so a service that is properly down is not hammered.
    expect(slept[3]!).toBeGreaterThan(slept[0]!);
  });

  it('still honours a real Retry-After', async () => {
    const slept: number[] = [];
    const client = new GraphClient(tokens, log, {
      fetchImpl: vi
        .fn()
        .mockResolvedValueOnce(response(429, {}, { 'retry-after': '2' }))
        .mockResolvedValueOnce(response(200, { ok: true })) as unknown as typeof fetch,
      sleep: async (ms) => {
        slept.push(ms);
      },
    });

    await client.request('/x');
    expect(slept).toEqual([2000]);
  });
});

/**
 * A row insert or delete is not idempotent, and "no answer" is not "did not
 * happen". A delete that succeeded server-side and then timed out, retried,
 * takes the patient who moved up into that row — and the identity check in
 * applyRemovals runs above this layer and has already passed.
 */
describe('a non-idempotent call', () => {
  const client = (fetchImpl: unknown) =>
    new GraphClient(tokens, log, {
      fetchImpl: fetchImpl as typeof fetch,
      sleep: async () => {},
    });

  it('is NOT retried after a 500, because the row may already be gone', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(response(500, { error: { code: 'InternalServerError' } }))
      .mockResolvedValueOnce(response(200, { ok: true }));

    await expect(
      client(fetchImpl).request('/delete', { method: 'POST', idempotent: false }),
    ).rejects.toBeInstanceOf(GraphError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('is NOT retried after a 503 or a 504 either', async () => {
    for (const status of [503, 504]) {
      const fetchImpl = vi
        .fn()
        .mockResolvedValueOnce(response(status, {}))
        .mockResolvedValueOnce(response(200, { ok: true }));

      await expect(
        client(fetchImpl).request('/delete', { method: 'POST', idempotent: false }),
      ).rejects.toBeInstanceOf(GraphError);
      expect(fetchImpl).toHaveBeenCalledTimes(1);
    }
  });

  it('is NOT retried after a client-side timeout', async () => {
    const fetchImpl = vi.fn().mockImplementationOnce(() => {
      const error = new Error('The operation was aborted due to timeout');
      error.name = 'TimeoutError';
      return Promise.reject(error);
    });

    await expect(
      client(fetchImpl).request('/delete', { method: 'POST', idempotent: false }),
    ).rejects.toThrow(/timeout/i);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('IS still retried on a lock or a throttle, which never reached the workbook', async () => {
    for (const status of [409, 423, 429]) {
      const fetchImpl = vi
        .fn()
        .mockResolvedValueOnce(response(status, {}))
        .mockResolvedValueOnce(response(200, { ok: true }));

      await expect(
        client(fetchImpl).request('/delete', { method: 'POST', idempotent: false }),
      ).resolves.toEqual({ ok: true });
      expect(fetchImpl).toHaveBeenCalledTimes(2);
    }
  });

  it('leaves ordinary calls alone: a cell PATCH still retries a 500', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(response(500, {}))
      .mockResolvedValueOnce(response(200, { ok: true }));

    await expect(client(fetchImpl).request('/x', { method: 'PATCH' })).resolves.toEqual({
      ok: true,
    });
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });
});
