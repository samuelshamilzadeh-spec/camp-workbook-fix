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

  it('does not retry a 403, which means the Sites.Selected grant is wrong', async () => {
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
