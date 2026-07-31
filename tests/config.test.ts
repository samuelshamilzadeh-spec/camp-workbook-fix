import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { loadConfig } from '../src/config';
import { createStateStore } from '../src/state/store';

/**
 * Settings that fail loudly.
 *
 * A number setting that silently became NaN did not throw, it degraded — and
 * every one of these degradations is invisible in the logs and costs patients
 * their queue row.
 */

const REQUIRED = {
  AZURE_TENANT_ID: 't',
  AZURE_CLIENT_ID: 'c',
  GRAPH_DRIVE_ID: 'd',
  GRAPH_ITEM_ID: 'i',
};

let saved: NodeJS.ProcessEnv;

beforeEach(() => {
  saved = { ...process.env };
  for (const key of Object.keys(process.env)) {
    if (key.startsWith('SYNC_') || key.startsWith('GRAPH_') || key.startsWith('AZURE_')) {
      delete process.env[key];
    }
  }
  Object.assign(process.env, REQUIRED);
});

afterEach(() => {
  process.env = saved;
});

describe('numeric settings', () => {
  it('take their documented defaults', () => {
    const config = loadConfig();
    expect(config.hotDaysBack).toBe(7);
    expect(config.hotDaysForward).toBe(14);
    expect(config.coldBatchSize).toBe(10);
    expect(config.maxSheetsPerCycle).toBe(90);
    expect(config.readConcurrency).toBe(8);
  });

  it('accept a valid override', () => {
    process.env.SYNC_MAX_SHEETS_PER_CYCLE = '120';
    expect(loadConfig().maxSheetsPerCycle).toBe(120);
  });

  it('refuse a value that is not a number', () => {
    // Silently, this made `daily.length <= NaN` false, switched the scan to
    // tiering, then took NaN sheets per rotation — so every sheet outside the
    // hot window stopped being read at all.
    process.env.SYNC_MAX_SHEETS_PER_CYCLE = 'ninety';
    expect(() => loadConfig()).toThrow(/SYNC_MAX_SHEETS_PER_CYCLE must be an integer/);
  });

  it('refuse a fraction and a negative', () => {
    process.env.SYNC_COLD_BATCH_SIZE = '2.5';
    expect(() => loadConfig()).toThrow(/SYNC_COLD_BATCH_SIZE/);

    delete process.env.SYNC_COLD_BATCH_SIZE;
    process.env.SYNC_HOT_DAYS_BACK = '-1';
    expect(() => loadConfig()).toThrow(/SYNC_HOT_DAYS_BACK/);
  });

  it('refuse a read concurrency that would just get us throttled', () => {
    process.env.SYNC_READ_CONCURRENCY = '64';
    expect(() => loadConfig()).toThrow(/SYNC_READ_CONCURRENCY/);
  });

  it('still refuses a bad SYNC_PHASE, as it always has', () => {
    process.env.SYNC_PHASE = '9';
    expect(() => loadConfig()).toThrow(/SYNC_PHASE/);
  });
});

describe('the checkpoint store', () => {
  const config = (overrides: Record<string, unknown> = {}) =>
    ({ stateContainer: 'camp-sync-state', stateConnectionString: undefined, ...overrides }) as any;

  it('falls back to a local file on a developer machine', () => {
    delete process.env.WEBSITE_INSTANCE_ID;
    expect(() => createStateStore(config())).not.toThrow();
  });

  it('refuses to fall back to a local file inside Azure', () => {
    // The filesystem is read-only there, so every state.write throws,
    // `lastSelfWriteModified` never persists, and the job treats its own writes
    // as staff changes every five seconds forever.
    process.env.WEBSITE_INSTANCE_ID = 'instance-0';
    try {
      expect(() => createStateStore(config())).toThrow(/read-only/);
    } finally {
      delete process.env.WEBSITE_INSTANCE_ID;
    }
  });

  it('uses blob storage when a connection string is present, in Azure or out', () => {
    process.env.WEBSITE_INSTANCE_ID = 'instance-0';
    try {
      const store = createStateStore(
        config({
          stateConnectionString:
            'DefaultEndpointsProtocol=https;AccountName=a;AccountKey=Zm9v;EndpointSuffix=core.windows.net',
        }),
      );
      expect(store).toBeDefined();
    } finally {
      delete process.env.WEBSITE_INSTANCE_ID;
    }
  });
});
