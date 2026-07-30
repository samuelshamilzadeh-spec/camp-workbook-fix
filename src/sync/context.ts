import { loadConfig, type RuntimeConfig } from '../config';
import { createCredential, createTokenProvider } from '../graph/auth';
import { GraphClient } from '../graph/client';
import { Workbook } from '../graph/workbook';
import { createLogger, type LogSink, type Logger } from '../logging';
import { createStateStore, type StateStore } from '../state/store';

export interface SyncContext {
  config: RuntimeConfig;
  workbook: Workbook;
  state: StateStore;
  log: Logger;
}

/**
 * Built once per worker process, not once per invocation. At one invocation
 * every 5 seconds, rebuilding the credential and Blob client each cycle is pure
 * waste, and the token cache only helps if the provider outlives the call.
 */
let cached: SyncContext | undefined;

export function createSyncContext(sink: LogSink, overrides: Partial<RuntimeConfig> = {}): SyncContext {
  const config = { ...loadConfig(), ...overrides };
  const log = createLogger(sink, { phase: config.phase, dryRun: config.dryRun });
  const tokens = createTokenProvider(createCredential(config));
  const graph = new GraphClient(tokens, log);

  return {
    config,
    log,
    workbook: new Workbook(graph, {
      driveId: config.driveId,
      siteId: config.siteId,
      itemId: config.itemId,
    }),
    state: createStateStore(config),
  };
}

export function getSyncContext(sink: LogSink): SyncContext {
  if (!cached) cached = createSyncContext(sink);
  return cached;
}

/** Test seam. */
export function resetSyncContext(): void {
  cached = undefined;
}
