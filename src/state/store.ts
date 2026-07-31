import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { BlobServiceClient, type ContainerClient } from '@azure/storage-blob';
import type { RuntimeConfig } from '../config';

/**
 * Checkpoint state.
 *
 * This has to live outside the workbook. The whole point of requirement 6 is
 * that a cycle where nothing changed costs one `lastModifiedDateTime` call and
 * exits — reading a control sheet to find out whether to read the workbook
 * would defeat that.
 */
export interface SyncState {
  /** `lastModifiedDateTime` observed at the end of the previous cycle. */
  lastSeenModified?: string;
  /** ETag observed at the end of the previous cycle. */
  lastSeenETag?: string;
  /**
   * Loop guard: `lastModifiedDateTime` as it stood immediately after our own
   * last write batch. When the file's current value matches this, the only
   * change since last cycle was ours and there is nothing to do.
   */
  lastSelfWriteModified?: string;
  /** Daily sheets touched last cycle. Diagnostic only — written, never read back. */
  lastScannedSheets?: string[];
  /**
   * Rotation position in the cold pool, so successive cycles sweep the rest of
   * the year rather than re-reading the same slice.
   */
  scanCursor?: number;
  /** ISO timestamp of the last cycle that actually opened the workbook. */
  lastFullCycleAt?: string;
}

export interface StateStore {
  read(): Promise<SyncState>;
  write(state: SyncState): Promise<void>;
}

const BLOB_NAME = 'sync-state.json';

class BlobStateStore implements StateStore {
  private container: ContainerClient | undefined;

  constructor(
    private readonly connectionString: string,
    private readonly containerName: string,
  ) {}

  private async getContainer(): Promise<ContainerClient> {
    if (!this.container) {
      const service = BlobServiceClient.fromConnectionString(this.connectionString);
      this.container = service.getContainerClient(this.containerName);
      await this.container.createIfNotExists();
    }
    return this.container;
  }

  async read(): Promise<SyncState> {
    const container = await this.getContainer();
    const blob = container.getBlockBlobClient(BLOB_NAME);
    try {
      const buffer = await blob.downloadToBuffer();
      return JSON.parse(buffer.toString('utf8')) as SyncState;
    } catch (error) {
      if ((error as { statusCode?: number }).statusCode === 404) return {};
      throw error;
    }
  }

  async write(state: SyncState): Promise<void> {
    const container = await this.getContainer();
    const body = JSON.stringify(state);
    await container.getBlockBlobClient(BLOB_NAME).upload(body, Buffer.byteLength(body), {
      blobHTTPHeaders: { blobContentType: 'application/json' },
    });
  }
}

/** Local development fallback. Holds no patient data — timestamps and sheet names only. */
class FileStateStore implements StateStore {
  constructor(private readonly path: string) {}

  async read(): Promise<SyncState> {
    try {
      return JSON.parse(await readFile(this.path, 'utf8')) as SyncState;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return {};
      throw error;
    }
  }

  async write(state: SyncState): Promise<void> {
    await mkdir(dirname(this.path), { recursive: true });
    await writeFile(this.path, JSON.stringify(state, null, 2), 'utf8');
  }
}

export class MemoryStateStore implements StateStore {
  constructor(private state: SyncState = {}) {}
  async read(): Promise<SyncState> {
    return { ...this.state };
  }
  async write(state: SyncState): Promise<void> {
    this.state = { ...state };
  }
}

/**
 * Blob storage when there is a connection string, a local file when there is
 * not — but never a local file when running in Azure.
 *
 * The fallback is for `npm run reconcile` on a laptop. Reached inside a Function
 * App it is silent and permanent damage: a Consumption app's filesystem is
 * read-only under `WEBSITE_RUN_FROM_PACKAGE`, so every `state.write` throws, the
 * checkpoint never persists, and `lastSelfWriteModified` never gets set. That
 * last one is the loop guard — without it the job treats its OWN writes as staff
 * changes and re-plans every five seconds forever, which is the throttling
 * incident the guard exists to prevent.
 *
 * It is also easy to reach by accident: switching `AzureWebJobsStorage` to the
 * identity-based `AzureWebJobsStorage__accountName` form leaves no connection
 * string for this to find. So in Azure it fails loudly at startup instead.
 */
export function createStateStore(config: RuntimeConfig): StateStore {
  if (config.stateConnectionString && !/UseDevelopmentStorage/i.test(config.stateConnectionString)) {
    return new BlobStateStore(config.stateConnectionString, config.stateContainer);
  }

  // Set by App Service on every instance; absent on a developer machine.
  if (process.env.WEBSITE_INSTANCE_ID) {
    throw new Error(
      'No storage connection string, so checkpoint state would fall back to the local ' +
        'filesystem — which is read-only here, making the loop guard silently ineffective. ' +
        'Set AZURE_STORAGE_CONNECTION_STRING or AzureWebJobsStorage (the connection-string ' +
        'form, not AzureWebJobsStorage__accountName).',
    );
  }

  return new FileStateStore(join(process.cwd(), '.state', BLOB_NAME));
}
