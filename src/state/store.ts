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
  /** Identity that last modified the file, so a self-write is recognizable. */
  lastSelfWriteBy?: string;
  /** Daily sheets touched last cycle, used to bound the next scan. */
  lastScannedSheets?: string[];
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

export function createStateStore(config: RuntimeConfig): StateStore {
  if (config.stateConnectionString && !/UseDevelopmentStorage/i.test(config.stateConnectionString)) {
    return new BlobStateStore(config.stateConnectionString, config.stateContainer);
  }
  return new FileStateStore(join(process.cwd(), '.state', BLOB_NAME));
}
