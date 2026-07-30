import type { GraphClient } from './client';

export interface FileMetadata {
  lastModifiedDateTime: string;
  eTag: string | undefined;
  lastModifiedBy: string | undefined;
}

export interface WorksheetInfo {
  id: string;
  name: string;
  position: number;
  visibility: string;
}

export interface RangeData {
  address: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
}

/**
 * Thin wrapper over the Graph Workbook API for one driveItem.
 *
 * Two things this deliberately does not do:
 *   - It never writes a formula. Every write goes through `values`, and
 *     `assertNoFormulas` rejects any payload whose cells start with `=`. The
 *     brief calls this the single most important constraint in the project, so
 *     it is enforced at the one chokepoint every write passes through rather
 *     than left to reviewer discipline.
 *   - It never logs a range payload.
 */
export interface WorkbookLocation {
  /**
   * The drive holding the workbook. Preferred, because `/drives/{id}/items/{id}`
   * addresses a file wherever it lives — a site's default library, a second
   * library on the same site, a Teams channel, or someone's OneDrive.
   */
  driveId?: string;
  /**
   * Fallback for when only the site is known. This resolves to the site's
   * DEFAULT document library, so it silently misses a workbook stored in any
   * other library on that site. Use `npm run resolve` to get a driveId instead.
   */
  siteId?: string;
  itemId: string;
}

export class Workbook {
  private sessionId: string | undefined;

  constructor(
    private readonly graph: GraphClient,
    private readonly location: WorkbookLocation,
  ) {
    if (!location.driveId && !location.siteId) {
      throw new Error(
        'Workbook needs either GRAPH_DRIVE_ID (preferred) or GRAPH_SITE_ID. ' +
          'Run `npm run resolve -- "<workbook URL>"` to get both.',
      );
    }
    if (!location.itemId) {
      throw new Error('Workbook needs GRAPH_ITEM_ID. Run `npm run resolve -- "<workbook URL>"`.');
    }
  }

  private get itemPath(): string {
    return this.location.driveId
      ? `/drives/${this.location.driveId}/items/${this.location.itemId}`
      : `/sites/${this.location.siteId}/drive/items/${this.location.itemId}`;
  }

  private get workbookPath(): string {
    return `${this.itemPath}/workbook`;
  }

  private get sessionHeaders(): Record<string, string> {
    return this.sessionId ? { 'workbook-session-id': this.sessionId } : {};
  }

  /**
   * The cheap change check. One call, no workbook session, no sheet opened.
   * Most of the ~17,000 daily invocations should stop right after this.
   */
  async getFileMetadata(): Promise<FileMetadata> {
    const item = await this.graph.request<{
      lastModifiedDateTime: string;
      eTag?: string;
      lastModifiedBy?: { user?: { id?: string; displayName?: string } };
    }>(`${this.itemPath}?$select=lastModifiedDateTime,eTag,lastModifiedBy`);

    return {
      lastModifiedDateTime: item.lastModifiedDateTime,
      eTag: item.eTag,
      // The user id, not the display name — a display name is a person's name.
      lastModifiedBy: item.lastModifiedBy?.user?.id,
    };
  }

  /**
   * A persisted session batches our writes against one workbook instance and
   * keeps Graph from re-loading the file on every call. Non-persisted sessions
   * discard changes, so read-only phases use one too — it is faster and cannot
   * modify anything.
   */
  async createSession(persistChanges: boolean): Promise<void> {
    const session = await this.graph.request<{ id: string }>(
      `${this.workbookPath}/createSession`,
      { method: 'POST', body: { persistChanges } },
    );
    this.sessionId = session.id;
  }

  async closeSession(): Promise<void> {
    if (!this.sessionId) return;
    const id = this.sessionId;
    this.sessionId = undefined;
    try {
      await this.graph.request(`${this.workbookPath}/closeSession`, {
        method: 'POST',
        headers: { 'workbook-session-id': id },
        maxAttempts: 2,
      });
    } catch {
      // A session we failed to close expires on its own. Never fail a cycle here.
    }
  }

  async listWorksheets(): Promise<WorksheetInfo[]> {
    const result = await this.graph.request<{ value: WorksheetInfo[] }>(
      `${this.workbookPath}/worksheets?$select=id,name,position,visibility`,
      { headers: this.sessionHeaders },
    );
    return result.value;
  }

  /**
   * One call for a whole sheet. Reads are cheap and filtering in memory is free,
   * so this is preferred over targeted per-row reads everywhere.
   */
  async getUsedRange(sheetName: string): Promise<RangeData> {
    return this.graph.request<RangeData>(
      `${this.workbookPath}/worksheets/${encodeSheet(sheetName)}/usedRange(valuesOnly=true)?$select=address,rowCount,columnCount,values`,
      { headers: this.sessionHeaders },
    );
  }

  async getRange(sheetName: string, address: string): Promise<RangeData> {
    return this.graph.request<RangeData>(
      `${this.workbookPath}/worksheets/${encodeSheet(sheetName)}/range(address='${encodeURIComponent(address)}')?$select=address,rowCount,columnCount,values`,
      { headers: this.sessionHeaders },
    );
  }

  /**
   * The only write path for cell values. Batch multi-row writes into one call:
   * it costs less and it narrows the window in which a concurrent human edit can
   * collide.
   */
  async writeRange(sheetName: string, address: string, values: unknown[][]): Promise<void> {
    assertNoFormulas(values, sheetName, address);
    await this.graph.request(
      `${this.workbookPath}/worksheets/${encodeSheet(sheetName)}/range(address='${encodeURIComponent(address)}')`,
      { method: 'PATCH', body: { values }, headers: this.sessionHeaders },
    );
  }

  async setFill(sheetName: string, address: string, color: string): Promise<void> {
    await this.graph.request(
      `${this.workbookPath}/worksheets/${encodeSheet(sheetName)}/range(address='${encodeURIComponent(address)}')/format/fill`,
      { method: 'PATCH', body: { color }, headers: this.sessionHeaders },
    );
  }

  async clearFill(sheetName: string, address: string): Promise<void> {
    await this.graph.request(
      `${this.workbookPath}/worksheets/${encodeSheet(sheetName)}/range(address='${encodeURIComponent(address)}')/format/fill/clear`,
      { method: 'POST', body: {}, headers: this.sessionHeaders },
    );
  }

  async listNames(): Promise<{ name: string; value: string; scope: string }[]> {
    const result = await this.graph.request<{
      value: { name: string; value: string; scope: string }[];
    }>(`${this.workbookPath}/names?$select=name,value,scope`);
    return result.value;
  }
}

/** Excel sheet names in a Graph path are quoted, and inner quotes are doubled. */
export function encodeSheet(sheetName: string): string {
  return `('${encodeURIComponent(sheetName.replace(/'/g, "''"))}')`;
}

/**
 * The workbook is being rescued from a formula load that made it unusable.
 * Writing one back — even accidentally, even a stray `=` a staff member typed
 * that we round-tripped through write-back — would reintroduce the problem this
 * project exists to solve.
 */
export function assertNoFormulas(values: unknown[][], sheetName: string, address: string): void {
  for (const row of values) {
    for (const cell of row) {
      // Only a leading `=` is rejected. Excel also treats a leading `+` or `@`
      // as formula-ish, but `+1 555 …` is a perfectly ordinary phone number and
      // refusing to sync it would be worse than the risk. See the open question
      // in the README about how Excel stores such a value on write-back.
      if (typeof cell === 'string' && /^\s*=/.test(cell)) {
        throw new Error(
          `Refusing to write a formula-like value to ${sheetName}!${address}. ` +
            'This codebase never writes formulas; see src/graph/workbook.ts.',
        );
      }
    }
  }
}
