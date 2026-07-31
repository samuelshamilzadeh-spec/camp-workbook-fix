import { offsetColumn } from '../domain/cells';
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

/**
 * Graph signals an expired or unknown session through the error code rather
 * than a distinct status. A session is cheap to recreate, so anything
 * session-shaped is treated as recoverable.
 */
function isSessionError(error: unknown): boolean {
  const code = (error as { graphCode?: string })?.graphCode;
  return typeof code === 'string' && /session/i.test(code);
}

export class Workbook {
  private sessionId: string | undefined;
  private sessionPersists: boolean | undefined;

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
   * Opens a workbook session, reusing the existing one when it already matches.
   *
   * This is THE performance lever for this workbook, not a nicety. Measured
   * against the live file, 3.5 MB:
   *
   *   without a session   single cell 8300ms, whole sheet 6613ms
   *   within a session    single cell  195ms, whole sheet  195ms
   *
   * A single cell costing the same as a whole sheet gives it away: the cost is
   * the Excel service opening the workbook, once per request, not the data. A
   * full 46-sheet scan is 66s session-less and 9.3s in a session — of which
   * ~9s is the first call paying the open, and every later call is ~200ms.
   *
   * So the session must OUTLIVE the cycle. Creating and closing one per cycle
   * pays that cold open every five seconds forever, which is what the first
   * implementation did. It is held on the Workbook instance, which lives for
   * the worker process, and Graph keeps a session warm for several minutes of
   * inactivity — far longer than the gap between cycles.
   *
   * A persisted session is required for writes; reads use a non-persisted one,
   * which is faster and structurally incapable of modifying the file.
   */
  async ensureSession(persistChanges: boolean): Promise<void> {
    if (this.sessionId && this.sessionPersists === persistChanges) return;
    if (this.sessionId) await this.closeSession();

    const session = await this.graph.request<{ id: string }>(
      `${this.workbookPath}/createSession`,
      { method: 'POST', body: { persistChanges } },
    );
    this.sessionId = session.id;
    this.sessionPersists = persistChanges;
  }

  /** Back-compat alias. Prefer ensureSession, which reuses. */
  async createSession(persistChanges: boolean): Promise<void> {
    await this.ensureSession(persistChanges);
  }

  /**
   * Runs a request, and if the session turned out to be expired, opens a fresh
   * one and tries once more. `run` reads `sessionHeaders` on each call, so the
   * retry picks up the new id.
   */
  private async withSession<T>(run: () => Promise<T>): Promise<T> {
    try {
      return await run();
    } catch (error) {
      if (!this.sessionId || !isSessionError(error)) throw error;
      const persists = this.sessionPersists ?? false;
      this.sessionId = undefined;
      this.sessionPersists = undefined;
      await this.ensureSession(persists);
      return run();
    }
  }

  /** True when a session is currently open, for logging a cold start. */
  get hasSession(): boolean {
    return this.sessionId !== undefined;
  }

  async closeSession(): Promise<void> {
    if (!this.sessionId) return;
    const id = this.sessionId;
    this.sessionId = undefined;
    this.sessionPersists = undefined;
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
    const result = await this.withSession(() =>
      this.graph.request<{ value: WorksheetInfo[] }>(
        `${this.workbookPath}/worksheets?$select=id,name,position,visibility`,
        { headers: this.sessionHeaders },
      ),
    );
    return result.value;
  }

  /**
   * One call for a whole sheet. Reads are cheap and filtering in memory is free,
   * so this is preferred over targeted per-row reads everywhere.
   */
  async getUsedRange(sheetName: string): Promise<RangeData> {
    return this.withSession(() =>
      this.graph.request<RangeData>(
        `${worksheetPath(this.workbookPath, sheetName)}/usedRange(valuesOnly=true)?$select=address,rowCount,columnCount,values`,
        { headers: this.sessionHeaders },
      ),
    );
  }

  async getRange(sheetName: string, address: string): Promise<RangeData> {
    return this.withSession(() =>
      this.graph.request<RangeData>(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')?$select=address,rowCount,columnCount,values`,
        { headers: this.sessionHeaders },
      ),
    );
  }

  /**
   * The only write path for cell values. Batch multi-row writes into one call:
   * it costs less and it narrows the window in which a concurrent human edit can
   * collide.
   */
  async writeRange(sheetName: string, address: string, values: unknown[][]): Promise<void> {
    assertNoFormulas(values, sheetName, address);
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')`,
        { method: 'PATCH', body: { values }, headers: this.sessionHeaders },
      ),
    );
  }

  /**
   * Inserts `count` blank WHOLE rows at `firstRow`, shifting everything below
   * down.
   *
   * The office's hard constraint is that one row is one patient and nothing may
   * break that alignment. A range insert scoped to the data columns would shift
   * A:Q down and leave BA — the SyncID column — where it was, silently pairing
   * every row below with the previous patient's identifier. So the insert is
   * always applied to the entire row.
   *
   * Two address forms reach the same operation, and which one Graph accepts is
   * not something this codebase can settle from the documentation:
   *
   *   `range(address='A10:A12')/entireRow/insert`   — a range, widened
   *   `range(address='10:12')/insert`               — Excel's whole-row notation
   *
   * The first is tried and the second is the fallback, because a 400 here means
   * Graph rejected the URL before doing anything — nothing is inserted, nothing
   * is half-done, and retrying with the other spelling is safe. `GraphClient`
   * does not retry a 400, so this cannot loop.
   */
  async insertRows(sheetName: string, firstRow: number, count: number): Promise<void> {
    if (count <= 0) return;
    if (!Number.isInteger(firstRow) || firstRow < 1) {
      throw new Error(`Not a row number: ${firstRow}`);
    }
    const lastRow = firstRow + count - 1;
    const sheet = worksheetPath(this.workbookPath, sheetName);

    const insert = (path: string) =>
      this.withSession(() =>
        this.graph.request(path, {
          method: 'POST',
          body: { shift: 'Down' },
          headers: this.sessionHeaders,
        }),
      );

    try {
      await insert(
        `${sheet}/range(address='${encodeURIComponent(`A${firstRow}:A${lastRow}`)}')/entireRow/insert`,
      );
    } catch (error) {
      if ((error as { statusCode?: number })?.statusCode !== 400) throw error;
      await insert(
        `${sheet}/range(address='${encodeURIComponent(`${firstRow}:${lastRow}`)}')/insert`,
      );
    }
  }

  /**
   * Empties cells.
   *
   * This exists because **writing `null` does not clear a cell.** Graph accepts
   * a values PATCH containing nulls, returns 200, and leaves every one of those
   * cells exactly as it was. Verified against the live workbook: three cells
   * holding SyncIDs, PATCHed with `[[null],[null],[null]]`, still held their
   * SyncIDs afterwards; the `/clear` action emptied them.
   *
   * So anything that needs a cell to become empty must come through here. The
   * one that mattered was clearing column B on a daily sheet after a row is
   * marked resolved — with a null write that silently did nothing, leaving the
   * keyword in place and the patient re-appended on the next cycle forever.
   *
   * `Contents` leaves formatting alone, which is what almost every caller wants:
   * the red shading on a blank required field should survive the value being
   * cleared.
   */
  async clearRange(
    sheetName: string,
    address: string,
    applyTo: 'All' | 'Contents' | 'Formats' = 'Contents',
  ): Promise<void> {
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')/clear`,
        { method: 'POST', body: { applyTo }, headers: this.sessionHeaders },
      ),
    );
  }

  async setFill(sheetName: string, address: string, color: string): Promise<void> {
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')/format/fill`,
        { method: 'PATCH', body: { color }, headers: this.sessionHeaders },
      ),
    );
  }

  async clearFill(sheetName: string, address: string): Promise<void> {
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')/format/fill/clear`,
        { method: 'POST', body: {}, headers: this.sessionHeaders },
      ),
    );
  }

  /**
   * Inserts `count` blank WHOLE columns at `firstColumn`, shifting everything to
   * the right of it right.
   *
   * Same reasoning as `insertRows`: scoped to a few rows it would shear the
   * sheet, so it is always the entire column. This is how `Resolved` gets added
   * between `Source Row` and `Last Name` on a tab that already holds patients.
   */
  async insertColumns(sheetName: string, firstColumn: string, count: number): Promise<void> {
    if (count <= 0) return;
    const lastColumn = offsetColumn(firstColumn, count - 1);
    const sheet = worksheetPath(this.workbookPath, sheetName);

    const insert = (path: string) =>
      this.withSession(() =>
        this.graph.request(path, {
          method: 'POST',
          body: { shift: 'Right' },
          headers: this.sessionHeaders,
        }),
      );

    try {
      await insert(
        `${sheet}/range(address='${encodeURIComponent(`${firstColumn}1:${lastColumn}1`)}')/entireColumn/insert`,
      );
    } catch (error) {
      if ((error as { statusCode?: number })?.statusCode !== 400) throw error;
      await insert(
        `${sheet}/range(address='${encodeURIComponent(`${firstColumn}:${lastColumn}`)}')/insert`,
      );
    }
  }

  /**
   * Deletes whole rows, shifting everything below up.
   *
   * The one destructive operation in this codebase. It is reached only from an
   * explicit `Resolved` marker a human typed, never from an inference about what
   * a row looks like.
   */
  async deleteRows(sheetName: string, firstRow: number, count: number): Promise<void> {
    if (count <= 0) return;
    const lastRow = firstRow + count - 1;
    const sheet = worksheetPath(this.workbookPath, sheetName);

    const remove = (path: string) =>
      this.withSession(() =>
        this.graph.request(path, {
          method: 'POST',
          body: { shift: 'Up' },
          headers: this.sessionHeaders,
        }),
      );

    try {
      await remove(
        `${sheet}/range(address='${encodeURIComponent(`A${firstRow}:A${lastRow}`)}')/entireRow/delete`,
      );
    } catch (error) {
      if ((error as { statusCode?: number })?.statusCode !== 400) throw error;
      await remove(`${sheet}/range(address='${encodeURIComponent(`${firstRow}:${lastRow}`)}')/delete`);
    }
  }

  /**
   * Sets a number format over a range.
   *
   * Graph wants a grid matching the range's shape, so the caller passes the row
   * and column counts and this fills it. A date written as a serial without this
   * shows up as `46233`, which is worse than the text it replaced.
   */
  async setNumberFormat(
    sheetName: string,
    address: string,
    format: string,
    rows: number,
    columns = 1,
  ): Promise<void> {
    const numberFormat = Array.from({ length: rows }, () =>
      Array.from({ length: columns }, () => format),
    );
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')`,
        { method: 'PATCH', body: { numberFormat }, headers: this.sessionHeaders },
      ),
    );
  }

  async setFont(
    sheetName: string,
    address: string,
    font: { bold?: boolean; color?: string; size?: number; name?: string; italic?: boolean },
  ): Promise<void> {
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')/format/font`,
        { method: 'PATCH', body: font, headers: this.sessionHeaders },
      ),
    );
  }

  /** Alignment, row height and column width all live on the range's format. */
  async setRangeFormat(
    sheetName: string,
    address: string,
    format: {
      horizontalAlignment?: 'General' | 'Left' | 'Center' | 'Right';
      verticalAlignment?: 'Top' | 'Center' | 'Bottom';
      rowHeight?: number;
      columnWidth?: number;
      wrapText?: boolean;
    },
  ): Promise<void> {
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')/format`,
        { method: 'PATCH', body: format, headers: this.sessionHeaders },
      ),
    );
  }

  async setBorder(
    sheetName: string,
    address: string,
    edge: 'EdgeTop' | 'EdgeBottom' | 'EdgeLeft' | 'EdgeRight' | 'InsideHorizontal' | 'InsideVertical',
    border: { style: string; color?: string; weight?: 'Hairline' | 'Thin' | 'Medium' | 'Thick' },
  ): Promise<void> {
    await this.withSession(() =>
      this.graph.request(
        `${worksheetPath(this.workbookPath, sheetName)}/range(address='${encodeURIComponent(address)}')/format/borders/${edge}`,
        { method: 'PATCH', body: border, headers: this.sessionHeaders },
      ),
    );
  }

  async listNames(): Promise<{ name: string; value: string; scope: string }[]> {
    const result = await this.graph.request<{
      value: { name: string; value: string; scope: string }[];
    }>(`${this.workbookPath}/names?$select=name,value,scope`);
    return result.value;
  }
}

/**
 * Builds the worksheet path segment: `.../workbook/worksheets('Sheet%20Name')`.
 *
 * This returns the whole segment, `worksheets(...)` included, rather than just
 * the parenthesised key. An earlier version returned only `('Name')` and every
 * call site wrote `/worksheets/${encodeSheet(name)}`, which produces
 * `/worksheets/('Name')` — an empty path segment between the slash and the
 * paren. Graph rejects that with a 400 "Empty segment encountered in request
 * URL", which reads like a problem with the workbook rather than with the URL.
 *
 * Keeping the two halves together means they cannot be joined wrongly.
 */
export function worksheetPath(workbookPath: string, sheetName: string): string {
  return `${workbookPath}/worksheets${encodeSheet(sheetName)}`;
}

/**
 * The parenthesised OData key for a sheet name. Excel quotes sheet names and
 * doubles any inner apostrophe. Prefer `worksheetPath` — see above.
 */
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
