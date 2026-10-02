/**
 * Local persistence.
 *
 * Workbooks are stored in IndexedDB as plain JSON so they survive restarts
 * without a server. Autosave is debounced; `saveNow` forces a write.
 */

import type { Cell, ChartSpec, ConditionalRule, Sheet, Workbook } from '../core/types';
import { unpackCol, unpackRow } from '../core/ref';
import { uid } from '../core/engine';

const DB_NAME = 'sheetcraft';
const DB_VERSION = 1;
const STORE = 'workbooks';
const SETTINGS = 'settings';

export interface StoredWorkbookMeta {
  id: string;
  name: string;
  savedAt: number;
  sheetCount: number;
  cellCount: number;
}

interface SerialisedCell {
  k: number;
  raw?: number | string | boolean | null;
  f?: string;
  s?: number;
  cm?: string;
}

interface SerialisedSheet {
  id: string;
  name: string;
  cells: SerialisedCell[];
  rows: number;
  cols: number;
  cw: [number, number][];
  rh: [number, number][];
  hr: number[];
  hc: number[];
  fr: number;
  fc: number;
  tabColor?: string;
  charts: ChartSpec[];
  condFormats: Sheet['condFormats'];
  dw: number;
  dh: number;
}

interface SerialisedWorkbook {
  id: string;
  name: string;
  sheets: SerialisedSheet[];
  activeSheet: number;
  namedRanges: Record<string, string>;
  styles: Workbook['styles'];
  created: number;
  modified: number;
  version: number;
}

const SCHEMA_VERSION = 1;

/* ------------------------------------------------------------------- shape */

export function serialise(wb: Workbook): SerialisedWorkbook {
  return {
    id: wb.id,
    name: wb.name,
    version: SCHEMA_VERSION,
    activeSheet: wb.activeSheet,
    namedRanges: { ...wb.namedRanges },
    styles: wb.styles,
    created: wb.created,
    modified: wb.modified,
    sheets: wb.sheets.map((s) => ({
      id: s.id,
      name: s.name,
      rows: s.rows,
      cols: s.cols,
      cw: [...s.colWidths.entries()],
      rh: [...s.rowHeights.entries()],
      hr: [...s.hiddenRows],
      hc: [...s.hiddenCols],
      fr: s.frozenRows,
      fc: s.frozenCols,
      tabColor: s.tabColor,
      charts: s.charts,
      condFormats: s.condFormats,
      dw: s.defColWidth,
      dh: s.defRowHeight,
      cells: [...s.cells.entries()].map(([k, cell]) => {
        const out: SerialisedCell = { k };
        if (cell.raw !== undefined) out.raw = cell.raw;
        if (cell.formula !== undefined) out.f = cell.formula;
        if (cell.styleId !== undefined) out.s = cell.styleId;
        if (cell.comment !== undefined) out.cm = cell.comment;
        return out;
      }),
    })),
  };
}

export function deserialise(data: SerialisedWorkbook): Workbook {
  const sheets: Sheet[] = data.sheets.map((s) => {
    const cells = new Map<number, Cell>();
    for (const entry of s.cells ?? []) {
      const cell: Cell = {};
      if (entry.raw !== undefined) cell.raw = entry.raw;
      if (entry.f !== undefined) cell.formula = entry.f;
      if (entry.s !== undefined) cell.styleId = entry.s;
      if (entry.cm !== undefined) cell.comment = entry.cm;
      cells.set(entry.k, cell);
    }
    return {
      id: s.id ?? uid('sh'),
      name: s.name,
      cells,
      rows: s.rows ?? 100,
      cols: s.cols ?? 26,
      colWidths: new Map(s.cw ?? []),
      rowHeights: new Map(s.rh ?? []),
      hiddenRows: new Set(s.hr ?? []),
      hiddenCols: new Set(s.hc ?? []),
      frozenRows: s.fr ?? 0,
      frozenCols: s.fc ?? 0,
      tabColor: s.tabColor,
      charts: s.charts ?? [],
      condFormats: s.condFormats ?? [],
      defColWidth: s.dw ?? 96,
      defRowHeight: s.dh ?? 22,
    };
  });

  return {
    id: data.id ?? uid('wb'),
    name: data.name ?? 'Book1',
    sheets: sheets.length ? sheets : [blankSheet()],
    activeSheet: data.activeSheet ?? 0,
    namedRanges: data.namedRanges ?? {},
    styles: data.styles?.length ? data.styles : [{ id: 0, fmt: {} }],
    revision: 0,
    created: data.created ?? Date.now(),
    modified: data.modified ?? Date.now(),
  };
}

function blankSheet(): Sheet {
  return {
    id: uid('sh'),
    name: 'Sheet1',
    cells: new Map(),
    rows: 100,
    cols: 26,
    colWidths: new Map(),
    rowHeights: new Map(),
    hiddenRows: new Set(),
    hiddenCols: new Set(),
    frozenRows: 0,
    frozenCols: 0,
    charts: [],
    condFormats: [],
    defColWidth: 96,
    defRowHeight: 22,
  };
}

/* ---------------------------------------------------------------- IndexedDB */

let dbPromise: Promise<IDBDatabase> | null = null;

function openDb(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('IndexedDB is not available in this environment'));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        const store = db.createObjectStore(STORE, { keyPath: 'id' });
        store.createIndex('savedAt', 'savedAt');
      }
      if (!db.objectStoreNames.contains(SETTINGS)) {
        db.createObjectStore(SETTINGS, { keyPath: 'key' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('Could not open the local database'));
  });
  return dbPromise;
}

function tx<T>(
  store: string,
  mode: IDBTransactionMode,
  fn: (s: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  return openDb().then(
    (db) =>
      new Promise<T>((resolve, reject) => {
        const transaction = db.transaction(store, mode);
        const req = fn(transaction.objectStore(store));
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error ?? new Error('Local database request failed'));
      }),
  );
}

/* -------------------------------------------------------------------- API */

export async function saveWorkbook(wb: Workbook): Promise<void> {
  const data = serialise(wb);
  await tx(STORE, 'readwrite', (s) => s.put(data));
}

export async function loadWorkbook(id: string): Promise<Workbook | null> {
  const data = await tx<SerialisedWorkbook | undefined>(STORE, 'readonly', (s) => s.get(id));
  return data ? deserialise(data) : null;
}

export async function deleteWorkbook(id: string): Promise<void> {
  await tx(STORE, 'readwrite', (s) => s.delete(id));
}

/** Workbooks ordered by most recent save. */
export async function listWorkbooks(): Promise<StoredWorkbookMeta[]> {
  const all = await tx<SerialisedWorkbook[]>(STORE, 'readonly', (s) => s.getAll());
  return all
    .map((wb) => ({
      id: wb.id,
      name: wb.name,
      savedAt: wb.modified,
      sheetCount: wb.sheets.length,
      cellCount: wb.sheets.reduce((n, sh) => n + (sh.cells?.length ?? 0), 0),
    }))
    .sort((a, b) => b.savedAt - a.savedAt);
}

export async function saveSetting(key: string, value: unknown): Promise<void> {
  await tx(SETTINGS, 'readwrite', (s) => s.put({ key, value }));
}

export async function loadSetting<T>(key: string): Promise<T | null> {
  const row = await tx<{ key: string; value: T } | undefined>(SETTINGS, 'readonly', (s) => s.get(key));
  return row ? row.value : null;
}

/* ------------------------------------------------------------- autosaving */

export class AutoSaver {
  private timer: number | null = null;
  private pending = false;
  private lastError: string | null = null;

  constructor(
    private getWorkbook: () => Workbook,
    private delay = 1200,
  ) {}

  /** Queue a save; repeated calls within the delay collapse into one write. */
  schedule(onSaved?: () => void, onError?: (message: string) => void): void {
    this.pending = true;
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      this.timer = null;
      void this.flush(onSaved, onError);
    }, this.delay) as unknown as number;
  }

  async flush(onSaved?: () => void, onError?: (message: string) => void): Promise<void> {
    if (!this.pending) return;
    this.pending = false;
    if (this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    try {
      await saveWorkbook(this.getWorkbook());
      this.lastError = null;
      onSaved?.();
    } catch (e) {
      this.lastError = (e as Error).message;
      onError?.(this.lastError);
    }
  }

  get error(): string | null {
    return this.lastError;
  }

  dispose(): void {
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
  }
}

/** Rows/cols used by a sheet, for the workbook list. */
export function sheetExtent(sheet: Sheet): { rows: number; cols: number } {
  let maxR = 0;
  let maxC = 0;
  for (const k of sheet.cells.keys()) {
    const r = unpackRow(k);
    const c = unpackCol(k);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  return { rows: maxR + 1, cols: maxC + 1 };
}

export type { ConditionalRule };
