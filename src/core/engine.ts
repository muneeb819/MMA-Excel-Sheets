/**
 * The workbook engine: cell storage, dependency-tracked recalculation,
 * dynamic-array spilling and undo/redo.
 *
 * Cells are held sparsely (`Map<packedCoord, Cell>`). Formula results are
 * cached; edits mark the edited cell and its transitive dependents dirty, then
 * a single recalculation pass rebuilds everything in dependency order.
 */

import {
  DEFAULT_COL_W,
  DEFAULT_ROW_H,
  err,
  isError,
  type Cell,
  type CellFormat,
  type ChartSpec,
  type Range,
  type Sheet,
  type Style,
  type Value,
  type Workbook,
} from './types';
import { MAX_COLS, MAX_ROWS } from './types';
import {
  colName,
  normRange,
  pack,
  quoteSheet,
  rangeToA1,
  translateFormula,
  unpackCol,
  unpackRow,
} from './ref';
import { formatValue, parseInput, toText } from './coerce';
import { evalFormulaRaw, type Resolver } from './formula/evaluator';
import { isArray, isRange, type FnRegistry } from './formula/fntypes';
import { buildRegistry } from './formula/functions/index';

let idCounter = 0;
export function uid(prefix: string): string {
  idCounter += 1;
  return `${prefix}_${Date.now().toString(36)}_${idCounter.toString(36)}`;
}

/* ------------------------------------------------------------ construction */

export function newSheet(name: string): Sheet {
  return {
    id: uid('sh'),
    name,
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
    defColWidth: DEFAULT_COL_W,
    defRowHeight: DEFAULT_ROW_H,
  };
}

export function newWorkbook(name = 'Book1'): Workbook {
  return {
    id: uid('wb'),
    name,
    sheets: [newSheet('Sheet1')],
    activeSheet: 0,
    namedRanges: {},
    styles: [{ id: 0, fmt: {} }],
    revision: 0,
    created: Date.now(),
    modified: Date.now(),
  };
}

/* ------------------------------------------------------------------ engine */

interface SheetSnapshot {
  id: string;
  name: string;
  cells: Map<number, Cell>;
  rows: number;
  cols: number;
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  hiddenRows: number[];
  hiddenCols: number[];
  frozenRows: number;
  frozenCols: number;
  tabColor?: string;
  charts: ChartSpec[];
  condFormats: Sheet['condFormats'];
  defColWidth: number;
  defRowHeight: number;
}

interface Snapshot {
  sheets: SheetSnapshot[];
  activeSheet: number;
  namedRanges: Record<string, string>;
}

export class Engine {
  wb: Workbook;
  readonly fns: FnRegistry;

  /** formula results, keyed `${sheetId}:${r}:${c}` */
  private cache = new Map<string, Value>();
  /** precedent cell -> set of dependent cells */
  private dependents = new Map<string, Set<string>>();
  /** `${sheetId}:${r}:${c}` -> the cell that spilled into it */
  private spillOwner = new Map<string, string>();
  private dirtySet = new Set<string>();

  private undoStack: Snapshot[] = [];
  private redoStack: Snapshot[] = [];
  private historyLimit = 120;
  private batching = 0;

  /** listeners notified after every recalculation so the UI can repaint */
  revisionListeners = new Set<() => void>();

  constructor(wb: Workbook = newWorkbook(), fns: FnRegistry = buildRegistry()) {
    this.wb = wb;
    this.fns = fns;
  }

  /* ------------------------------------------------------------ accessors */

  get active(): Sheet {
    return this.wb.sheets[this.wb.activeSheet] ?? this.wb.sheets[0];
  }

  sheetIndex(idOrName: string): number {
    return this.wb.sheets.findIndex(
      (s) => s.id === idOrName || s.name.toLowerCase() === idOrName.toLowerCase(),
    );
  }

  sheetById(id: string): Sheet | undefined {
    return this.wb.sheets.find((s) => s.id === id);
  }

  sheetByName(name: string): Sheet | null {
    const want = name.toLowerCase();
    return this.wb.sheets.find((s) => s.name.toLowerCase() === want) ?? null;
  }

  cellAt(sheet: Sheet, r: number, c: number): Cell | undefined {
    return sheet.cells.get(pack(r, c));
  }

  /** Value of a cell, including dynamic-array spill results. */
  valueAt(sheet: Sheet, r: number, c: number): Value {
    const k = key(sheet.id, r, c);
    const cell = sheet.cells.get(pack(r, c));
    if (cell) {
      if (cell.formula !== undefined) {
        const cached = this.cache.get(k);
        return cached === undefined ? null : cached;
      }
      return cell.raw ?? null;
    }
    return this.cache.get(k) ?? null;
  }

  /** Text shown in the formula bar for a cell. */
  editText(sheet: Sheet, r: number, c: number): string {
    const cell = sheet.cells.get(pack(r, c));
    if (!cell) return '';
    if (cell.formula !== undefined) return `=${cell.formula}`;
    if (cell.raw === null || cell.raw === undefined) return '';
    if (typeof cell.raw === 'boolean') return cell.raw ? 'TRUE' : 'FALSE';
    return toText(cell.raw);
  }

  /** Display string honouring the cell's number format. */
  displayAt(sheet: Sheet, r: number, c: number): string {
    const cell = sheet.cells.get(pack(r, c));
    const fmt = cell?.styleId !== undefined ? this.formatOf(cell.styleId) : undefined;
    return formatValue(this.valueAt(sheet, r, c), fmt?.numFmt);
  }

  formatOf(styleId: number | undefined): CellFormat | undefined {
    if (styleId === undefined) return undefined;
    return this.wb.styles.find((s) => s.id === styleId)?.fmt;
  }

  /** Intern a format into the style table, returning its id. */
  styleIdFor(fmt: CellFormat): number {
    const keyStr = JSON.stringify(fmt);
    const found = this.wb.styles.find((s) => JSON.stringify(s.fmt) === keyStr);
    if (found) return found.id;
    const id = this.wb.styles.length;
    this.wb.styles.push({ id, fmt });
    return id;
  }

  /* ------------------------------------------------------------- mutations */

  /** Set a cell from raw user input; parses formulas and literal types. */
  setInput(sheet: Sheet, r: number, c: number, text: string): void {
    const k = pack(r, c);
    const existing = sheet.cells.get(k);
    const cell: Cell = { ...(existing ?? {}), cached: undefined, dirty: undefined };

    if (text.startsWith('=')) {
      cell.formula = text.slice(1);
      delete cell.raw;
    } else {
      delete cell.formula;
      const parsed = parseInput(text);
      if (parsed === null) delete cell.raw;
      else cell.raw = parsed;
    }

    if (
      cell.raw === undefined &&
      cell.formula === undefined &&
      cell.comment === undefined &&
      cell.styleId === undefined
    ) {
      sheet.cells.delete(k);
    } else {
      sheet.cells.set(k, cell);
    }
    this.markDirty(sheet, r, c);
    this.touchExtent(sheet, r, c);
  }

  /** Set an explicit value (imports, data tools, chart data loading). */
  setValue(sheet: Sheet, r: number, c: number, value: Value, styleId?: number): void {
    const k = pack(r, c);
    const cell: Cell = { ...(this.wb ? sheet.cells.get(k) : undefined) };
    if (styleId !== undefined) cell.styleId = styleId;
    cell.cached = undefined;

    if (isError(value)) {
      delete cell.formula;
      delete cell.raw;
      cell.cached = value;
    } else if (typeof value === 'string' && value.startsWith('=')) {
      delete cell.raw;
      cell.formula = value.slice(1);
    } else {
      delete cell.formula;
      if (value === null) delete cell.raw;
      else if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
        cell.raw = value;
      } else {
        delete cell.raw;
        cell.cached = value;
      }
    }

    sheet.cells.set(k, cell);
    this.markDirty(sheet, r, c);
    this.touchExtent(sheet, r, c);
  }

  setStyleId(sheet: Sheet, r: number, c: number, styleId: number | undefined): void {
    const k = pack(r, c);
    const existing = sheet.cells.get(k);
    if (existing) {
      existing.styleId = styleId;
      sheet.cells.set(k, existing);
    } else if (styleId !== undefined) {
      sheet.cells.set(k, { styleId });
    }
  }

  /** Remove values (and styles) across a range. */
  clearRange(sheet: Sheet, range: Range): void {
    const n = normRange(range);
    for (let r = n.r1; r <= n.r2; r++) {
      for (let c = n.c1; c <= n.c2; c++) {
        const k = pack(r, c);
        const cell = sheet.cells.get(k);
        if (cell) {
          delete cell.raw;
          delete cell.formula;
          delete cell.comment;
          if (cell.styleId === undefined) sheet.cells.delete(k);
          else sheet.cells.set(k, cell);
        }
        this.cache.delete(key(sheet.id, r, c));
        this.spillOwner.delete(key(sheet.id, r, c));
      }
    }
    this.markDirtyRange(sheet, n);
  }

  /**
   * Copy `source` across `target`, tiling it when the target is larger, and
   * shift relative references by the offset between the two blocks.
   */
  copyRange(
    sheet: Sheet,
    target: Range,
    source: Range,
    mode: 'copy' | 'move' = 'copy',
    translateRefs = true,
  ): void {
    const t = normRange(target);
    const s = normRange(source);
    const height = s.r2 - s.r1 + 1;
    const width = s.c2 - s.c1 + 1;

    // Snapshot first so overlapping copy/paste behaves predictably.
    const block: (Cell | undefined)[] = [];
    for (let r = s.r1; r <= s.r2; r++) {
      for (let c = s.c1; c <= s.c2; c++) block.push(sheet.cells.get(pack(r, c)));
    }

    let i = 0;
    for (let r = t.r1; r <= t.r2; r++) {
      for (let c = t.c1; c <= t.c2; c++) {
        const src = block[i % block.length];
        i++;
        // Which cell of the source block this destination corresponds to.
        const srcRow = s.r1 + ((r - t.r1) % height);
        const srcCol = s.c1 + ((c - t.c1) % width);
        this.writeClonedCell(sheet, r, c, src, translateRefs, r - srcRow, c - srcCol);
      }
    }

    if (mode === 'move') this.clearRange(sheet, s);
    this.markDirtyRange(sheet, t);
  }

  /**
   * Write one cloned cell. `dr`/`dc` are how far the destination sits from the
   * source cell, which is what relative references are shifted by.
   */
  private writeClonedCell(
    sheet: Sheet,
    r: number,
    c: number,
    src: Cell | undefined,
    translateRefs: boolean,
    dr: number,
    dc: number,
  ): void {
    const k = pack(r, c);
    if (!src) {
      sheet.cells.delete(k);
      this.markDirty(sheet, r, c);
      return;
    }
    const cell: Cell = { ...src, cached: undefined, dirty: undefined };
    if (translateRefs && cell.formula !== undefined) {
      cell.formula = translateFormula(cell.formula, dr, dc);
    }
    sheet.cells.set(k, cell);
    this.markDirty(sheet, r, c);
  }

  /* ---------------------------------------------------- structural changes */

  insertRows(sheet: Sheet, at: number, count: number): void {
    sheet.cells = shiftCells(sheet.cells, at, 0, count, 1, count);
    sheet.rowHeights = shiftMap(sheet.rowHeights, at, count, 1);
    shiftSet(sheet.hiddenRows, at, count, 1);
    sheet.frozenRows = sheet.frozenRows >= at ? sheet.frozenRows + count : sheet.frozenRows;
    this.invalidateAll();
  }

  deleteRows(sheet: Sheet, at: number, count: number): void {
    sheet.cells = shiftCells(sheet.cells, at, 0, count, -1, count);
    sheet.rowHeights = shiftMap(sheet.rowHeights, at, count, -1);
    shiftSet(sheet.hiddenRows, at, count, -1);
    sheet.frozenRows = Math.max(0, sheet.frozenRows > at ? sheet.frozenRows - count : sheet.frozenRows);
    this.invalidateAll();
  }

  insertCols(sheet: Sheet, at: number, count: number): void {
    sheet.cells = shiftCells(sheet.cells, 0, at, count, 1, count);
    sheet.colWidths = shiftMap(sheet.colWidths, at, count, 1);
    shiftSet(sheet.hiddenCols, at, count, 1);
    sheet.frozenCols = sheet.frozenCols >= at ? sheet.frozenCols + count : sheet.frozenCols;
    this.invalidateAll();
  }

  deleteCols(sheet: Sheet, at: number, count: number): void {
    sheet.cells = shiftCells(sheet.cells, 0, at, count, -1, count);
    sheet.colWidths = shiftMap(sheet.colWidths, at, count, -1);
    shiftSet(sheet.hiddenCols, at, count, -1);
    sheet.frozenCols = Math.max(0, sheet.frozenCols > at ? sheet.frozenCols - count : sheet.frozenCols);
    this.invalidateAll();
  }

  /* ------------------------------------------------------------ sheet ops */

  addSheet(name?: string, at = this.wb.sheets.length): Sheet {
    const sheet = newSheet(
      uniqueSheetName(this.wb, name ?? `Sheet${this.wb.sheets.length + 1}`),
    );
    this.wb.sheets.splice(at, 0, sheet);
    this.wb.activeSheet = at;
    this.invalidateAll();
    return sheet;
  }

  removeSheet(sheet: Sheet): void {
    if (this.wb.sheets.length === 1) return;
    const i = this.wb.sheets.indexOf(sheet);
    if (i < 0) return;
    this.wb.sheets.splice(i, 1);
    this.wb.activeSheet = Math.max(0, Math.min(this.wb.sheets.length - 1, this.wb.activeSheet));
    this.invalidateAll();
  }

  renameSheet(sheet: Sheet, name: string): void {
    const clean = name.trim() || sheet.name;
    if (this.wb.sheets.some((s) => s !== sheet && s.name.toLowerCase() === clean.toLowerCase())) return;
    sheet.name = clean;
    this.invalidateAll();
  }

  duplicateSheet(sheet: Sheet): Sheet {
    const copy = newSheet(uniqueSheetName(this.wb, `${sheet.name} (2)`));
    copy.cells = new Map(sheet.cells);
    copy.rows = sheet.rows;
    copy.cols = sheet.cols;
    copy.colWidths = new Map(sheet.colWidths);
    copy.rowHeights = new Map(sheet.rowHeights);
    copy.charts = sheet.charts.map((c) => ({ ...c, id: uid('ch'), series: c.series.map((x) => ({ ...x })) }));
    copy.condFormats = sheet.condFormats.map((f) => ({ ...f }));
    copy.frozenRows = sheet.frozenRows;
    copy.frozenCols = sheet.frozenCols;
    copy.defColWidth = sheet.defColWidth;
    copy.defRowHeight = sheet.defRowHeight;
    const i = this.wb.sheets.indexOf(sheet);
    this.wb.sheets.splice(i + 1, 0, copy);
    this.wb.activeSheet = i + 1;
    this.invalidateAll();
    return copy;
  }

  moveSheet(from: number, to: number): void {
    if (from === to || from < 0 || from >= this.wb.sheets.length) return;
    const [s] = this.wb.sheets.splice(from, 1);
    this.wb.sheets.splice(Math.max(0, Math.min(this.wb.sheets.length, to)), 0, s);
    this.wb.activeSheet = this.wb.sheets.indexOf(this.active);
  }

  setNamedRange(name: string, sheet: Sheet, range: Range): void {
    this.wb.namedRanges[name] = `${quoteSheet(sheet.name)}!${rangeToA1(normRange(range), true)}`;
    this.invalidateAll();
  }

  deleteNamedRange(name: string): void {
    delete this.wb.namedRanges[name];
    this.invalidateAll();
  }

  /* --------------------------------------------------------- recalculation */

  markDirty(sheet: Sheet, r: number, c: number): void {
    const start = key(sheet.id, r, c);
    const stack = [start];
    const seen = new Set<string>();
    while (stack.length) {
      const k = stack.pop()!;
      if (seen.has(k)) continue;
      seen.add(k);
      this.dirtySet.add(k);
      this.cache.delete(k);
      const deps = this.dependents.get(k);
      if (deps) for (const d of deps) stack.push(d);
    }
  }

  markDirtyRange(sheet: Sheet, range: Range): void {
    const n = normRange(range);
    const cells = (n.r2 - n.r1 + 1) * (n.c2 - n.c1 + 1);
    if (cells > 400_000) {
      this.invalidateAll();
      return;
    }
    for (let r = n.r1; r <= n.r2; r++) {
      for (let c = n.c1; c <= n.c2; c++) this.markDirty(sheet, r, c);
    }
  }

  /** Recompute every dirty cell. Safe to call as often as needed. */
  recalc(): void {
    if (this.dirtySet.size === 0) return;
    this.recalcKeys([...this.dirtySet]);
  }

  /** Recompute every formula cell in the workbook. */
  recalcAll(): void {
    const all: string[] = [];
    for (const sheet of this.wb.sheets) {
      for (const [k, cell] of sheet.cells) {
        if (cell.formula !== undefined) {
          all.push(key(sheet.id, unpackRow(k), unpackCol(k)));
        }
      }
    }
    this.cache.clear();
    this.spillOwner.clear();
    this.recalcKeys(all);
  }

  private recalcKeys(keys: string[]): void {
    for (const k of keys) this.cache.delete(k);

    // Evaluate in dependency order where known; `compute` recurses into
    // precedents anyway, so this only matters for spill footprints.
    const ordered = this.topoSort(keys);
    for (const k of ordered) {
      const loc = parseKey(k);
      if (!loc) continue;
      const sheet = this.sheetById(loc.sheetId);
      if (!sheet) continue;
      this.compute(sheet, loc.r, loc.c, new Set());
    }

    for (const k of keys) this.dirtySet.delete(k);
    this.wb.revision += 1;
    this.wb.modified = Date.now();
    this.emit();
  }

  /** DFS over the dependent graph so precedents evaluate first. */
  private topoSort(keys: string[]): string[] {
    const out: string[] = [];
    const visited = new Set<string>();
    const active = new Set<string>();
    const target = new Set(keys);

    const visit = (k: string): void => {
      if (visited.has(k) || active.has(k)) return;
      active.add(k);
      const deps = this.dependents.get(k);
      if (deps) for (const d of deps) if (target.has(d)) visit(d);
      active.delete(k);
      visited.add(k);
      out.push(k);
    };

    for (const k of keys) visit(k);
    return out;
  }

  /**
   * Value of a cell, computing formulas on demand.
   *
   * Recursion means evaluation order does not depend on the dependency graph
   * being complete, and `stack` turns genuine cycles into #CIRCULAR! instead of
   * blowing the JavaScript call stack.
   */
  private compute(sheet: Sheet, r: number, c: number, stack: Set<string>): Value {
    const k = key(sheet.id, r, c);

    const cached = this.cache.get(k);
    if (cached !== undefined) return cached;

    const cell = sheet.cells.get(pack(r, c));
    if (!cell) return null;
    if (cell.formula === undefined) return cell.raw ?? null;

    if (stack.has(k)) return err('CYCLE', 'Circular reference');
    stack.add(k);

    // Record this cell's precedents so later edits can invalidate it.
    const refs: string[] = [];
    const resolver = this.resolverFor(sheet, stack, (s, r1, c1, r2, c2) => {
      if (!s) return;
      const own = s === sheet;
      const rEnd = own ? r2 : r1;
      const cEnd = own ? c2 : c1;
      for (let rr = r1; rr <= rEnd; rr++) {
        for (let cc = c1; cc <= cEnd; cc++) {
          if (own && rr === r && cc === c) continue;
          const pk = key(s.id, rr, cc);
          refs.push(pk);
          let set = this.dependents.get(pk);
          if (!set) {
            set = new Set();
            this.dependents.set(pk, set);
          }
          set.add(k);
        }
      }
    });

    const result = evalFormulaRaw(cell.formula, sheet, resolver, this.fns);

    if (isArray(result)) {
      if (this.applySpill(sheet, r, c, result.rows, k)) {
        stack.delete(k);
        this.pruneDependents(k, refs);
        return this.cache.get(k) ?? null;
      }
      this.cache.set(k, err('SPILL', 'Spill range is not empty'));
      stack.delete(k);
      this.pruneDependents(k, refs);
      return this.cache.get(k) as Value;
    }

    if (isRange(result)) {
      // OFFSET/INDIRECT hand back a live reference; take its anchor cell.
      const v = result.sheet ? this.compute(result.sheet, result.r1, result.c1, stack) : null;
      this.cache.set(k, v);
    } else {
      this.cache.set(k, result);
    }

    this.pruneDependents(k, refs);
    stack.delete(k);
    return this.cache.get(k) ?? null;
  }

  /** Drop precedent edges that this cell no longer reads. */
  private pruneDependents(k: string, after: string[]): void {
    for (const bk of this.dependents.keys()) {
      const set = this.dependents.get(bk);
      if (set && set.has(k) && !after.includes(bk)) {
        set.delete(k);
        if (set.size === 0) this.dependents.delete(bk);
      }
    }
  }

  /** Write an array result into neighbouring cells; false when blocked. */
  private applySpill(
    sheet: Sheet,
    r: number,
    c: number,
    rows: Value[][],
    ownerKey: string,
  ): boolean {
    const height = rows.length;
    const width = rows[0]?.length ?? 1;

    if (height === 1 && width === 1) {
      this.releaseSpill(ownerKey);
      this.cache.set(ownerKey, rows[0][0] ?? null);
      return true;
    }

    // Everything except the anchor must be free.
    for (let i = 0; i < height; i++) {
      for (let j = 0; j < width; j++) {
        if (i === 0 && j === 0) continue;
        const rr = r + i;
        const cc = c + j;
        if (rr >= MAX_ROWS || cc >= MAX_COLS) return false;
        const occupant = sheet.cells.get(pack(rr, cc));
        if (occupant && (occupant.raw !== undefined || occupant.formula !== undefined)) return false;
        const owner = this.spillOwner.get(key(sheet.id, rr, cc));
        if (owner && owner !== ownerKey) return false;
      }
    }

    this.releaseSpill(ownerKey);
    this.cache.set(ownerKey, rows[0][0] ?? null);
    for (let i = 0; i < height; i++) {
      for (let j = 0; j < width; j++) {
        if (i === 0 && j === 0) continue;
        const k = key(sheet.id, r + i, c + j);
        this.spillOwner.set(k, ownerKey);
        this.cache.set(k, rows[i][j] ?? null);
      }
    }
    return true;
  }

  private releaseSpill(ownerKey: string): void {
    for (const [k, owner] of [...this.spillOwner]) {
      if (owner === ownerKey) {
        this.spillOwner.delete(k);
        this.cache.delete(k);
      }
    }
  }

  private resolverFor(
    sheet: Sheet,
    stack: Set<string>,
    onRef: NonNullable<Resolver['onRef']>,
  ): Resolver {
    return {
      sheet: (name) => (name === null ? sheet : this.sheetByName(name)),
      // Recurse so a formula never sees a stale (or missing) precedent.
      cellValue: (s, r, c) => this.compute(s, r, c, stack),
      named: (name) => {
        const up = name.toUpperCase();
        for (const [k, v] of Object.entries(this.wb.namedRanges)) {
          if (k.toUpperCase() === up) return v;
        }
        return null;
      },
      onRef,
    };
  }

  /** Read a rectangular block as plain values. */
  readBlock(sheet: Sheet, range: Range): Value[][] {
    const n = normRange(range);
    const out: Value[][] = [];
    for (let r = n.r1; r <= n.r2; r++) {
      const row: Value[] = [];
      for (let c = n.c1; c <= n.c2; c++) row.push(this.valueAt(sheet, r, c));
      out.push(row);
    }
    return out;
  }

  /* --------------------------------------------------------- undo / redo */

  begin(): void {
    this.batching += 1;
    if (this.batching === 1) this.undoStack.push(this.snapshot());
  }

  commit(): void {
    this.batching = Math.max(0, this.batching - 1);
    if (this.batching === 0) {
      this.redoStack = [];
      if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
    }
  }

  /** Group several mutations into a single undo step. */
  transact<T>(fn: () => T): T {
    this.begin();
    try {
      return fn();
    } finally {
      this.commit();
    }
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): boolean {
    const snap = this.undoStack.pop();
    if (!snap) return false;
    this.redoStack.push(this.snapshot());
    this.restore(snap);
    return true;
  }

  redo(): boolean {
    const snap = this.redoStack.pop();
    if (!snap) return false;
    this.undoStack.push(this.snapshot());
    this.restore(snap);
    return true;
  }

  private snapshot(): Snapshot {
    return {
      sheets: this.wb.sheets.map((s) => ({
        id: s.id,
        name: s.name,
        cells: new Map(s.cells),
        rows: s.rows,
        cols: s.cols,
        colWidths: new Map(s.colWidths),
        rowHeights: new Map(s.rowHeights),
        hiddenRows: [...s.hiddenRows],
        hiddenCols: [...s.hiddenCols],
        frozenRows: s.frozenRows,
        frozenCols: s.frozenCols,
        tabColor: s.tabColor,
        charts: s.charts.map((c) => ({ ...c, series: c.series.map((x) => ({ ...x })) })),
        condFormats: s.condFormats.map((f) => ({
          range: { ...f.range },
          rules: f.rules.map((r) => ({ ...r, values: [...r.values] })),
        })),
        defColWidth: s.defColWidth,
        defRowHeight: s.defRowHeight,
      })),
      activeSheet: this.wb.activeSheet,
      namedRanges: { ...this.wb.namedRanges },
    };
  }

  private restore(snap: Snapshot): void {
    this.wb.sheets = snap.sheets.map(
      (s): Sheet => ({
        id: s.id,
        name: s.name,
        cells: new Map(s.cells),
        rows: s.rows,
        cols: s.cols,
        colWidths: new Map(s.colWidths),
        rowHeights: new Map(s.rowHeights),
        hiddenRows: new Set(s.hiddenRows),
        hiddenCols: new Set(s.hiddenCols),
        frozenRows: s.frozenRows,
        frozenCols: s.frozenCols,
        tabColor: s.tabColor,
        charts: s.charts,
        condFormats: s.condFormats,
        defColWidth: s.defColWidth,
        defRowHeight: s.defRowHeight,
      }),
    );
    this.wb.activeSheet = snap.activeSheet;
    this.wb.namedRanges = { ...snap.namedRanges };
    this.invalidateAll();
  }

  /* ---------------------------------------------------------- misc helpers */

  /**
   * Drop every cached result and queue all formula cells for recalculation.
   * Called after anything that can move cells around.
   */
  invalidateAll(): void {
    this.dirtySet.clear();
    this.cache.clear();
    this.spillOwner.clear();
    this.dependents.clear();
    for (const sheet of this.wb.sheets) {
      for (const [k, cell] of sheet.cells) {
        if (cell.formula !== undefined) {
          this.dirtySet.add(key(sheet.id, unpackRow(k), unpackCol(k)));
        }
      }
    }
  }

  private touchExtent(sheet: Sheet, r: number, c: number): void {
    if (r + 1 > sheet.rows) sheet.rows = Math.min(MAX_ROWS, r + 1);
    if (c + 1 > sheet.cols) sheet.cols = Math.min(MAX_COLS, c + 1);
  }

  /** Shrink the used range when trailing cells are removed. */
  trimExtent(sheet: Sheet): void {
    let maxR = -1;
    let maxC = -1;
    for (const k of sheet.cells.keys()) {
      const r = unpackRow(k);
      const c = unpackCol(k);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
    sheet.rows = Math.max(1, maxR + 1);
    sheet.cols = Math.max(1, maxC + 1);
  }

  /** Trim the used range of every sheet. */
  trimAll(): void {
    for (const sheet of this.wb.sheets) this.trimExtent(sheet);
  }

  colWidth(sheet: Sheet, c: number): number {
    return sheet.colWidths.get(c) ?? sheet.defColWidth;
  }
  rowHeight(sheet: Sheet, r: number): number {
    return sheet.rowHeights.get(r) ?? sheet.defRowHeight;
  }

  onRevision(fn: () => void): () => void {
    this.revisionListeners.add(fn);
    return () => {
      this.revisionListeners.delete(fn);
    };
  }

  emit(): void {
    for (const fn of this.revisionListeners) fn();
  }

  usedRange(sheet: Sheet): Range {
    let maxR = 0;
    let maxC = 0;
    for (const k of sheet.cells.keys()) {
      const r = unpackRow(k);
      const c = unpackCol(k);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
    return { r1: 0, c1: 0, r2: maxR, c2: maxC };
  }
}

/* ------------------------------------------------------------------ module */

function key(sheetId: string, r: number, c: number): string {
  return `${sheetId}:${r}:${c}`;
}

/** Decode a `sheetId:r:c` cache key. */
function parseKey(k: string): { sheetId: string; r: number; c: number } | null {
  const sepC = k.lastIndexOf(':');
  const sepR = k.lastIndexOf(':', sepC - 1);
  if (sepR < 0) return null;
  const c = Number(k.slice(sepC + 1));
  const r = Number(k.slice(sepR + 1, sepC));
  if (Number.isNaN(r) || Number.isNaN(c)) return null;
  return { sheetId: k.slice(0, sepR), r, c };
}

function uniqueSheetName(wb: Workbook, wanted: string): string {
  const base = wanted.slice(0, 31);
  if (!wb.sheets.some((s) => s.name.toLowerCase() === base.toLowerCase())) return base;
  for (let i = 2; i < 1000; i++) {
    const candidate = `${base.slice(0, 28)} ${i}`.slice(0, 31);
    if (!wb.sheets.some((s) => s.name.toLowerCase() === candidate.toLowerCase())) return candidate;
  }
  return `${base} ${Date.now()}`;
}

/** Shift packed keys to make room for inserted/deleted rows or columns. */
function shiftCells(
  cells: Map<number, Cell>,
  atRow: number,
  atCol: number,
  count: number,
  dir: 1 | -1,
  n: number,
): Map<number, Cell> {
  const out = new Map<number, Cell>();
  for (const [k, cell] of cells) {
    let r = unpackRow(k);
    let c = unpackCol(k);
    if (dir === 1) {
      if (r >= atRow) r += n;
      if (c >= atCol) c += n;
    } else if (atCol === 0) {
      // row deletion
      if (r >= atRow && r < atRow + count) continue;
      if (r >= atRow) r -= n;
    } else {
      // column deletion
      if (c >= atCol && c < atCol + count) continue;
      if (c >= atCol) c -= n;
    }
    out.set(pack(r, c), cell);
  }
  return out;
}

function shiftMap(
  m: Map<number, number>,
  at: number,
  count: number,
  dir: 1 | -1,
): Map<number, number> {
  const out = new Map<number, number>();
  for (const [k, v] of m) {
    if (dir === 1) {
      out.set(k >= at ? k + count : k, v);
    } else {
      if (k >= at && k < at + count) continue;
      out.set(k >= at ? k - count : k, v);
    }
  }
  return out;
}

function shiftSet(set: Set<number>, at: number, count: number, dir: 1 | -1): void {
  const out = new Set<number>();
  for (const k of set) {
    if (dir === 1) out.add(k >= at ? k + count : k);
    else {
      if (k >= at && k < at + count) continue;
      out.add(k >= at ? k - count : k);
    }
  }
  set.clear();
  for (const v of out) set.add(v);
}

export { colName, rangeToA1 };
export type { Style };
