/**
 * SheetCraft core type definitions.
 *
 * Everything in the app is built on these types. Values follow Excel semantics
 * closely enough that a .xlsx round-trip preserves meaning.
 */

/* ------------------------------------------------------------------ values */

export type ErrorKind =
  | 'NULL'
  | 'DIV/0'
  | 'VALUE'
  | 'REF'
  | 'NAME'
  | 'NUM'
  | 'N/A'
  | 'CYCLE'
  | 'CALC'
  | 'PARSE'
  | 'SPILL'
  | 'GETTING_DATA';

export interface CellError {
  readonly __error: ErrorKind;
  readonly message?: string;
}

/** A value that can live in a cell or be produced by the formula engine. */
export type Value = number | string | boolean | null | CellError;

export const isError = (v: unknown): v is CellError =>
  typeof v === 'object' && v !== null && '__error' in (v as object);

export const err = (kind: ErrorKind, message?: string): CellError =>
  message === undefined ? { __error: kind } : { __error: kind, message };

export const isBlank = (v: Value): boolean =>
  v === null || v === '' || (isError(v) === false && false);

/* ----------------------------------------------------------------- formats */

export type AlignH = 'left' | 'center' | 'right';

export interface CellFormat {
  bold?: 1;
  italic?: 1;
  underline?: 1;
  strike?: 1;
  /** font colour, css hex */
  color?: string;
  /** fill colour, css hex */
  bg?: string;
  align?: AlignH;
  valign?: 'top' | 'middle' | 'bottom';
  wrap?: 1;
  /** Excel number-format code, e.g. "#,##0.00" or "0%" */
  numFmt?: string;
  font?: string;
  size?: number;
  border?: {
    top?: string;
    right?: string;
    bottom?: string;
    left?: string;
  };
}

export interface Style {
  id: number;
  fmt: CellFormat;
}

/* ------------------------------------------------------------------- cells */

/**
 * Stored cell. `raw` is the literal the user typed; `formula` is present when
 * the cell is calculated and never stored together with `raw`.
 */
export interface Cell {
  /** literal input; strings are kept verbatim so "007" survives */
  raw?: number | string | boolean | null;
  /** formula body WITHOUT the leading '=' */
  formula?: string;
  styleId?: number;
  comment?: string;
  /** cached last computed value, avoids recalculating the whole book on paint */
  cached?: Value;
  /** dirty flag consumed by the recalc engine */
  dirty?: 1;
}

/* ------------------------------------------------------------------ charts */

export type ChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'radar';

export interface ChartSeries {
  name?: string;
  /** A1-style range reference, e.g. `Sheet1!$B$1:$B$10` */
  values: string;
  /** category (x) reference for line/scatter/area */
  categories?: string;
  color?: string;
}

export interface ChartSpec {
  id: string;
  name: string;
  type: ChartType;
  title: string;
  sheetId: string;
  /** anchor cell (top-left of the chart frame) + size in grid units */
  anchor: { r: number; c: number };
  width: number;
  height: number;
  series: ChartSeries[];
  showLegend: boolean;
  showTitle: boolean;
  showGridlines: boolean;
  showDataLabels: boolean;
  stacked?: boolean;
  smooth?: boolean;
  /** pie/doughnut only: treat first series as single slice set */
  holeSize?: number;
}

/* ------------------------------------------------------- conditional format */

export type CondOperator =
  | 'gt'
  | 'gte'
  | 'lt'
  | 'lte'
  | 'eq'
  | 'neq'
  | 'between'
  | 'contains'
  | 'notContains'
  | 'dup'
  | 'empty'
  | 'top10';

export interface ConditionalRule {
  id: string;
  op: CondOperator;
  /** numeric thresholds; 1 or 2 entries depending on operator */
  values: (number | string)[];
  styleId: number;
}

/* ------------------------------------------------------------------- sheet */

export interface Sheet {
  id: string;
  name: string;
  /** sparse cell store keyed by pack(r,c) */
  cells: Map<number, Cell>;
  /** highest used row + 1 */
  rows: number;
  /** highest used col + 1 */
  cols: number;
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  frozenRows: number;
  frozenCols: number;
  tabColor?: string;
  charts: ChartSpec[];
  condFormats: { range: Range; rules: ConditionalRule[] }[];
  /** default width in px */
  defColWidth: number;
  defRowHeight: number;
}

/* ----------------------------------------------------------------- workbook */

export interface Workbook {
  id: string;
  name: string;
  sheets: Sheet[];
  activeSheet: number;
  /** name -> fully qualified reference, e.g. `Rates!$A$1:$B$9` */
  namedRanges: Record<string, string>;
  styles: Style[];
  /** recalculation counter, bumped on every structural change */
  revision: number;
  created: number;
  modified: number;
  /** undo stack is kept outside the workbook so saves stay small */
}

/* -------------------------------------------------------------- selection */

export interface Pos {
  r: number;
  c: number;
}

export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/* -------------------------------------------------------------- clipboard */

export interface ClipCell {
  v: string;
  f?: string;
  styleId?: number;
}

/* ------------------------------------------------------------------- misc */

export interface DirtyRegion {
  sheetId: string;
  range: Range;
}

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;
export const DEFAULT_COL_W = 96;
export const DEFAULT_ROW_H = 22;
export const HEADER_W = 52;
export const HEADER_H = 26;
