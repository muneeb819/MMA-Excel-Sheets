/**
 * Shared types for the built-in function library.
 *
 * Functions receive arguments already "lifted": a cell reference arrives as a
 * `RangeRef` so functions like COUNT can see the whole block, while plain
 * expressions arrive as scalars.
 */

import { err, type CellError, type Sheet, type Value } from '../types';
import type { Node } from './parser';

export interface RangeRef {
  readonly __range: true;
  readonly sheet: Sheet | null;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export interface ArrayRef {
  readonly __array: true;
  /** rows[r][c] — may contain errors */
  rows: Value[][];
}

export type Lifted = Value | RangeRef | ArrayRef;

export const isRange = (v: unknown): v is RangeRef =>
  typeof v === 'object' && v !== null && '__range' in (v as object);

export const isArray = (v: unknown): v is ArrayRef =>
  typeof v === 'object' && v !== null && '__array' in (v as object);

export function makeArray(rows: Value[][]): ArrayRef {
  return { __array: true, rows };
}

/** Context handed to every function call. */
export interface FnCtx {
  /** sheet the formula lives on */
  readonly sheet: Sheet;
  /** flatten a reference into row-major values */
  values(ref: RangeRef): Value[];
  rowsOf(ref: RangeRef): Value[][];
  shape(ref: RangeRef): { rows: number; cols: number };
  /** read a single cell from a resolved sheet (already scoped to `ref.sheet`) */
  cell(sheet: Sheet | null, r: number, c: number): Value;
  /** true when the referenced sheet exists */
  valid(sheet: Sheet | null): boolean;
  /** args as written, for lazy functions */
  evalNode(n: Node): Lifted;
  /** function name for implicit-intersection hints in error messages */
  fnName: string;
}

export interface FnDef {
  name: string;
  min?: number;
  /**
   * Maximum argument count. When omitted the function is treated as variadic
   * (no upper bound), which is what the aggregate family needs.
   */
  max?: number;
  /** re-evaluated on every recalculation pass (RAND, TODAY, OFFSET...) */
  volatile?: boolean;
  /** when set, `lazyFn` is used instead of `fn` */
  lazy?: boolean;
  /** eager implementation: arguments arrive already lifted */
  fn?: (args: Lifted[], ctx: FnCtx) => Value | ArrayRef;
  /** lazy implementation: arguments arrive as unevaluated AST nodes */
  lazyFn?: (args: Node[], ctx: FnCtx) => Value | ArrayRef;
}

/** Lower/upper arity bounds for a definition. */
export function arity(def: FnDef): { min: number; max: number } {
  const min = def.min ?? 0;
  return { min, max: def.max ?? Infinity };
}

export type FnRegistry = Map<string, FnDef>;

/* ---------------------------------------------------------- arg utilities */

/** Row-major flatten of any lifted argument. */
export function flat(arg: Lifted, ctx: FnCtx): Value[] {
  if (isRange(arg)) return ctx.values(arg);
  if (isArray(arg)) return arg.rows.flat();
  return [arg as Value];
}

/** All values across every argument. */
export function flatAll(args: Lifted[], ctx: FnCtx): Value[] {
  const out: Value[] = [];
  for (const a of args) out.push(...flat(a, ctx));
  return out;
}

/** The effective scalar of an argument: top-left cell for ranges. */
export function scalar(arg: Lifted, ctx: FnCtx): Value {
  if (isRange(arg)) {
    const vs = ctx.values(arg);
    return vs.length ? vs[0] : null;
  }
  if (isArray(arg)) {
    const vs = arg.rows.flat();
    return vs.length ? vs[0] : null;
  }
  return arg as Value;
}

export function argValues(args: Lifted[], ctx: FnCtx): Value[] {
  return args.map((a) => scalar(a, ctx));
}

/** Numeric values only, skipping text, booleans and blanks (AVERAGE, MAX...). */
export function numbersOnly(args: Lifted[], ctx: FnCtx): number[] {
  const out: number[] = [];
  for (const a of args) {
    for (const v of flat(a, ctx)) {
      if (typeof v === 'number') out.push(v);
    }
  }
  return out;
}

/** Numbers, booleans and date-serial numbers; skips text and blanks. */
export function numbersAndBools(args: Lifted[], ctx: FnCtx): number[] {
  const out: number[] = [];
  for (const a of args) {
    for (const v of flat(a, ctx)) {
      if (typeof v === 'number') out.push(v);
      else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    }
  }
  return out;
}

/** Counts non-empty cells. */
export function countValues(args: Lifted[], ctx: FnCtx): number {
  let n = 0;
  for (const a of args) for (const v of flat(a, ctx)) if (v !== null) n++;
  return n;
}

/* ------------------------------------------------------------- arity check */

export function checkArity(def: FnDef, count: number): CellError | null {
  const min = def.min ?? 0;
  const max = def.max ?? min;
  if (count < min || count > max) {
    return err(
      'VALUE',
      `${def.name} expects ${min === max ? min : `${min}-${max}`} argument(s), got ${count}`,
    );
  }
  return null;
}
