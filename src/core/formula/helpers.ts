/**
 * Shared helpers for the function library.
 *
 * Every function receives already-lifted arguments (scalars, `RangeRef` or
 * `ArrayRef`), so these helpers are the only place that needs to care about the
 * difference.
 */

import { err, isError, type CellError, type Value } from '../types';
import { toNumber, toText } from '../coerce';
import {
  flat,
  isArray,
  isRange,
  numbersOnly,
  type ArrayRef,
  type FnCtx,
  type Lifted,
} from './fntypes';

/* ------------------------------------------------------------ error helpers */

/** First error found anywhere in the arguments, or null. */
export function firstError(args: Lifted[], ctx: FnCtx): Value | null {
  for (const a of args) {
    if (isError(a)) return a;
    if (isArray(a)) {
      for (const row of a.rows) for (const v of row) if (isError(v)) return v;
    } else if (isRange(a)) {
      for (const v of ctx.values(a)) if (isError(v)) return v;
    }
  }
  return null;
}

/* ------------------------------------------------------------------ typing */

/** Every value across every argument, errors included. */
export function allValues(args: Lifted[], ctx: FnCtx): Value[] {
  const out: Value[] = [];
  for (const a of args) out.push(...flat(a, ctx));
  return out;
}

/** Only string values (used by text functions). */
export function textValues(args: Lifted[], ctx: FnCtx): string[] {
  const out: string[] = [];
  for (const a of args) {
    for (const v of flat(a, ctx)) if (typeof v === 'string') out.push(v);
  }
  return out;
}

/** Text form of the first argument, errors propagated. */
export function firstText(args: Lifted[], ctx: FnCtx): string | CellError {
  const v = firstScalar(args, ctx);
  if (isError(v)) return v;
  return toText(v);
}

/** Top-left scalar of the first argument. */
export function firstScalar(args: Lifted[], ctx: FnCtx): Value {
  const a = args[0];
  if (a === undefined) return null;
  if (isRange(a)) {
    const vs = ctx.values(a);
    return vs.length ? vs[0] : null;
  }
  if (isArray(a)) {
    const vs = a.rows.flat();
    return vs.length ? vs[0] : null;
  }
  return a as Value;
}

/** All scalars (each argument collapsed to its top-left value). */
export function scalars(args: Lifted[], ctx: FnCtx): Value[] {
  return args.map((a) => firstScalar([a], ctx));
}

/** Number of the first argument, or an error value. */
export function firstNumber(args: Lifted[], ctx: FnCtx): number | CellError {
  const v = firstScalar(args, ctx);
  if (isError(v)) return v;
  return toNumber(v);
}

export function numberAt(args: Lifted[], ctx: FnCtx, i: number): number | CellError {
  if (i < 0 || i >= args.length) return err('VALUE', 'Missing argument');
  const v = firstScalar([args[i]], ctx);
  if (isError(v)) return v;
  return toNumber(v);
}

/** Optional numeric argument with a fallback, e.g. ROUND's digits. */
export function optNumber(
  args: Lifted[],
  ctx: FnCtx,
  i: number,
  fallback: number,
): number | CellError {
  if (args.length <= i) return fallback;
  const v = firstScalar([args[i]], ctx);
  if (v === null || v === '') return fallback;
  if (isError(v)) return v;
  const n = toNumber(v);
  return n;
}

export function optText(args: Lifted[], ctx: FnCtx, i: number, fallback: string): string {
  if (args.length <= i) return fallback;
  const v = firstScalar([args[i]], ctx);
  if (v === null) return fallback;
  return toText(v);
}

/**
 * Optional numeric argument narrowed to `number | CellError` so the result can
 * feed arithmetic after a single isError check.
 */
export function numOrDefault(
  args: Lifted[],
  ctx: FnCtx,
  i: number,
  fallback: number,
): number | CellError {
  return optNumber(args, ctx, i, fallback);
}

/* ------------------------------------------------------------------ arrays */

/** Build a 2D array from a flat list, laid out column-major by `cols`. */
export function toMatrix(values: Value[], cols: number): ArrayRef {
  const rows: Value[][] = [];
  const n = Math.max(1, cols);
  for (let i = 0; i < values.length; i += n) rows.push(values.slice(i, i + n));
  if (rows.length === 0) rows.push([]);
  return { __array: true, rows };
}

/** Transpose a matrix. */
export function transpose(m: ArrayRef): ArrayRef {
  const width = Math.max(0, ...m.rows.map((r) => r.length));
  const out: Value[][] = [];
  for (let c = 0; c < width; c++) {
    out.push(m.rows.map((r) => (c < r.length ? r[c] : null)));
  }
  return { __array: true, rows: out.length ? out : [[]] };
}

/** Matrix of the first argument (range -> values, array -> as-is). */
export function argMatrix(arg: Lifted | undefined, ctx: FnCtx): ArrayRef {
  if (!arg) return { __array: true, rows: [[]] };
  if (isArray(arg)) return arg;
  if (isRange(arg)) return { __array: true, rows: ctx.rowsOf(arg) };
  return { __array: true, rows: [[arg as Value]] };
}

/** Flat row-major values of the first argument. */
export function argFlat(arg: Lifted | undefined, ctx: FnCtx): Value[] {
  if (!arg) return [];
  if (isArray(arg)) return arg.rows.flat();
  if (isRange(arg)) return ctx.values(arg);
  return [arg as Value];
}

/* ------------------------------------------------------------- misc utils */

/** Convert a 1-based Excel index into a zero-based one, clamped. */
export function zeroIndex(v: number): number {
  return Math.max(0, Math.floor(v) - 1);
}

export function notFound(): CellError {
  return err('N/A', 'Value not found');
}

export function numOut(v: number): number | CellError {
  return Number.isFinite(v) ? v : err('NUM', `Result is not a finite number (${v})`);
}

/** Divide safely, producing #DIV/0! like Excel. */
export function div(a: number, b: number): number {
  if (b === 0) return NaN;
  return a / b;
}
