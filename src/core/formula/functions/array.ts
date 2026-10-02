/**
 * Dynamic-array functions.
 *
 * These return a 2D array; the recalculation engine spills the result into
 * neighbouring cells when they are empty, and reports #SPILL when they are not.
 */

import { err, isError, type CellError, type Value } from '../../types';
import { toNumber, toText } from '../../coerce';
import {
  isArray,
  isRange,
  makeArray,
  type ArrayRef,
  type FnCtx,
  type FnDef,
  type Lifted,
} from '../fntypes';
import { argFlat, argMatrix, firstError, numOrDefault } from '../helpers';

/** Stable sort so equal keys keep their input order. */
function stableSort<T>(items: T[], cmp: (a: T, b: T) => number): T[] {
  return items
    .map((v, i) => [v, i] as [T, number])
    .sort((a, b) => cmp(a[0], b[0]) || a[1] - b[1])
    .map((p) => p[0]);
}

/** Compare two values with Excel's ordering: number < text < logical. */
function compareValues(a: Value, b: Value): number {
  if (a === null && b === null) return 0;
  if (a === null) return typeof b === 'number' ? -1 : -1;
  if (b === null) return typeof a === 'number' ? 1 : 1;
  if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : a < b ? -1 : 1;
  if (typeof a === 'string' && typeof b === 'string') {
    const c = a.localeCompare(b, undefined, { sensitivity: 'base' });
    return c === 0 ? 0 : c < 0 ? -1 : 1;
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') {
    return a === b ? 0 : a ? 1 : -1;
  }
  if (typeof a === 'number') return -1;
  if (typeof b === 'number') return 1;
  if (typeof a === 'boolean') return 1;
  return -1;
}

function uniqueRows(rows: Value[][], byCol: boolean, exactlyOnce: boolean): Value[][] {
  const key = (row: Value[]) => (byCol ? row.join('') : row.join(''));
  const counts = new Map<string, number>();
  for (const row of rows) counts.set(key(row), (counts.get(key(row)) ?? 0) + 1);
  const seen = new Set<string>();
  const out: Value[][] = [];
  for (const row of rows) {
    const k = key(row);
    if (seen.has(k)) continue;
    seen.add(k);
    if (exactlyOnce && counts.get(k) !== 1) continue;
    out.push(row);
  }
  return out;
}

function matchModePredicate(mode: number): (needle: Value, hay: Value) => boolean {
  if (mode === 2) {
    const re = (s: string) =>
      new RegExp(
        '^' +
          s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') +
          '$',
        'i',
      );
    return (n, h) => typeof h === 'string' && re(toText(n)).test(h);
  }
  const loose = (n: Value, h: Value): number => {
    if (typeof n === 'number' && typeof h === 'number') return n === h ? 0 : n < h ? -1 : 1;
    if (typeof n === 'string' && typeof h === 'string') {
      return n.localeCompare(h, undefined, { sensitivity: 'base' });
    }
    if (n === null && h === null) return 0;
    return -1;
  };
  if (mode === -1) return (n, h) => loose(n, h) >= 0;
  if (mode === 1) return (n, h) => loose(n, h) <= 0;
  return (n, h) => loose(n, h) === 0;
}

export const arrayFns: FnDef[] = [
  {
    name: 'SEQUENCE', min: 1, max: 4,
    fn: (a, ctx): Value | ArrayRef => {
      const rows = intArg(a, ctx, 0, 1);
      if (isError(rows)) return rows;
      const cols = intArg(a, ctx, 1, 1);
      if (isError(cols)) return cols;
      const start = numOr(a, ctx, 2, 1);
      if (isError(start)) return start;
      const step = numOr(a, ctx, 3, 1);
      if (isError(step)) return step;
      if (rows < 1 || cols < 1 || rows * cols > 1_000_000) return err('NUM', 'SEQUENCE size out of range');
      const out: Value[][] = [];
      for (let r = 0; r < rows; r++) {
        const row: Value[] = [];
        for (let c = 0; c < cols; c++) row.push(start + (r * cols + c) * step);
        out.push(row);
      }
      return makeArray(out);
    },
  },

  {
    name: 'UNIQUE', min: 1, max: 3,
    fn: (a, ctx): Value | ArrayRef => {
      const m = argMatrix(a[0], ctx);
      const byCol = truthyOpt(a[1], false);
      const exactlyOnce = truthyOpt(a[2], false);
      const rows = byCol
        ? (m.rows[0]?.length ?? 0) > 0 && (m.rows.length === 0 || m.rows[0].length !== m.rows.length)
          ? transposeMat(m).rows
          : m.rows
        : m.rows;
      const uniq = uniqueRows(rows.filter((r) => r.some((v) => v !== null)), byCol, exactlyOnce);
      return makeArray(uniq.length ? uniq : [[null]]);
    },
  },

  {
    name: 'SORT', min: 1, max: 4,
    fn: (a, ctx): Value | ArrayRef => {
      const e = firstError(a, ctx);
      if (e) return e;
      const m = argMatrix(a[0], ctx);
      let rows = m.rows.map((r) => r.slice());
      if (rows.length === 1 && (rows[0]?.length ?? 0) > 1) rows = rows[0].map((v) => [v]);

      const index = intArg(a, ctx, 1, 1);
      if (isError(index)) return index;
      const order = numOr(a, ctx, 2, 1);
      if (isError(order)) return order;
      const byCol = truthyOpt(a[3], false);
      const desc = order < 0;

      if (byCol) {
        const width = rows.length;
        let cols = rows.map((_, c) => rows.map((row) => row[c]));
        const keyCol = index - 1;
        cols = stableSort(cols, (x, y) => compareValues(x[keyCol], y[keyCol]) * (desc ? -1 : 1));
        const out: Value[][] = [];
        for (let c = 0; c < width; c++) out.push(cols.map((col) => col[c]));
        return makeArray(out);
      }

      const keyCol = index - 1;
      rows = stableSort(rows, (x, y) => compareValues(x[keyCol], y[keyCol]) * (desc ? -1 : 1));
      return makeArray(rows.length ? rows : [[null]]);
    },
  },

  {
    name: 'SORTBY', min: 2, max: Infinity,
    fn: (a, ctx): Value | ArrayRef => {
      const e = firstError(a, ctx);
      if (e) return e;
      const m = argMatrix(a[0], ctx);
      const rows = m.rows.map((r) => r.slice());

      // Remaining arguments are (by_array, [sort_order]) pairs.
      const keys: { values: Value[]; desc: boolean }[] = [];
      for (let i = 1; i < a.length; ) {
        const values = argFlat(a[i], ctx);
        let order = 1;
        if (i + 1 < a.length) {
          const o = toNumber(argFlat(a[i + 1], ctx)[0] ?? 1);
          if (typeof o === 'number' && !isError(o)) order = o;
          i += 2;
        } else {
          i += 1;
        }
        keys.push({ values, desc: order < 0 });
      }

      // Rank once per key, then sort row indices by their rank tuples.
      let idx = rows.map((_, i) => i);
      for (const key of keys) {
        const byValue = stableSort(
          idx,
          (x, y) => compareValues(key.values[x] ?? null, key.values[y] ?? null),
        );
        const rank = new Map<number, number>();
        byValue.forEach((rowIdx, position) => rank.set(rowIdx, position));
        const dir = key.desc ? -1 : 1;
        idx = stableSort(idx, (x, y) => dir * ((rank.get(x) ?? 0) - (rank.get(y) ?? 0)));
      }
      return makeArray(idx.map((i) => rows[i]));
    },
  },

  {
    name: 'FILTER', min: 2, max: 3,
    fn: (a, ctx): Value | ArrayRef => {
      const e = firstError(a, ctx);
      if (e) return e;
      const m = argMatrix(a[0], ctx);
      const include = argFlat(a[1], ctx);
      if (include.length !== m.rows.length) {
        return err('VALUE', 'FILTER include range must be one column tall');
      }
      if (a.length > 2) {
        if (isError(a[2] as Value)) return a[2] as Value;
        return makeArray([[a[2] as Value]]);
      }
      const out = m.rows.filter((_, i) => {
        const v = include[i];
        return typeof v === 'boolean' ? v : typeof v === 'number' && v !== 0;
      });
      if (!out.length) return err('CALC', 'FILTER returned no rows');
      return makeArray(out);
    },
  },

  {
    name: 'XMATCH', min: 2, max: 4,
    fn: (a, ctx): Value | ArrayRef => {
      const needle = argFlat(a[0], ctx)[0] ?? null;
      const hay = argFlat(a[1], ctx);
      const mode = intArg(a, ctx, 2, 0);
      const search = intArg(a, ctx, 3, 1);
      if (isError(mode)) return mode;
      if (isError(search)) return search;
      const pred = matchModePredicate(mode);
      const indices = hay.map((_, i) => i);
      const ordered = search < 0 ? [...indices].reverse() : indices;
      for (const i of ordered) {
        if (pred(needle, hay[i])) return i + 1;
      }
      return err('N/A', 'XMATCH found no match');
    },
  },

  { name: 'TAKE', min: 2, max: 3, fn: (a, ctx) => takeDrop(a, ctx, true) },
  { name: 'DROP', min: 2, max: 3, fn: (a, ctx) => takeDrop(a, ctx, false) },

  {
    name: 'VSTACK', min: 1, fn: (a, ctx): Value | ArrayRef => {
      const mats = a.map((x) => argMatrix(x, ctx));
      const width = Math.max(...mats.map((m) => Math.max(0, ...m.rows.map((r) => r.length))));
      const rows: Value[][] = [];
      for (const m of mats) {
        for (const r of m.rows) {
          const copy = r.slice();
          while (copy.length < width) copy.push(null);
          rows.push(copy);
        }
      }
      return makeArray(rows.length ? rows : [[null]]);
    },
  },

  {
    name: 'HSTACK', min: 1, fn: (a, ctx): Value | ArrayRef => {
      const mats = a.map((x) => argMatrix(x, ctx));
      const height = Math.max(...mats.map((m) => m.rows.length));
      const rows: Value[][] = [];
      for (let r = 0; r < height; r++) {
        const row: Value[] = [];
        for (const m of mats) {
          const src = m.rows[r] ?? [];
          row.push(...src);
        }
        rows.push(row);
      }
      return makeArray(rows.length ? rows : [[null]]);
    },
  },

  {
    name: 'TOROW', min: 1, max: 3, fn: (a, ctx): Value | ArrayRef => reduceTo(a, ctx, true) },
  {
    name: 'TOCOL', min: 1, max: 3, fn: (a, ctx): Value | ArrayRef => reduceTo(a, ctx, false) },
  {
    name: 'WRAPROWS', min: 2, max: 3, fn: (a, ctx): Value | ArrayRef => {
      const vals = argFlat(a[0], ctx);
      const width = intArg(a, ctx, 1, 1);
      if (isError(width)) return width;
      if (width < 1) return err('NUM');
      const pad = truthyOpt(a[2], false);
      const rows: Value[][] = [];
      for (let i = 0; i < vals.length; i += width) {
        const chunk = vals.slice(i, i + width);
        while (pad && chunk.length < width) chunk.push(null);
        rows.push(chunk);
      }
      return makeArray(rows.length ? rows : [[null]]);
    },
  },

  { name: 'CHOOSEROWS', min: 2, fn: (a, ctx): Value | ArrayRef => chooseDim(a, ctx, true) },
  { name: 'CHOOSECOLS', min: 2, fn: (a, ctx): Value | ArrayRef => chooseDim(a, ctx, false) },
];

/* ---------------------------------------------------------------- helpers */

function numOr(args: Lifted[], ctx: FnCtx, i: number, fallback: number): number | CellError {
  return numOrDefault(args, ctx, i, fallback);
}

/** Integer-valued optional argument, checked for errors before truncating. */
function intArg(args: Lifted[], ctx: FnCtx, i: number, fallback: number): number | CellError {
  const v = numOrDefault(args, ctx, i, fallback);
  return isError(v) ? v : Math.trunc(v);
}

function truthyOpt(arg: Lifted | undefined, fallback: boolean): boolean {
  if (arg === undefined) return fallback;
  const v = isRange(arg) || isArray(arg) ? null : (arg as Value);
  if (v === null) return fallback;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const up = toText(v).toUpperCase();
  if (up === 'TRUE') return true;
  if (up === 'FALSE') return false;
  return fallback;
}

function transposeMat(m: ArrayRef): ArrayRef {
  const width = Math.max(0, ...m.rows.map((r) => r.length));
  const out: Value[][] = [];
  for (let c = 0; c < width; c++) out.push(m.rows.map((r) => (c < r.length ? r[c] : null)));
  return makeArray(out.length ? out : [[]]);
}

function takeDrop(args: Lifted[], ctx: FnCtx, take: boolean): Value | ArrayRef {
  const e = firstError(args, ctx);
  if (e) return e;
  const m = argMatrix(args[0], ctx);
  const nRows = intArg(args, ctx, 1, 0);
  const nCols = intArg(args, ctx, 2, 0);
  if (isError(nRows)) return nRows;
  if (isError(nCols)) return nCols;

  let rows = m.rows.map((r) => r.slice());
  if (rows.length === 1 && (rows[0]?.length ?? 0) > 1) rows = rows[0].map((v) => [v]);

  if (nRows !== 0) {
    rows = take ? rows.slice(0, nRows) : (nRows > 0 ? rows.slice(nRows) : rows.slice(0));
    if (!take && nRows < 0) rows = rows.slice(0, nRows);
  }
  if (nCols !== 0) {
    rows = rows.map((r) => {
      const slice = take ? r.slice(0, nCols) : r.slice(nCols);
      if (!take && nCols < 0) return r.slice(0, nCols);
      return slice;
    });
  }
  if (!rows.length) return err('CALC', 'Result is empty');
  return makeArray(rows);
}

function reduceTo(args: Lifted[], ctx: FnCtx, byRow: boolean): Value | ArrayRef {
  const e = firstError(args, ctx);
  if (e) return e;
  const m = argMatrix(args[0], ctx);
  const ignore = args[1] === undefined ? true : truthyOpt(args[1], true);
  const scanByRow = args[2] === undefined ? !byRow : truthyOpt(args[2], !byRow);
  let vals = scanByRow ? m.rows.flat() : m.rows.map((r) => r[0] ?? null);
  if (ignore) vals = vals.filter((v) => v !== null && v !== '');
  if (!vals.length) return makeArray([[null]]);
  return makeArray([vals]);
}

function chooseDim(args: Lifted[], ctx: FnCtx, rows: boolean): Value | ArrayRef {
  const e = firstError(args, ctx);
  if (e) return e;
  const m = argMatrix(args[0], ctx);
  const raw = argFlat(args[1], ctx);
  const idx: number[] = [];
  for (const v of raw) {
    const n = toNumber(v);
    if (isError(n)) return err('VALUE', 'Indices must be numbers');
    idx.push(Math.trunc(n));
  }
  if (!idx.length || idx.some((i) => i < 1)) return err('VALUE', 'Indices start at 1');

  if (rows) {
    return makeArray(idx.map((i) => (m.rows[i - 1] ?? []).slice()));
  }
  const out: Value[][] = [];
  for (let r = 0; r < m.rows.length; r++) {
    out.push(idx.map((i) => m.rows[r][i - 1] ?? null));
  }
  return makeArray(out.length ? out : [[null]]);
}
