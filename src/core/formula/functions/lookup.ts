/**
 * Lookup / reference functions and the database (DSUM-style) family.
 *
 * Ranges are rectangular, so table-shaped arguments are treated as 2D
 * matrices and indexed as `matrix[row][col]`.
 */

import { err, isError, type CellError, type Value } from '../../types';
import { toNumber, toText } from '../../coerce';
import { colName, parseRangeExpr } from '../../ref';
import {
  isArray,
  isRange,
  makeArray,
  type ArrayRef,
  type FnCtx,
  type FnDef,
  type Lifted,
  type RangeRef,
} from '../fntypes';
import { argFlat, argMatrix, firstError, transpose } from '../helpers';

/* ---------------------------------------------------------------- helpers */

type Matrix = Value[][];

/** Safe matrix read. */
function at(m: Matrix, r: number, c: number): Value {
  const row = m[r];
  if (!row) return null;
  return c < row.length ? row[c] : null;
}

function width(m: Matrix): number {
  return m.reduce((w, r) => Math.max(w, r.length), 0);
}

/** Top-left scalar of an argument. */
function argValue(arg: Lifted | undefined, ctx: FnCtx): Value {
  if (arg === undefined) return null;
  if (isRange(arg)) {
    const vals = ctx.values(arg);
    return vals.length ? vals[0] : null;
  }
  if (isArray(arg)) {
    const vals = (arg as ArrayRef).rows.flat();
    return vals.length ? vals[0] : null;
  }
  return arg as Value;
}

/** Row-major flatten of an argument. */
function argRow(arg: Lifted | undefined, ctx: FnCtx): Value[] {
  if (arg === undefined) return [];
  if (isRange(arg)) return ctx.values(arg);
  if (isArray(arg)) return (arg as ArrayRef).rows.flat();
  return [arg as Value];
}

function asMatrix(arg: Lifted | undefined, ctx: FnCtx): Matrix {
  if (arg === undefined) return [];
  if (isRange(arg)) return ctx.rowsOf(arg);
  if (isArray(arg)) return (arg as ArrayRef).rows;
  return [[arg as Value]];
}

/**
 * Excel equality: numbers compare numerically, text case-insensitively, and
 * blanks only match other blanks.
 */
function eqValue(a: Value, b: Value): boolean {
  if (a === null && b === null) return true;
  if (a === null || b === null) return false;
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase() === b.toLowerCase();
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b;
  return false;
}

/** Excel's ordering across types: number < text < logical. */
function cmpOrdered(a: Value, b: Value): number {
  const rank = (v: Value): number => {
    if (v === null) return -1;
    if (typeof v === 'number') return 0;
    if (typeof v === 'string') return 1;
    return 2;
  };
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : a < b ? -1 : 1;
  if (typeof a === 'string' && typeof b === 'string') {
    const c = a.localeCompare(b, undefined, { sensitivity: 'base' });
    return c === 0 ? 0 : c < 0 ? -1 : 1;
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') {
    return a === b ? 0 : a ? 1 : -1;
  }
  return 0;
}

function wildcardRe(pattern: string): RegExp {
  let out = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~') {
      const nxt = pattern[i + 1];
      if (nxt === '*' || nxt === '?' || nxt === '~') {
        out += nxt.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        i++;
        continue;
      }
    }
    if (ch === '*') out += '[\\s\\S]*';
    else if (ch === '?') out += '[\\s\\S]';
    else out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${out}$`, 'i');
}

function matchesWildcard(pattern: Value, candidate: Value): boolean {
  if (typeof pattern !== 'string' || typeof candidate !== 'string') return false;
  return wildcardRe(pattern).test(candidate);
}

/**
 * Find the position of `needle` in a sorted vector using Excel's approximate
 * rules: the largest value <= needle.
 *
 * Excel requires ascending order for approximate matches and leaves the result
 * undefined otherwise, so an unsorted vector falls back to a linear scan
 * rather than producing a meaningless binary-search answer.
 */
function searchKeys(keys: Value[], target: Value, approximate: boolean): number {
  const nonBlank = keys.filter((k) => k !== null);
  let sorted = true;
  for (let i = 1; i < nonBlank.length; i++) {
    if (cmpOrdered(nonBlank[i - 1], nonBlank[i]) > 0) {
      sorted = false;
      break;
    }
  }

  if (approximate && sorted) {
    let lo = 0;
    let hi = nonBlank.length - 1;
    let best = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const v = nonBlank[mid];
      if (cmpOrdered(v, target) <= 0) {
        best = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    if (best >= 0) return indexOfKey(keys, nonBlank[best]);
    return -1;
  }

  if (approximate) {
    // Unsorted input: take the last value that is not greater than the target.
    let best = -1;
    for (let i = 0; i < nonBlank.length; i++) {
      if (cmpOrdered(nonBlank[i], target) <= 0) best = i;
    }
    return best >= 0 ? indexOfKey(keys, nonBlank[best]) : -1;
  }

  for (let i = 0; i < keys.length; i++) {
    const v = keys[i];
    if (v === null) continue;
    if (eqValue(v, target)) return i;
  }
  return -1;
}

function indexOfKey(keys: Value[], wanted: Value): number {
  for (let i = 0; i < keys.length; i++) if (keys[i] === wanted) return i;
  return -1;
}

function rangeOf(arg: Lifted): RangeRef | null {
  return isRange(arg) ? arg : null;
}

function errNotFound(what: string): Value {
  return err('N/A', `${what} found no match`);
}

/* ----------------------------------------------------------------- lookups */

export const lookupFns: FnDef[] = [
  {
    name: 'VLOOKUP', min: 3, max: 4,
    fn: (a, ctx) => {
      const e = firstError(a, ctx);
      if (e) return e;
      const target = argValue(a[0], ctx);
      const table = asMatrix(a[1], ctx);
      if (!table.length) return err('N/A', 'VLOOKUP table is empty');
      const colRaw = toNumber(argValue(a[2], ctx));
      if (isError(colRaw)) return colRaw;
      const col = Math.trunc(colRaw);
      if (col < 1 || col > width(table)) return err('VALUE', `VLOOKUP column ${colRaw} is out of range`);
      const approx = truthy4(a[3], ctx, true);

      if (approx) {
        const keys = table.map((_, i) => at(table, i, 0));
        const row = searchKeys(keys, target, true);
        if (row < 0) return errNotFound('VLOOKUP');
        return at(table, row, col - 1);
      }

      const pattern = typeof target === 'string' && /[*?]/.test(target) ? target : null;
      for (let i = 0; i < table.length; i++) {
        const key = at(table, i, 0);
        const hit = pattern ? matchesWildcard(pattern, key) : eqValue(key, target);
        if (hit) return at(table, i, col - 1);
      }
      return errNotFound('VLOOKUP');
    },
  },

  {
    name: 'HLOOKUP', min: 3, max: 4,
    fn: (a, ctx) => {
      const e = firstError(a, ctx);
      if (e) return e;
      const target = argValue(a[0], ctx);
      const table = asMatrix(a[1], ctx);
      const cols = width(table);
      if (!table.length || !cols) return err('N/A', 'HLOOKUP table is empty');
      const rowRaw = toNumber(argValue(a[2], ctx));
      if (isError(rowRaw)) return rowRaw;
      const row = Math.trunc(rowRaw);
      if (row < 1 || row > table.length) return err('VALUE', `HLOOKUP row ${rowRaw} is out of range`);
      const approx = truthy4(a[3], ctx, true);

      const keyOf = (c: number): Value[] => table.map((_, i) => at(table, i, c));
      if (approx) {
        const keys = keyOf(0);
        const col = searchKeys(keys, target, true);
        if (col < 0) return errNotFound('HLOOKUP');
        return at(table, row - 1, col);
      }
      for (let c = 0; c < cols; c++) {
        if (eqValue(keyOf(c)[0], target)) return at(table, row - 1, c);
      }
      return errNotFound('HLOOKUP');
    },
  },

  {
    name: 'XLOOKUP', min: 3, max: 6,
    fn: (a, ctx) => {
      const e = firstError(a, ctx);
      if (e) return e;
      const target = argValue(a[0], ctx);
      const haystack = argRow(a[1], ctx);
      const returns = asMatrix(a[2], ctx);
      const notFoundArg = a.length > 3 ? argValue(a[3], ctx) : undefined;
      const mode = intOr(a[4], ctx, 0);
      if (isError(mode)) return mode;
      const search = intOr(a[5], ctx, 1);
      if (isError(search)) return search;

      const order: number[] = [];
      for (let i = 0; i < haystack.length; i++) order.push(i);
      if (search < 0) order.reverse();

      let hit = -1;
      for (const i of order) {
        const v = haystack[i];
        if (mode === 2) {
          if (matchesWildcard(target, v)) { hit = i; break; }
        } else if (mode === -1) {
          if (v !== null && cmpOrdered(target, v) >= 0) { hit = i; break; }
        } else if (mode === 1) {
          if (v !== null && cmpOrdered(v, target) >= 0) { hit = i; break; }
        } else if (eqValue(v, target)) {
          hit = i;
          break;
        }
      }

      if (hit < 0) {
        if (notFoundArg !== undefined) return notFoundArg;
        return errNotFound('XLOOKUP');
      }
      // A single-column return range indexes positionally; otherwise by row.
      const flat = returns.flat();
      return returns.length > 1 || (returns[0]?.length ?? 0) > 1
        ? flat[hit]
        : flat[hit];
    },
  },

  {
    name: 'LOOKUP', min: 2, max: 3,
    fn: (a, ctx) => {
      const e = firstError(a, ctx);
      if (e) return e;
      const target = argValue(a[0], ctx);
      const keys = argRow(a[1], ctx);
      const results = a.length > 2 ? argRow(a[2], ctx) : keys;
      if (results.length !== keys.length) {
        return err('N/A', 'LOOKUP result and lookup vectors must be the same size');
      }
      const pos = searchKeys(keys, target, true);
      if (pos < 0) return errNotFound('LOOKUP');
      return results[pos];
    },
  },

  {
    name: 'INDEX', min: 1, max: 3,
    fn: (a, ctx) => {
      const m = asMatrix(a[0], ctx);
      if (!m.length) return err('VALUE', 'INDEX needs a range');

      // INDEX(range) returns the whole block.
      if (a.length === 1) return makeArray(m);

      const rowRaw = toNumber(argValue(a[1], ctx));
      if (isError(rowRaw)) return rowRaw;
      const row = Math.trunc(rowRaw);

      // Single-row or single-column ranges accept one index.
      if (a.length === 2) {
        if (m.length === 1) {
          const cols = width(m);
          if (row === 0) return makeArray([m[0].slice()]);
          if (row < 1 || row > cols) return err('REF', `INDEX column ${row} is out of range`);
          return at(m, 0, row - 1);
        }
        const cols = width(m);
        if (cols === 1) {
          if (row === 0) return makeArray(m.map((_, i) => [at(m, i, 0)]));
          if (row < 1 || row > m.length) return err('REF', `INDEX row ${row} is out of range`);
          return at(m, row - 1, 0);
        }
        if (row === 0) return makeArray(m.map((r) => r.slice()));
        if (row < 1 || row > m.length) return err('REF', `INDEX row ${row} is out of range`);
        return makeArray([m[row - 1].slice()]);
      }

      const colRaw = toNumber(argValue(a[2], ctx));
      if (isError(colRaw)) return colRaw;
      const col = Math.trunc(colRaw);
      if (row === 0 && col === 0) return makeArray(m);
      if (col < 1 || col > width(m)) return err('REF', `INDEX column ${col} is out of range`);
      if (row === 0) return makeArray(m.map((_, i) => [at(m, i, col - 1)]));
      if (row < 1 || row > m.length) return err('REF', `INDEX row ${row} is out of range`);
      return at(m, row - 1, col - 1);
    },
  },

  {
    name: 'MATCH', min: 2, max: 3,
    fn: (a, ctx) => {
      const e = firstError(a, ctx);
      if (e) return e;
      const target = argValue(a[0], ctx);
      const haystack = argRow(a[1], ctx);
      const type = intOr(a[2], ctx, 1);
      if (isError(type)) return type;

      if (type === 0) {
        const pattern = typeof target === 'string' && /[*?]/.test(target) ? target : null;
        for (let i = 0; i < haystack.length; i++) {
          const v = haystack[i];
          if (pattern ? matchesWildcard(pattern, v) : eqValue(v, target)) return i + 1;
        }
        return errNotFound('MATCH');
      }

      const nonBlank = haystack.filter((v) => v !== null);
      if (type > 0) {
        // largest value <= target
        let best = -1;
        for (let i = 0; i < nonBlank.length; i++) {
          if (cmpOrdered(nonBlank[i], target) <= 0) best = i;
          else break;
        }
        if (best < 0) return errNotFound('MATCH');
        return indexOfKey(haystack, nonBlank[best]) + 1;
      }
      // smallest value >= target, scanning from the end
      let best = -1;
      for (let i = nonBlank.length - 1; i >= 0; i--) {
        if (cmpOrdered(nonBlank[i], target) >= 0) best = i;
        else break;
      }
      if (best < 0) return errNotFound('MATCH');
      return indexOfKey(haystack, nonBlank[best]) + 1;
    },
  },

  {
    name: 'ROW', min: 0, max: 1,
    fn: (a) => {
      const ref = a.length ? rangeOf(a[0]) : null;
      if (!ref) return err('VALUE', 'ROW() needs a reference');
      return ref.r1 + 1;
    },
  },
  {
    name: 'COLUMN', min: 0, max: 1,
    fn: (a) => {
      const ref = a.length ? rangeOf(a[0]) : null;
      if (!ref) return err('VALUE', 'COLUMN() needs a reference');
      return ref.c1 + 1;
    },
  },
  {
    name: 'ROWS', min: 1, max: 1,
    fn: (a) => {
      const ref = rangeOf(a[0]);
      if (!ref) return err('VALUE', 'ROWS() needs a reference');
      return ref.r2 - ref.r1 + 1;
    },
  },
  {
    name: 'COLUMNS', min: 1, max: 1,
    fn: (a) => {
      const ref = rangeOf(a[0]);
      if (!ref) return err('VALUE', 'COLUMNS() needs a reference');
      return ref.c2 - ref.c1 + 1;
    },
  },

  {
    // OFFSET returns a live reference so SUM(OFFSET(...)) sees the whole block.
    name: 'OFFSET', min: 3, max: 5, volatile: true,
    fn: (a, ctx) => {
      const base = rangeOf(a[0]);
      if (!base) return err('VALUE', 'OFFSET needs a reference');
      const dr = toNumber(argValue(a[1], ctx));
      if (isError(dr)) return dr;
      const dc = toNumber(argValue(a[2], ctx));
      if (isError(dc)) return dc;
      const height = a.length > 3 ? intOr(a[3], ctx, base.r2 - base.r1 + 1) : base.r2 - base.r1 + 1;
      if (isError(height)) return height;
      const widthArg = a.length > 4 ? intOr(a[4], ctx, base.c2 - base.c1 + 1) : base.c2 - base.c1 + 1;
      if (isError(widthArg)) return widthArg;
      if (height < 1 || widthArg < 1) return err('REF', 'OFFSET height/width must be at least 1');

      const r1 = base.r1 + Math.trunc(dr);
      const c1 = base.c1 + Math.trunc(dc);
      const r2 = r1 + height - 1;
      const c2 = c1 + widthArg - 1;
      if (r1 < 0 || c1 < 0) return err('REF', 'OFFSET moved above or left of column A');
      const ref: RangeRef = {
        __range: true,
        sheet: base.sheet,
        r1,
        c1,
        r2,
        c2,
      };
      return ref as unknown as Value;
    },
  },

  {
    name: 'INDIRECT', min: 1, max: 2,
    fn: (a, ctx) => {
      const text = toText(argValue(a[0], ctx));
      const parsed = parseRangeExpr(text);
      if (!parsed) return err('REF', `INDIRECT cannot read "${text}"`);
      if (parsed.sheet && parsed.sheet.toLowerCase() !== ctx.sheet.name.toLowerCase()) {
        return err('REF', 'INDIRECT can only reference the current sheet');
      }
      const ref: RangeRef = {
        __range: true,
        sheet: ctx.sheet,
        r1: parsed.r1,
        c1: parsed.c1,
        r2: parsed.r2,
        c2: parsed.c2,
      };
      return ref as unknown as Value;
    },
  },

  {
    name: 'ADDRESS', min: 2, max: 5,
    fn: (a, ctx) => {
      const row = toNumber(argValue(a[0], ctx));
      if (isError(row)) return row;
      const col = toNumber(argValue(a[1], ctx));
      if (isError(col)) return col;
      if (row < 1 || col < 1) return err('VALUE', 'ADDRESS needs row and column >= 1');
      const abs = intOr(a[2], ctx, 1);
      if (isError(abs)) return abs;
      const sheetText = a.length > 4 ? toText(argValue(a[4], ctx)) : '';
      const r = Math.trunc(row);
      const c = Math.trunc(col);
      const dollarR = abs === 1 || abs === 3 ? '$' : '';
      const dollarC = abs === 1 || abs === 2 ? '$' : '';
      const cell = `${dollarC}${colName(c - 1)}${dollarR}${r}`;
      if (!sheetText) return cell;
      return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetText)
        ? `${sheetText}!${cell}`
        : `'${sheetText.replace(/'/g, "''")}'!${cell}`;
    },
  },

  {
    name: 'CHOOSE', min: 2, max: Infinity,
    fn: (a, ctx) => {
      const idx = toNumber(argValue(a[0], ctx));
      if (isError(idx)) return idx;
      const n = Math.trunc(idx);
      if (n < 1 || n > a.length - 1) return err('VALUE', `CHOOSE index ${n} is out of range`);
      return argValue(a[n], ctx);
    },
  },

  {
    name: 'SWITCH', min: 3, max: Infinity,
    fn: (a, ctx) => {
      const subject = argValue(a[0], ctx);
      if (isError(subject)) return subject;
      let i = 1;
      for (; i + 1 < a.length; i += 2) {
        if (eqValue(subject, argValue(a[i], ctx))) return argValue(a[i + 1], ctx);
      }
      if (i < a.length) return argValue(a[i], ctx);
      return err('N/A', 'SWITCH found no match and no default');
    },
  },

  {
    name: 'TRANSPOSE', min: 1, max: 1,
    fn: (a, ctx) => makeArray(transpose(argMatrix(a[0], ctx)).rows),
  },

  {
    name: 'HYPERLINK', min: 1, max: 2,
    fn: (a, ctx) => {
      const link = toText(argValue(a[0], ctx));
      if (a.length > 1) return toText(argValue(a[1], ctx));
      return link;
    },
  },
];

function truthy4(arg: Lifted | undefined, ctx: FnCtx, fallback: boolean): boolean {
  if (arg === undefined) return fallback;
  const v = argValue(arg, ctx);
  if (v === null) return fallback;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return fallback;
  const up = toText(v).toUpperCase();
  if (up === 'TRUE') return true;
  if (up === 'FALSE') return false;
  return fallback;
}

function intOr(arg: Lifted | undefined, ctx: FnCtx, fallback: number): number | CellError {
  if (arg === undefined) return fallback;
  const v = argValue(arg, ctx);
  if (v === null) return fallback;
  const n = toNumber(v);
  if (isError(n)) return n;
  return Math.trunc(n);
}

/* ---------------------------------------------------------------- database */

interface Criteria {
  values: Value[];
  match: (v: Value) => boolean;
}

/** Build a predicate for a criteria cell ("<5", ">=x", "a*", "=text"). */
function criterion(crit: Value): (v: Value) => boolean {
  if (typeof crit === 'number') return (v) => typeof v === 'number' && v === crit;
  if (typeof crit === 'boolean') return (v) => typeof v === 'boolean' && v === crit;
  const text = toText(crit).trim();
  const m = /^(>=|<=|<>|>|<|=)([\s\S]*)$/.exec(text);
  if (!m) {
    if (/[*?]/.test(text)) {
      const re = wildcardRe(text);
      return (v) => typeof v === 'string' && re.test(v);
    }
    return (v) => typeof v === 'string' && v.toLowerCase() === text.toLowerCase();
  }
  const op = m[1];
  const operand = m[2].trim();
  const num = Number(operand);
  const isNum = operand !== '' && !Number.isNaN(num);

  if (op === '=' || op === '<>') {
    const hit = isNum
      ? (v: Value) => typeof v === 'number' && v === num
      : (v: Value) => typeof v === 'string' && v.toLowerCase() === operand.toLowerCase();
    return op === '=' ? hit : (v) => !hit(v);
  }

  const cmp = (v: number): number => (v < num ? -1 : v > num ? 1 : 0);
  const only = (v: Value): boolean => typeof v === 'number';
  if (op === '>') return (v) => only(v) && cmp(v as number) > 0;
  if (op === '<') return (v) => only(v) && cmp(v as number) < 0;
  if (op === '>=') return (v) => only(v) && cmp(v as number) >= 0;
  return (v) => only(v) && cmp(v as number) <= 0;
}

/** Resolve a field given either its name or a 1-based column number. */
function resolveField(db: Matrix, field: Value): number {
  if (typeof field === 'number') return Math.trunc(field) - 1;
  const wanted = toText(field).toLowerCase();
  const header = db[0] ?? [];
  for (let c = 0; c < header.length; c++) {
    if (toText(header[c]).toLowerCase() === wanted) return c;
  }
  return -1;
}

/**
 * Values of `col` for every record matching the criteria grid.
 *
 * The grid's first row names the fields; each later row is a set of
 * conditions AND-ed together, and multiple rows are OR-ed. A blank criterion
 * cell is ignored, and its column position decides which field is tested.
 */
function runCriteria(db: Matrix, col: number, critGrid: Matrix): Value[] {
  const out: Value[] = [];
  if (db.length < 2) return out;
  const header = db[0] ?? [];

  const rules = critGrid
    .slice(1)
    .filter((row) => row.some((v) => v !== null && v !== ''));
  if (rules.length === 0) rules.push([]); // no criteria: every record matches

  for (let r = 1; r < db.length; r++) {
    const record = db[r];
    const matched = rules.some((rule) =>
      rule.every((raw, c) => {
        if (raw === null || raw === '') return true; // ignored criterion
        if (c >= header.length) return true;
        return criterion(raw)(record[c] ?? null);
      }),
    );
    if (matched) out.push(record[col] ?? null);
  }
  return out;
}

function buildCriteria(a: Lifted[], ctx: FnCtx): { db: Matrix; col: number; crit: Matrix } | CellError {
  const db = asMatrix(a[0], ctx);
  if (db.length < 2) return err('VALUE', 'The database needs a header row and at least one record');
  const col = resolveField(db, argValue(a[1], ctx));
  if (col < 0 || col >= width(db)) return err('VALUE', 'Unknown database field');
  const crit = asMatrix(a[2], ctx);
  return { db, col, crit };
}

export const databaseFns: FnDef[] = [
  {
    name: 'DSUM', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => vals.reduce((s, v) => s + v, 0), 0),
  },
  {
    name: 'DCOUNT', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => vals.length, 0, true),
  },
  {
    name: 'DCOUNTA', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => vals.length, 0, false),
  },
  {
    name: 'DAVERAGE', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => (vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : NaN), err('DIV/0')),
  },
  {
    name: 'DMAX', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => (vals.length ? Math.max(...vals) : NaN), err('NUM')),
  },
  {
    name: 'DMIN', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => (vals.length ? Math.min(...vals) : NaN), err('NUM')),
  },
  {
    name: 'DPRODUCT', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => (vals.length ? vals.reduce((s, v) => s * v, 1) : 0), 0),
  },
  {
    name: 'DSTDEV', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => (vals.length > 1 ? Math.sqrt(varianceOf(vals, true)) : NaN), err('DIV/0'), true),
  },
  {
    name: 'DSTDEVP', min: 3, max: 3,
    fn: (a, ctx) => dbAgg(a, ctx, (vals) => (vals.length ? Math.sqrt(varianceOf(vals, false)) : NaN), err('DIV/0'), true),
  },
];

function varianceOf(xs: number[], sample: boolean): number {
  const m = xs.reduce((s, v) => s + v, 0) / xs.length;
  const ss = xs.reduce((s, v) => s + (v - m) ** 2, 0);
  return ss / (sample ? xs.length - 1 : xs.length);
}

function dbAgg(
  a: Lifted[],
  ctx: FnCtx,
  reduce: (nums: number[]) => number,
  empty: Value,
  numericOnly = true,
): Value {
  const built = buildCriteria(a, ctx);
  if (isError(built)) return built;
  const values = runCriteria(built.db, built.col, built.crit);
  const nums = numericOnly
    ? values.filter((v): v is number => typeof v === 'number')
    : (values.map((v) => toNumber(v)) as number[]).filter((v) => typeof v === 'number');
  if (isError(empty) && !nums.length) return empty;
  const out = reduce(nums);
  if (Number.isNaN(out)) return empty;
  return out;
}

export { argFlat };
