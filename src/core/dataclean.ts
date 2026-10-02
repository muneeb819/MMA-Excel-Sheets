/**
 * Data cleaning and profiling transforms.
 *
 * These are the "Data" ribbon operations: trim/case normalisation, numeric and
 * date coercion, duplicate removal, splitting columns and a quick profile of
 * each column so problems are visible before they reach a chart.
 */

import { err, isError, type Range, type Sheet, type Value } from './types';
import { iterRange, normRange } from './ref';
import { dateToSerial, parseDateish, toText } from './coerce';
import type { Engine } from './engine';

/** What a column looks like once its non-blank values are examined. */
export type ColumnKind = 'number' | 'date' | 'text' | 'boolean' | 'empty' | 'mixed';

export interface ColumnProfile {
  index: number;
  header: string;
  kind: ColumnKind;
  total: number;
  blanks: number;
  numbers: number;
  dates: number;
  text: number;
  booleans: number;
  errors: number;
  duplicates: number;
  /** suggested action for the Data tab */
  suggestion: string;
}

export interface ProfileOptions {
  /** treat the first row of the range as column names rather than data */
  hasHeader?: boolean;
}

/** Scan a range and describe every column. */
export function profileColumns(
  engine: Engine,
  sheet: Sheet,
  range: Range,
  options: ProfileOptions = {},
): ColumnProfile[] {
  const n = normRange(range);
  const width = n.c2 - n.c1 + 1;
  const headerRow = options.hasHeader ? n.r1 : -1;
  const firstDataRow = options.hasHeader ? n.r1 + 1 : n.r1;
  const out: ColumnProfile[] = [];

  for (let c = n.c1; c <= n.c2; c++) {
    const header = headerRow >= 0 ? toText(engine.valueAt(sheet, headerRow, c)) : '';

    let total = 0;
    let blanks = 0;
    let numbers = 0;
    let dates = 0;
    let text = 0;
    let booleans = 0;
    let errors = 0;
    const seen = new Map<string, number>();

    for (let r = firstDataRow; r <= n.r2; r++) {
      const v = engine.valueAt(sheet, r, c);
      if (v === null || v === '') {
        blanks++;
        continue;
      }
      total++;
      if (isError(v)) {
        errors++;
        continue;
      }
      const key = String(v).trim().toLowerCase();
      seen.set(key, (seen.get(key) ?? 0) + 1);

      if (typeof v === 'number') {
        numbers++;
        // Integers in the plausible date range may really be dates.
        if (v > 20000 && v < 80000 && Number.isInteger(v)) dates++;
      } else if (typeof v === 'boolean') {
        booleans++;
      } else {
        text++;
        if (parseDateish(v) !== null) dates++;
      }
    }

    let duplicates = 0;
    for (const count of seen.values()) if (count > 1) duplicates += count - 1;

    const kind: ColumnKind = classify({ numbers, dates, text, booleans, total });
    out.push({
      index: c - n.c1,
      header,
      kind,
      total: total + blanks,
      blanks,
      numbers,
      dates,
      text,
      booleans,
      errors,
      duplicates,
      suggestion: suggest({ kind, blanks, total: total + blanks, errors, duplicates, dates }),
    });
  }
  return out;
}

function classify(x: {
  numbers: number;
  dates: number;
  text: number;
  booleans: number;
  total: number;
}): ColumnKind {
  if (x.total === 0) return 'empty';
  if (x.booleans === x.total) return 'boolean';
  if (x.numbers + x.dates === x.total) return x.dates > x.numbers ? 'date' : 'number';
  if (x.text === x.total) return 'text';
  if (x.text > 0 && x.numbers + x.dates > 0) return 'mixed';
  return 'mixed';
}

function suggest(x: {
  kind: ColumnKind;
  blanks: number;
  total: number;
  errors: number;
  duplicates: number;
  dates: number;
}): string {
  if (x.kind === 'empty') return 'Empty — safe to remove';
  if (x.errors > 0) return 'Contains errors';
  if (x.duplicates > 0) return `${x.duplicates} duplicate value(s)`;
  if (x.total > 0 && x.blanks / x.total > 0.4) return 'Mostly empty';
  if (x.kind === 'mixed' && x.dates > 0) return 'Mixed — try Convert to date';
  if (x.kind === 'text' && x.dates > 0) return 'Text dates — convert';
  return 'Looks clean';
}

/* ------------------------------------------------------------- transforms */

export interface TransformResult {
  changed: number;
  description: string;
}

type CellFn = (v: Value) => Value;

/** Apply a per-cell function across a range as one undoable step. */
export function applyTransform(
  engine: Engine,
  sheet: Sheet,
  range: Range,
  label: string,
  fn: CellFn,
): TransformResult {
  const n = normRange(range);
  let changed = 0;
  engine.transact(() => {
    for (const { r, c } of iterRange(n)) {
      const before = engine.valueAt(sheet, r, c);
      if (before === null) continue;
      const after = fn(before);
      if (after === before) continue;
      engine.setValue(sheet, r, c, after, engine.cellAt(sheet, r, c)?.styleId);
      changed++;
    }
  });
  engine.recalc();
  return { changed, description: `${label}: ${changed} cell(s) updated` };
}

export const transforms = {
  trim: (v: Value): Value =>
    typeof v === 'string' ? v.trim() : v,

  collapseSpaces: (v: Value): Value =>
    typeof v === 'string' ? v.replace(/\s+/g, ' ') : v,

  upper: (v: Value): Value => (typeof v === 'string' ? v.toUpperCase() : v),

  lower: (v: Value): Value => (typeof v === 'string' ? v.toLowerCase() : v),

  proper: (v: Value): Value =>
    typeof v === 'string'
      ? v.replace(/(^|\s)\p{L}/gu, (m) => m.toUpperCase()).replace(/\p{L}+/gu, (w) =>
          w.charAt(0) + w.slice(1).toLowerCase(),
        )
      : v,

  toNumber: (v: Value): Value => {
    if (typeof v === 'number') return v;
    if (typeof v !== 'string') return v;
    const cleaned = v.replace(/[,\s]/g, '').replace(/%$/, '');
    const n = Number(cleaned);
    if (!Number.isFinite(n)) return v;
    return v.trim().endsWith('%') ? n / 100 : n;
  },

  toDate: (v: Value): Value => {
    if (typeof v === 'number') return v;
    if (typeof v !== 'string') return v;
    const serial = parseDateish(v);
    return serial === null ? v : serial;
  },

  toDateSerial: (v: Value): Value => {
    const serial = parseDateish(toText(v));
    return serial === null ? v : serial;
  },

  /** Remove thousands separators and stray symbols. */
  stripNonNumeric: (v: Value): Value => {
    if (typeof v === 'number') return v;
    if (typeof v !== 'string') return v;
    const n = Number(v.replace(/[^0-9.\-]/g, ''));
    return Number.isFinite(n) ? n : v;
  },

  zeroToBlank: (v: Value): Value => (v === 0 ? null : v),

  blankToZero: (v: Value): Value => (v === null || v === '' ? 0 : v),

  /** Replace occurrences of a substring; `find` empty means strip all digits. */
  replace: (find: string, replacement: string): CellFn => (v: Value): Value =>
    typeof v === 'string' ? v.split(find).join(replacement) : v,

  /** Keep only letters (useful for postcode-like codes). */
  lettersOnly: (v: Value): Value =>
    typeof v === 'string' ? v.replace(/[^\p{L}]+/gu, '') : v,

  digitsOnly: (v: Value): Value =>
    typeof v === 'string' ? v.replace(/\D+/g, '') : v,
};

/**
 * Remove rows whose key columns are all blank.
 * Returns the number of rows deleted.
 */
export function removeEmptyRows(engine: Engine, sheet: Sheet, range: Range): number {
  const n = normRange(range);
  const doomed: number[] = [];
  for (let r = n.r1; r <= n.r2; r++) {
    let empty = true;
    for (let c = n.c1; c <= n.c2; c++) {
      const v = engine.valueAt(sheet, r, c);
      if (v !== null && v !== '') {
        empty = false;
        break;
      }
    }
    if (empty) doomed.push(r);
  }
  if (!doomed.length) return 0;

  engine.transact(() => {
    // Delete from the bottom so earlier indices stay valid.
    for (let i = doomed.length - 1; i >= 0; i--) {
      engine.deleteRows(sheet, doomed[i], 1);
    }
  });
  engine.recalc();
  return doomed.length;
}

/**
 * Remove duplicate rows, keeping the first occurrence.
 * Returns how many rows were removed.
 */
export function removeDuplicateRows(engine: Engine, sheet: Sheet, range: Range): number {
  const n = normRange(range);
  const seen = new Set<string>();
  const doomed: number[] = [];

  for (let r = n.r1; r <= n.r2; r++) {
    const parts: string[] = [];
    for (let c = n.c1; c <= n.c2; c++) parts.push(toText(engine.valueAt(sheet, r, c)).trim().toLowerCase());
    const key = parts.join('');
    if (seen.has(key)) doomed.push(r);
    else seen.add(key);
  }

  if (!doomed.length) return 0;
  engine.transact(() => {
    for (let i = doomed.length - 1; i >= 0; i--) engine.deleteRows(sheet, doomed[i], 1);
  });
  engine.recalc();
  return doomed.length;
}

/** Replace an error value in a range with a fallback. */
export function replaceErrors(
  engine: Engine,
  sheet: Sheet,
  range: Range,
  replacement: Value = '',
): number {
  const n = normRange(range);
  let changed = 0;
  engine.transact(() => {
    for (const { r, c } of iterRange(n)) {
      const v = engine.valueAt(sheet, r, c);
      if (!isError(v)) continue;
      engine.setValue(sheet, r, c, replacement, engine.cellAt(sheet, r, c)?.styleId);
      changed++;
    }
  });
  engine.recalc();
  return changed;
}

/**
 * Split one column on a separator into several columns to its right.
 */
export function splitColumn(
  engine: Engine,
  sheet: Sheet,
  range: Range,
  separator: string,
  maxParts = 8,
): number {
  const n = normRange(range);
  const col = n.c1;
  const writeStart = n.c2 + 1;
  let widest = 0;

  const rows: string[][] = [];
  for (let r = n.r1; r <= n.r2; r++) {
    const parts = toText(engine.valueAt(sheet, r, col)).split(separator);
    const trimmed = parts.slice(0, maxParts).map((p) => p.trim());
    widest = Math.max(widest, trimmed.length);
    rows.push(trimmed);
  }

  engine.transact(() => {
    for (let r = n.r1; r <= n.r2; r++) {
      for (let i = 0; i < widest; i++) {
        engine.setValue(sheet, r, writeStart + i, rows[r - n.r1][i] ?? null);
      }
    }
  });
  engine.recalc();
  return widest;
}

/** Pad a numeric-looking string with leading zeroes to `width` characters. */
export function padZeros(engine: Engine, sheet: Sheet, range: Range, width: number): number {
  const n = normRange(range);
  let changed = 0;
  engine.transact(() => {
    for (const { r, c } of iterRange(n)) {
      const v = engine.valueAt(sheet, r, c);
      if (typeof v !== 'string') continue;
      if (!/^\d+$/.test(v.trim())) continue;
      const padded = v.trim().padStart(width, '0');
      if (padded === v) continue;
      engine.setValue(sheet, r, c, padded, engine.cellAt(sheet, r, c)?.styleId);
      changed++;
    }
  });
  engine.recalc();
  return changed;
}

/** A date serial from loose text, exposed for the cleaner's date column. */
export { dateToSerial, err };
