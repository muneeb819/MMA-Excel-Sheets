/**
 * Text and single-cell information functions.
 *
 * These follow Excel's string semantics closely: text arguments are coerced
 * with `toText`, numbers go through the number-format engine (`formatValue`),
 * and every error found in the inputs is propagated untouched. Positions and
 * counts are 1-based and truncated, ranges collapse to their top-left cell,
 * and SEARCH honours the `*` / `?` / `~` wildcards while FIND is literal.
 */

import type { FnCtx, FnDef, Lifted } from '../fntypes';
import { allValues, firstError, firstScalar, firstText } from '../helpers';
import { err, isError, type CellError, type Value } from '../../types';
import { formatValue, toNumber, toText } from '../../coerce';

/* ------------------------------------------------------------------ helpers */

/** Text form of argument `i`; a missing argument counts as empty. */
function at(args: Lifted[], ctx: FnCtx, i: number): string | CellError {
  if (i >= args.length) return '';
  return firstText([args[i] as Lifted], ctx);
}

/** Apply `f` to a text argument, letting errors through untouched. */
function map(v: Value, f: (s: string) => Value): Value {
  if (isError(v)) return v;
  return f(toText(v));
}

/** Numeric argument `i`, using `fallback` only when the argument is absent. */
function numAt(
  args: Lifted[],
  ctx: FnCtx,
  i: number,
  fallback: number,
): number | CellError {
  if (i >= args.length) return fallback;
  const v = firstScalar([args[i] as Lifted], ctx);
  if (isError(v)) return v;
  return toNumber(v);
}

/** `numAt` truncated to a whole number, as Excel does for positions/counts. */
function intAt(
  args: Lifted[],
  ctx: FnCtx,
  i: number,
  fallback: number,
): number | CellError {
  const n = numAt(args, ctx, i, fallback);
  return typeof n === 'number' ? Math.trunc(n) : n;
}

/** Excel truthiness: FALSE, 0 and empty text are false. */
function truthy(v: Value): boolean {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return true;
  const s = v.trim().toUpperCase();
  return s !== 'FALSE' && s !== '';
}

const RE_SPECIAL = /[.*+?^$()|[\]\\]/g;
const escapeRe = (s: string): string => s.replace(RE_SPECIAL, '\\$&');

/** Translate a SEARCH pattern (`*`, `?`, `~`) into a case-insensitive regex. */
function wildcard(needle: string): RegExp {
  let src = '';
  for (let i = 0; i < needle.length; i++) {
    const ch = needle[i];
    if (ch === '~' && i + 1 < needle.length) {
      src += escapeRe(needle[i + 1]);
      i++;
    } else if (ch === '*') src += '[\\s\\S]*';
    else if (ch === '?') src += '[\\s\\S]';
    else src += escapeRe(ch);
  }
  return new RegExp(src, 'gi');
}

/** Insert thousands separators into the integer part of a rendered number. */
function groupThousands(body: string): string {
  const neg = body.startsWith('-');
  const s = neg ? body.slice(1) : body;
  const dot = s.indexOf('.');
  const int = dot < 0 ? s : s.slice(0, dot);
  if (!/^\d+$/.test(int)) return body;
  const rest = dot < 0 ? '' : s.slice(dot);
  return (neg ? '-' : '') + int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + rest;
}

/** Pad the fraction out to exactly `d` digits (the format engine trims zeros). */
function padDecimals(body: string, d: number): string {
  if (d <= 0) return body;
  const dot = body.indexOf('.');
  if (dot < 0) return `${body}.${'0'.repeat(d)}`;
  return body.slice(0, dot) + '.' + body.slice(dot + 1).slice(0, d).padEnd(d, '0');
}

/** Fixed-point rendering with optional grouping; keeps the leading `-`. */
function fixedNumber(n: number, decimals: number, grouped: boolean): string {
  let v = n;
  let d = Math.min(Math.max(Math.trunc(decimals), 0), 30);
  if (decimals < 0) {
    const f = 10 ** Math.min(-decimals, 30);
    v = Math.sign(v) * Math.round(Math.abs(v) / f) * f;
    d = 0;
  }
  const int = grouped ? '#,##0' : '0';
  const pattern = d > 0 ? `${int}.${'0'.repeat(d)}` : int;
  const body = padDecimals(formatValue(v, pattern), d);
  return grouped ? groupThousands(body) : body;
}

/** Drop a stray currency symbol ahead of explicit number parsing. */
function uncurry(s: string): string {
  return s.replace(/[$€£¥]/g, '');
}

/** Full-width ASCII and the ideographic space folded back to half-width. */
function halfWidth(v: string | Value): string | Value {
  if (isError(v)) return v;
  let out = '';
  for (const ch of toText(v)) {
    const cp = ch.codePointAt(0);
    if (cp === undefined) continue;
    if (cp === 0x3000) out += ' ';
    else if (cp >= 0xff01 && cp <= 0xff5e) out += String.fromCharCode(cp - 0xfee0);
    else out += ch;
  }
  return out;
}

/** Join every value across the arguments, error propagation aside. */
function concatAll(args: Lifted[], ctx: FnCtx): Value {
  const e = firstError(args, ctx);
  if (e) return e;
  return allValues(args, ctx).map((v) => toText(v)).join('');
}

/* ------------------------------------------------------------- the functions */

export const textFns: FnDef[] = [
  { name: 'TEXT', min: 2, fn: (args, ctx) => {
    const v = firstScalar(args, ctx);
    if (isError(v)) return v;
    const f = at(args, ctx, 1);
    if (isError(f)) return f;
    return formatValue(v, f);
  } },

  { name: 'DOLLAR', min: 1, max: 2, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const n = numAt(args, ctx, 0, 0);
    if (isError(n)) return n;
    const d = intAt(args, ctx, 1, 2);
    if (isError(d)) return d;
    const body = `$${fixedNumber(Math.abs(n), d, true)}`;
    return n < 0 ? `(${body})` : body;
  } },

  { name: 'FIXED', min: 1, max: 3, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const n = numAt(args, ctx, 0, 0);
    if (isError(n)) return n;
    const d = intAt(args, ctx, 1, 2);
    if (isError(d)) return d;
    const noCommas = args.length > 2 && truthy(firstScalar([args[2] as Lifted], ctx));
    return fixedNumber(n, d, !noCommas);
  } },

  { name: 'VALUE', min: 1, fn: (args, ctx) => {
    const v = firstScalar(args, ctx);
    if (isError(v)) return v;
    if (typeof v === 'number') return v;
    if (v === null) return 0;
    if (typeof v !== 'string' || v.trim() === '') {
      return err('VALUE', `Cannot convert "${toText(v)}" to a number`);
    }
    return toNumber(uncurry(v.trim()));
  } },

  { name: 'NUMBERVALUE', min: 1, max: 3, fn: (args, ctx) => {
    const v = firstScalar(args, ctx);
    if (isError(v)) return v;
    if (typeof v === 'number') return v;
    if (v === null) return 0;
    if (typeof v !== 'string') {
      return err('VALUE', `Cannot convert "${toText(v)}" to a number`);
    }
    const rawDec = at(args, ctx, 1);
    const rawGrp = at(args, ctx, 2);
    if (isError(rawDec)) return rawDec;
    if (isError(rawGrp)) return rawGrp;
    const dec = rawDec === '' ? '.' : rawDec;
    const grp = rawGrp === '' ? ',' : rawGrp;
    if (dec.length > 1 || grp.length > 1) {
      return err('VALUE', 'Separators must be a single character');
    }
    if (dec === grp) return err('VALUE', 'Decimal and group separators must differ');
    let s = uncurry(v.trim());
    const pct = s.endsWith('%');
    if (pct) s = s.slice(0, -1);
    s = s.split(grp).join('').split(dec).join('.');
    if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(s)) {
      return err('VALUE', `Cannot convert "${toText(v)}" to a number`);
    }
    return pct ? Number(s) / 100 : Number(s);
  } },

  { name: 'T', min: 1, fn: (args, ctx) => {
    const v = firstScalar(args, ctx);
    if (isError(v)) return v;
    return typeof v === 'string' ? v : '';
  } },

  { name: 'CLEAN', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => s.replace(/[\x00-\x1f]/g, '')) },
  { name: 'TRIM', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => s.replace(/ {2,}/g, ' ').trim()) },
  { name: 'LOWER', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => s.toLowerCase()) },
  { name: 'UPPER', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => s.toUpperCase()) },
  { name: 'LEN', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => s.length) },
  { name: 'ASC', min: 1, fn: (args, ctx) => halfWidth(firstText(args, ctx)) },
  { name: 'JIS', min: 1, fn: (args, ctx) => halfWidth(firstText(args, ctx)) },

  { name: 'PROPER', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => s
    .toLowerCase()
    .replace(/(^|[^A-Za-z])([a-z])/g, (_m, pre: string, ch: string) => pre + ch.toUpperCase())) },

  { name: 'LEFT', min: 1, max: 2, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const s = firstText(args, ctx);
    if (isError(s)) return s;
    const n = intAt(args, ctx, 1, 1);
    if (isError(n)) return n;
    if (n < 0) return err('VALUE', 'LEFT num_chars must not be negative');
    return s.slice(0, n);
  } },

  { name: 'RIGHT', min: 1, max: 2, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const s = firstText(args, ctx);
    if (isError(s)) return s;
    const n = intAt(args, ctx, 1, 1);
    if (isError(n)) return n;
    if (n < 0) return err('VALUE', 'RIGHT num_chars must not be negative');
    return n >= s.length ? s : s.slice(s.length - n);
  } },

  { name: 'MID', min: 3, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const s = firstText(args, ctx);
    if (isError(s)) return s;
    const start = intAt(args, ctx, 1, 1);
    if (isError(start)) return start;
    const n = intAt(args, ctx, 2, 0);
    if (isError(n)) return n;
    if (start < 1) return err('VALUE', 'MID start_num must be at least 1');
    if (n < 0) return err('VALUE', 'MID num_chars must not be negative');
    return s.slice(start - 1, start - 1 + n);
  } },

  { name: 'REPLACE', min: 4, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const s = firstText(args, ctx);
    if (isError(s)) return s;
    const rep = at(args, ctx, 3);
    if (isError(rep)) return rep;
    const start = intAt(args, ctx, 1, 1);
    if (isError(start)) return start;
    const count = intAt(args, ctx, 2, 0);
    if (isError(count)) return count;
    if (start < 1) return err('VALUE', 'REPLACE start_num must be at least 1');
    if (count < 0) return err('VALUE', 'REPLACE num_chars must not be negative');
    return s.slice(0, start - 1) + rep + s.slice(start - 1 + count);
  } },

  { name: 'SUBSTITUTE', min: 3, max: 4, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const s = firstText(args, ctx);
    if (isError(s)) return s;
    const oldText = at(args, ctx, 1);
    if (isError(oldText)) return oldText;
    const newText = at(args, ctx, 2);
    if (isError(newText)) return newText;
    if (oldText === '') return s;
    if (args.length < 4) return s.split(oldText).join(newText);
    const which = intAt(args, ctx, 3, 1);
    if (isError(which)) return which;
    if (which < 1) return err('VALUE', 'SUBSTITUTE instance_num must be at least 1');
    let count = 0;
    let idx = s.indexOf(oldText);
    while (idx >= 0) {
      count++;
      idx = s.indexOf(oldText, idx + oldText.length);
    }
    if (which > count) return err('VALUE', `Instance ${which} does not exist`);
    let pos = -1;
    for (let i = 0; i < which; i++) pos = s.indexOf(oldText, pos + 1);
    return s.slice(0, pos) + newText + s.slice(pos + oldText.length);
  } },

  { name: 'REPT', min: 2, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const s = firstText(args, ctx);
    if (isError(s)) return s;
    const n = intAt(args, ctx, 1, 0);
    if (isError(n)) return n;
    if (n < 0) return err('VALUE', 'REPT num_times must not be negative');
    if (s.length * n > 32_767) return err('VALUE', 'REPT result is too long');
    return s.repeat(n);
  } },

  { name: 'FIND', min: 2, max: 3, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const needle = at(args, ctx, 0);
    if (isError(needle)) return needle;
    const hay = at(args, ctx, 1);
    if (isError(hay)) return hay;
    const start = intAt(args, ctx, 2, 1);
    if (isError(start)) return start;
    if (start < 1) return err('VALUE', 'FIND start_num must be at least 1');
    const pos = hay.indexOf(needle, start - 1);
    if (pos < 0) return err('VALUE', `FIND could not locate "${needle}"`);
    return pos + 1;
  } },

  { name: 'SEARCH', min: 2, max: 3, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const needle = at(args, ctx, 0);
    if (isError(needle)) return needle;
    const hay = at(args, ctx, 1);
    if (isError(hay)) return hay;
    const start = intAt(args, ctx, 2, 1);
    if (isError(start)) return start;
    if (start < 1) return err('VALUE', 'SEARCH start_num must be at least 1');
    if (start > hay.length + 1) {
      return err('VALUE', 'SEARCH start_num is past the end of the text');
    }
    if (needle === '') return start;
    const re = wildcard(needle);
    re.lastIndex = start - 1;
    const m = re.exec(hay);
    if (!m) return err('VALUE', `SEARCH could not locate "${needle}"`);
    return m.index + 1;
  } },

  { name: 'EXACT', min: 2, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    return firstScalar([args[0] as Lifted], ctx) === firstScalar([args[1] as Lifted], ctx);
  } },

  { name: 'CONCATENATE', min: 1, max: Infinity, fn: (args, ctx) => concatAll(args, ctx) },
  { name: 'CONCAT', min: 1, max: Infinity, fn: (args, ctx) => concatAll(args, ctx) },

  { name: 'TEXTJOIN', min: 3, max: Infinity, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const sep = at(args, ctx, 0);
    if (isError(sep)) return sep;
    const skip = truthy(firstScalar([args[1] as Lifted], ctx));
    const parts: string[] = [];
    for (const v of allValues(args.slice(2), ctx)) {
      const s = toText(v);
      if (skip && (v === null || s === '')) continue;
      parts.push(s);
    }
    return parts.join(sep);
  } },

  { name: 'CHAR', min: 1, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const n = intAt(args, ctx, 0, 0);
    if (isError(n)) return n;
    if (n < 1 || n > 255) return err('VALUE', `CHAR needs 1-255, got ${n}`);
    return String.fromCharCode(n);
  } },

  { name: 'UNICHAR', min: 1, fn: (args, ctx) => {
    const e = firstError(args, ctx);
    if (e) return e;
    const n = intAt(args, ctx, 0, 0);
    if (isError(n)) return n;
    if (n < 1 || n > 0x10ffff) return err('VALUE', `UNICHAR needs 1-1114111, got ${n}`);
    return String.fromCodePoint(n);
  } },

  { name: 'CODE', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => {
    const cp = s.codePointAt(0);
    return cp === undefined ? err('VALUE', 'CODE requires at least one character') : cp;
  }) },

  { name: 'UNICODE', min: 1, fn: (args, ctx) => map(firstText(args, ctx), (s) => {
    const cp = s.codePointAt(0);
    return cp === undefined ? err('VALUE', 'UNICODE requires at least one character') : cp;
  }) },
];
