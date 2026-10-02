/**
 * Information functions: ISBLANK/ISNUMBER-family, ERROR handling and CELL/INFO.
 */

import { err, isError, type Value } from '../../types';
import { toNumber, toText } from '../../coerce';
import { isArray, isRange, type FnCtx, type FnDef, type Lifted } from '../fntypes';
import { firstScalar } from '../helpers';

type Kind = 'number' | 'text' | 'logical' | 'error' | 'blank' | 'mixed';

function kindOf(v: Lifted, ctx: FnCtx): Kind {
  const vals = isRange(v) ? ctx.values(v) : isArray(v) ? v.rows.flat() : [v as Value];
  if (!vals.length) return 'blank';
  let hasNum = false;
  let hasText = false;
  let hasBool = false;
  let hasErr = false;
  for (const x of vals) {
    if (x === null || x === '') continue;
    if (isError(x)) hasErr = true;
    else if (typeof x === 'number') hasNum = true;
    else if (typeof x === 'boolean') hasBool = true;
    else hasText = true;
  }
  const present = [hasNum, hasText, hasBool, hasErr].filter(Boolean).length;
  if (present === 0) return 'blank';
  if (present > 1) return 'mixed';
  if (hasNum) return 'number';
  if (hasText) return 'text';
  if (hasBool) return 'logical';
  return 'error';
}

/** IS* functions return FALSE rather than an error, even for bad input. */
const isFn =
  (test: (k: Kind) => boolean) =>
  (args: Lifted[], ctx: FnCtx): Value =>
    test(kindOf(args[0] ?? null, ctx));

export const infoFns: FnDef[] = [
  { name: 'ISBLANK', min: 1, max: 1, fn: isFn((k) => k === 'blank') },
  { name: 'ISNUMBER', min: 1, max: 1, fn: isFn((k) => k === 'number') },
  { name: 'ISTEXT', min: 1, max: 1, fn: isFn((k) => k === 'text') },
  { name: 'ISNONTEXT', min: 1, max: 1, fn: isFn((k) => k !== 'text') },
  { name: 'ISLOGICAL', min: 1, max: 1, fn: isFn((k) => k === 'logical') },
  { name: 'ISERROR', min: 1, max: 1, fn: isFn((k) => k === 'error') },
  { name: 'ISERR', min: 1, max: 1, fn: (a, ctx) => {
    const k = kindOf(a[0] ?? null, ctx);
    return k === 'error' && !valsAreNA(a[0], ctx);
  } },
  { name: 'ISNA', min: 1, max: 1, fn: (a, ctx) => valsAreNA(a[0] ?? null, ctx) },
  { name: 'ISREF', min: 1, max: 1, fn: (a) => isRange(a[0] ?? null) },
  { name: 'ISFORMULA', min: 1, max: 1, fn: () => false },

  {
    name: 'NA', min: 0, max: 0,
    fn: () => err('N/A', '#N/A'),
  },
  {
    name: 'ERROR.TYPE', min: 1, max: 1,
    fn: (a, ctx) => {
      const v = firstScalar(a, ctx);
      if (!isError(v)) return err('N/A', 'ERROR.TYPE needs an error value');
      const order = ['NULL', 'DIV/0', 'VALUE', 'REF', 'NAME', 'NUM', 'N/A'];
      const i = order.indexOf(v.__error);
      return i < 0 ? err('N/A', 'Unknown error type') : i + 1;
    },
  },
  {
    name: 'TYPE', min: 1, max: 1,
    fn: (a, ctx) => {
      const v = firstScalar(a, ctx);
      if (isRange(a[0] ?? null)) {
        const ref = a[0] as Extract<Lifted, { __range: true }>;
        return ref.r1 === ref.r2 && ref.c1 === ref.c2 ? 1 : 64;
      }
      if (typeof v === 'number' || v === null) return 1;
      if (typeof v === 'string') return 2;
      if (typeof v === 'boolean') return 4;
      return 16;
    },
  },
  {
    name: 'INFO', min: 1, max: 1,
    fn: (a, ctx) => {
      const kind = toText(firstScalar(a, ctx)).toUpperCase();
      switch (kind) {
        case 'OSVERSION': return 'pcdos';
        case 'RELEASE': return '1.0';
        case 'SYSTEM': return 'pcdos';
        case 'NUMFILE': return 1;
        case 'DIRECTORY': return 'C:\\';
        case 'NFILE': return 1;
        case 'PCTCHANGE': return '0.00';
        case 'CD': return 'C:\\';
        case 'CALC': return 'Automatic';
        default: return 'N/A';
      }
    },
  },
  {
    // CELL(info_type, [reference]) — returns metadata about a reference.
    name: 'CELL', min: 1, max: 2,
    fn: (a, ctx) => {
      const what = toText(firstScalar([a[0]], ctx)).toUpperCase();
      if (a.length < 2) return err('VALUE', 'CELL needs a reference');
      const ref = a[1];
      if (!isRange(ref)) return err('VALUE', 'CELL needs a reference');
      const r = ref as Extract<Lifted, { __range: true }>;
      const value = firstScalar([ref], ctx);
      switch (what) {
        case 'ADDRESS': return `${columnLetter(r.c1)}${r.r1 + 1}`;
        case 'COL': return r.c1 + 1;
        case 'ROW': return r.r1 + 1;
        case 'WIDTH': return 8;
        case 'FORMAT':
        case 'PATTERN': return 'General';
        case 'TYPE': return value === null ? 'b' : 'v';
        case 'PROTECT': return 1;
        case 'PARENTHESES': return 0;
        case 'PREFIX': return "'";
        case 'COLOR': return 0;
        case 'FILENAME': return '';
        case 'CONTENTS': return value;
        default: return err('VALUE', `CELL does not understand "${what}"`);
      }
    },
  },
  {
    name: 'SHEET', min: 0, max: 1,
    fn: (a, ctx) => {
      if (!a.length) return sheetIndexOf(ctx);
      const v = a[0];
      if (isRange(v)) {
        const ref = v as Extract<Lifted, { __range: true }>;
        return sheetIndexOf(ctx, ref.sheet);
      }
      return err('VALUE', 'SHEET needs a sheet reference');
    },
  },
  {
    name: 'SHEETS', min: 0, max: 1,
    fn: () => err('VALUE', 'SHEETS is not supported'),
  },
  {
    name: 'REGEXTEST', min: 2, max: 3,
    fn: (a, ctx) => {
      const text = toText(firstScalar([a[0]], ctx));
      const pattern = toText(firstScalar([a[1]], ctx));
      const flags = a.length > 2 ? toText(firstScalar([a[2]], ctx)) : '';
      const extra = flags.includes('i') ? 'i' : '';
      try {
        return new RegExp(pattern, extra).test(text);
      } catch {
        return err('VALUE', 'REGEXTEST pattern is invalid');
      }
    },
  },
  {
    name: 'REGEXEXTRACT', min: 2, max: 3,
    fn: (a, ctx) => {
      const text = toText(firstScalar([a[0]], ctx));
      const pattern = toText(firstScalar([a[1]], ctx));
      const flags = a.length > 2 ? toText(firstScalar([a[2]], ctx)) : '';
      try {
        const m = new RegExp(pattern, flags.includes('i') ? 'i' : '').exec(text);
        if (!m) return err('N/A', 'REGEXEXTRACT found no match');
        return m.length > 1 ? m[1] : m[0];
      } catch {
        return err('VALUE', 'REGEXEXTRACT pattern is invalid');
      }
    },
  },
  {
    name: 'REGEXREPLACE', min: 3, max: 4,
    fn: (a, ctx) => {
      const text = toText(firstScalar([a[0]], ctx));
      const pattern = toText(firstScalar([a[1]], ctx));
      const replacement = toText(firstScalar([a[2]], ctx));
      const flags = a.length > 3 ? toText(firstScalar([a[3]], ctx)) : '';
      try {
        return text.replace(new RegExp(pattern, flags.includes('g') || flags.includes('i') ? flags : `${flags}g`), replacement);
      } catch {
        return err('VALUE', 'REGEXREPLACE pattern is invalid');
      }
    },
  },
];

function valsAreNA(v: Lifted, ctx: FnCtx): boolean {
  const vals = isRange(v) ? ctx.values(v) : isArray(v) ? v.rows.flat() : [v as Value];
  return vals.length > 0 && vals.every((x) => isError(x) && x.__error === 'N/A');
}

function sheetIndexOf(ctx: FnCtx, sheet?: SheetLike | null): Value {
  const all = ctx.sheet;
  void all;
  void sheet;
  return 1;
}

interface SheetLike {
  id: string;
}

function columnLetter(c: number): string {
  let n = c + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export { toNumber };
