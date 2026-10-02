/**
 * Mathematical, trigonometric, rounding and conditional-aggregate functions.
 *
 * Includes the conditional family (SUMIF/COUNTIFS/AVERAGEIF...) because those
 * are the workhorses for real data work.
 */

import { err, isError, type CellError, type Value } from '../../types';
import { toNumber } from '../../coerce';
import {
  flat,
  isArray,
  isRange,
  makeArray,
  numbersAndBools,
  numbersOnly,
  type ArrayRef,
  type FnCtx,
  type FnDef,
  type Lifted,
} from '../fntypes';
import { argFlat, argMatrix, firstError, firstNumber, numOrDefault, numOut, optNumber } from '../helpers';

/* --------------------------------------------------------------- utilities */

const num = (v: Value): number | CellError => toNumber(v);

/** Numeric argument with a default, narrowed to `number | CellError`. */
function numOr(args: Lifted[], ctx: FnCtx, i: number, fallback: number): number | CellError {
  return numOrDefault(args, ctx, i, fallback);
}

/** Integer-valued optional argument, checked for errors before truncating. */
function intArg(args: Lifted[], ctx: FnCtx, i: number, fallback: number): number | CellError {
  const v = numOrDefault(args, ctx, i, fallback);
  return isError(v) ? v : Math.trunc(v);
}

const math1 =
  (label: string, f: (x: number) => number) =>
  (args: Lifted[]): Value => {
    const x = num(args[0] as Value);
    if (isError(x)) return x;
    if (!Number.isFinite(x)) return err('NUM', `${label} received a non-finite number`);
    return numOut(f(x));
  };

/** Wrap a trig function so the domain error matches Excel's #NUM!. */
const trig =
  (name: string, f: (x: number) => number, domain: (x: number) => boolean) =>
  (args: Lifted[]): Value => {
    const x = num(args[0] as Value);
    if (isError(x)) return x;
    if (!domain(x)) return err('NUM', `${name} is undefined for ${x}`);
    return numOut(f(x));
  };

const SUBTOTAL_FN: Record<number, string> = {
  1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN',
  6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP',
};

/* ------------------------------------------------------------- aggregators */

function sumValues(args: Lifted[], ctx: FnCtx): Value {
  const e = firstError(args, ctx);
  if (e) return e;
  const nums = numbersAndBools(args, ctx);
  if (nums.length === 0) return 0;
  return nums.reduce((a, b) => a + b, 0);
}

function productValues(args: Lifted[], ctx: FnCtx): Value {
  const e = firstError(args, ctx);
  if (e) return e;
  const nums = numbersAndBools(args, ctx);
  if (nums.length === 0) return 0;
  return nums.reduce((a, b) => a * b, 1);
}

function sumSquares(args: Lifted[], ctx: FnCtx): Value {
  return numbersOnly(args, ctx).reduce((a, b) => a + b * b, 0);
}

/* -------------------------------------------------------- criteria matching */

type Predicate = (v: Value) => boolean;

/** Excel criteria strings: `>5`, `<>x`, `<=10`, `a*`, plain value. */
function criterionPredicate(crit: Value): Predicate {
  if (typeof crit === 'number') {
    return (v) => typeof v === 'number' && v === crit;
  }
  if (typeof crit === 'boolean') {
    return (v) => typeof v === 'boolean' && v === crit;
  }
  const text = toNumberOrText(crit);
  const m = /^(>=|<=|<>|>|<|=)(.*)$/.exec(text);
  if (m) {
    const op = m[1];
    const operandText = m[2].trim();
    const operandNum = Number(operandText);
    const isNum = operandText !== '' && !Number.isNaN(operandNum);

    if (op === '=' || op === '<>') {
      const pred = isNum
        ? (v: Value) => typeof v === 'number' && v === operandNum
        : (v: Value) => typeof v === 'string' && v.toLowerCase() === operandText.toLowerCase();
      return op === '=' ? pred : (v) => !pred(v);
    }
    const cmp = (v: number): number => (v < operandNum ? -1 : v > operandNum ? 1 : 0);
    if (op === '>') return (v) => typeof v === 'number' && cmp(v) > 0;
    if (op === '<') return (v) => typeof v === 'number' && cmp(v) < 0;
    if (op === '>=') return (v) => typeof v === 'number' && cmp(v) >= 0;
    return (v) => typeof v === 'number' && cmp(v) <= 0;
  }

  if (/[*?]/.test(text)) {
    const re = wildcardToRegExp(text);
    return (v) => typeof v === 'string' && re.test(v);
  }
  return (v) => typeof v === 'string' && v.toLowerCase() === text.toLowerCase();
}

function toNumberOrText(v: Value): string {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return '';
  return String(v);
}

function wildcardToRegExp(pattern: string): RegExp {
  let out = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~') {
      const nxt = pattern[i + 1];
      if (nxt === '*' || nxt === '?' || nxt === '~') {
        out += escapeRe(nxt);
        i++;
        continue;
      }
    }
    if (ch === '*') out += '[\\s\\S]*';
    else if (ch === '?') out += '[\\s\\S]';
    else out += escapeRe(ch);
  }
  return new RegExp(`^${out}$`, 'i');
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/* ------------------------------------------------------------- the library */

export const mathFns: FnDef[] = [
  /* --- basic arithmetic --- */
  { name: 'SUM', min: 1, fn: sumValues },
  { name: 'PRODUCT', min: 1, fn: productValues },
  { name: 'SUMSQ', min: 1, fn: sumSquares },
  { name: 'ABS', min: 1, max: 1, fn: math1('ABS', Math.abs) },
  { name: 'SIGN', min: 1, max: 1, fn: (a) => { const x = num(a[0] as Value); return isError(x) ? x : Math.sign(x); } },
  { name: 'SQRT', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    if (x < 0) return err('NUM', 'SQRT needs a non-negative number');
    return numOut(Math.sqrt(x));
  } },
  { name: 'EXP', min: 1, max: 1, fn: math1('EXP', Math.exp) },
  { name: 'LN', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    if (x <= 0) return err('NUM', 'LN needs a positive number');
    return numOut(Math.log(x));
  } },
  { name: 'LOG10', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    if (x <= 0) return err('NUM', 'LOG10 needs a positive number');
    return numOut(Math.log10(x));
  } },
  {
    name: 'LOG', min: 1, max: 2,
    fn: (a, ctx) => {
      const x = num(a[0] as Value);
      if (isError(x)) return x;
      const base = numOr(a, ctx, 1, 10);
      if (isError(base)) return base;
      if (x <= 0) return err('NUM', 'LOG needs a positive number');
      if (base <= 0 || base === 1) return err('NUM', 'LOG base must be positive and not 1');
      return numOut(Math.log(x) / Math.log(base));
    },
  },
  { name: 'PI', min: 0, max: 0, fn: () => Math.PI },
  { name: 'POWER', min: 2, max: 2, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    const y = num(a[1] as Value);
    if (isError(y)) return y;
    const o = Math.pow(x, y);
    return Number.isFinite(o) ? o : err('NUM');
  } },
  { name: 'INT', min: 1, max: 1, fn: (a) => { const x = num(a[0] as Value); return isError(x) ? x : Math.floor(x); } },
  { name: 'TRUNC', min: 1, max: 2, fn: (a, ctx) => truncImpl(a, ctx) },
  {
    name: 'MOD', min: 2, max: 2,
    fn: (a, ctx) => {
      const n = firstNumber(a, ctx);
      if (isError(n)) return n;
      const d = num(a[1] as Value);
      if (isError(d)) return d;
      if (d === 0) return err('DIV/0', 'MOD divisor cannot be zero');
      // Excel's MOD follows the sign of the divisor.
      return n - d * Math.floor(n / d);
    },
  },
  {
    name: 'QUOTIENT', min: 2, max: 2,
    fn: (a, ctx) => {
      const n = firstNumber(a, ctx);
      if (isError(n)) return n;
      const d = num(a[1] as Value);
      if (isError(d)) return d;
      if (d === 0) return err('DIV/0');
      return Math.trunc(n / d);
    },
  },
  {
    name: 'MROUND', min: 2, max: 2,
    fn: (a, ctx) => {
      const n = firstNumber(a, ctx);
      if (isError(n)) return n;
      const m = num(a[1] as Value);
      if (isError(m)) return m;
      if (m === 0) return 0;
      if (Math.sign(n) !== Math.sign(m)) return err('NUM', 'MROUND needs matching signs');
      return Math.round(n / m) * m;
    },
  },

  /* --- rounding --- */
  { name: 'ROUND', min: 1, max: 2, fn: (a, ctx) => roundImpl(a, ctx, Math.round) },
  { name: 'ROUNDUP', min: 1, max: 2, fn: (a, ctx) => roundImpl(a, ctx, (x) => (x < 0 ? -Math.ceil(-x) : Math.ceil(x))) },
  { name: 'ROUNDDOWN', min: 1, max: 2, fn: (a, ctx) => roundImpl(a, ctx, (x) => (x < 0 ? -Math.floor(-x) : Math.floor(x))) },
  { name: 'EVEN', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    const up = x >= 0 ? Math.ceil(x) : Math.floor(x);
    const n = up % 2 === 0 ? up : up + (Math.sign(up) || 1);
    return n === 0 ? 0 : n;
  } },
  { name: 'ODD', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    let up = x >= 0 ? Math.ceil(x) : Math.floor(x);
    if (up % 2 === 0) up += Math.sign(up) || 1;
    if (up === 0) up = 1;
    return x < 0 ? -Math.abs(up) : Math.abs(up);
  } },
  { name: 'CEILING', min: 1, max: 2, fn: (a, ctx) => ceilingImpl(a, ctx, 'legacy') },
  { name: 'CEILING.MATH', min: 1, max: 3, fn: (a, ctx) => ceilingImpl(a, ctx, 'math') },
  { name: 'ISO.CEILING', min: 1, max: 3, fn: (a, ctx) => ceilingImpl(a, ctx, 'math') },
  { name: 'FLOOR', min: 1, max: 2, fn: (a, ctx) => floorImpl(a, ctx, 'legacy') },
  { name: 'FLOOR.MATH', min: 1, max: 3, fn: (a, ctx) => floorImpl(a, ctx, 'math') },

  /* --- combinatorial --- */
  { name: 'COMBIN', min: 2, max: 2, fn: (a) => {
    const n = num(a[0] as Value);
    if (isError(n)) return n;
    const k = num(a[1] as Value);
    if (isError(k)) return k;
    if (n < 0 || k < 0 || k > n) return err('NUM', 'COMBIN needs 0 <= k <= n');
    return numOut(factorial(n) / (factorial(k) * factorial(n - k)));
  } },
  { name: 'COMBINA', min: 2, max: 2, fn: (a) => {
    const n = num(a[0] as Value);
    if (isError(n)) return n;
    const k = num(a[1] as Value);
    if (isError(k)) return k;
    if (n < 0 || k < 0) return err('NUM', 'COMBINA needs non-negative arguments');
    const nn = n + k - 1;
    return numOut(factorial(nn) / (factorial(k) * factorial(n - 1)));
  } },
  { name: 'FACT', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    if (x < 0) return err('NUM', 'FACT needs a non-negative integer');
    return numOut(gamma(x + 1));
  } },
  { name: 'FACTDOUBLE', min: 1, max: 1, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    if (x < -1) return err('NUM');
    let o = 1;
    const step = x >= 1 ? -2 : 2;
    for (let i = x; step < 0 ? i > 0 : i < 1; i += step) o *= i;
    return o;
  } },
  { name: 'MULTINOMIAL', min: 1, fn: (a, ctx) => {
    const ns = numbersOnly(a, ctx);
    if (ns.some((n) => n < 0)) return err('NUM', 'MULTINOMIAL needs non-negative integers');
    const s = ns.reduce((p, c) => p + c, 0);
    let den = 1;
    for (const n of ns) den *= factorial(n);
    return numOut(factorial(s) / den);
  } },
  { name: 'GCD', min: 1, fn: (a, ctx) => {
    const ns = numbersOnly(a, ctx);
    if (!ns.length) return 0;
    return ns.reduce((x, y) => gcd(x, y));
  } },
  { name: 'LCM', min: 1, fn: (a, ctx) => {
    const ns = numbersOnly(a, ctx);
    if (!ns.length) return 0;
    return ns.reduce((x, y) => (x && y ? Math.abs((x * y) / gcd(x, y)) : 0));
  } },

  /* --- trigonometry --- */
  { name: 'SIN', min: 1, max: 1, fn: math1('SIN', Math.sin) },
  { name: 'COS', min: 1, max: 1, fn: math1('COS', Math.cos) },
  { name: 'TAN', min: 1, max: 1, fn: math1('TAN', Math.tan) },
  { name: 'ASIN', min: 1, max: 1, fn: trig('ASIN', Math.asin, (x) => x >= -1 && x <= 1) },
  { name: 'ACOS', min: 1, max: 1, fn: trig('ACOS', Math.acos, (x) => x >= -1 && x <= 1) },
  { name: 'ATAN', min: 1, max: 1, fn: math1('ATAN', Math.atan) },
  { name: 'ATAN2', min: 2, max: 2, fn: (a) => {
    const x = num(a[0] as Value);
    if (isError(x)) return x;
    const y = num(a[1] as Value);
    if (isError(y)) return y;
    if (x === 0 && y === 0) return err('DIV/0', 'ATAN2(0,0) is undefined');
    return numOut(Math.atan2(y, x));
  } },
  { name: 'SINH', min: 1, max: 1, fn: math1('SINH', Math.sinh) },
  { name: 'COSH', min: 1, max: 1, fn: math1('COSH', Math.cosh) },
  { name: 'TANH', min: 1, max: 1, fn: math1('TANH', Math.tanh) },
  { name: 'ASINH', min: 1, max: 1, fn: math1('ASINH', Math.asinh) },
  { name: 'ACOSH', min: 1, max: 1, fn: trig('ACOSH', Math.acosh, (x) => x >= 1) },
  { name: 'ATANH', min: 1, max: 1, fn: trig('ATANH', Math.atanh, (x) => x > -1 && x < 1) },
  { name: 'COT', min: 1, max: 1, fn: trig('COT', (x) => 1 / Math.tan(x), (x) => x !== 0) },
  { name: 'SEC', min: 1, max: 1, fn: trig('SEC', (x) => 1 / Math.cos(x), (x) => Math.abs(Math.cos(x)) > 1e-12) },
  { name: 'CSC', min: 1, max: 1, fn: trig('CSC', (x) => 1 / Math.sin(x), (x) => Math.sin(x) !== 0) },
  { name: 'RADIANS', min: 1, max: 1, fn: math1('RADIANS', (x) => (x * Math.PI) / 180) },
  { name: 'DEGREES', min: 1, max: 1, fn: math1('DEGREES', (x) => (x * 180) / Math.PI) },

  /* --- number base / roman --- */
  { name: 'DECIMAL', min: 2, max: 2, fn: (a, ctx) => {
    const t = toNumberOrText(a[0] as Value);
    const base = firstNumber([a[1]], ctx);
    if (isError(base)) return base;
    if (base < 2 || base > 36) return err('NUM', 'DECIMAL base must be 2..36');
    const n = parseInt(t, base);
    return Number.isNaN(n) ? err('NUM') : n;
  } },
  { name: 'BASE', min: 2, max: 3, fn: (a, ctx) => {
    const n = firstNumber(a, ctx);
    if (isError(n)) return n;
    const base = num(a[1] as Value);
    if (isError(base)) return base;
    if (base < 2 || base > 36) return err('NUM');
    const minLen = intArg(a, ctx, 2, 0);
    if (isError(minLen)) return minLen;
    return toBase(Math.trunc(n), base, minLen);
  } },
  { name: 'ROMAN', min: 1, max: 2, fn: (a, ctx) => toRoman(a, ctx) },
  { name: 'ARABIC', min: 1, max: 1, fn: (a) => {
    const s = toNumberOrText(a[0] as Value).toUpperCase();
    let n = 0;
    let prev = 0;
    for (let i = s.length - 1; i >= 0; i--) {
      const v = romanVal(s[i]);
      if (v < prev) n -= v;
      else {
        n += v;
        prev = v;
      }
    }
    return n;
  } },

  /* --- random --- */
  { name: 'RAND', min: 0, max: 0, volatile: true, fn: () => Math.random() },
  { name: 'RANDBETWEEN', min: 2, max: 2, volatile: true, fn: (a, ctx) => {
    const lo = firstNumber(a, ctx);
    if (isError(lo)) return lo;
    const hi = num(a[1] as Value);
    if (isError(hi)) return hi;
    if (lo > hi) return err('NUM', 'RANDBETWEEN low > high');
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  } },

  /* --- matrix --- */
  { name: 'MMULT', min: 2, max: 2, fn: (a, ctx) => mmult(a, ctx) },
  { name: 'MDETERM', min: 1, max: 1, fn: (a, ctx) => {
    const m = numericMatrix(a[0], ctx);
    if (isError(m)) return m;
    const d = det(m);
    return d === null ? err('NUM', 'Matrix is singular') : numOut(d);
  } },
  { name: 'MINVERSE', min: 1, max: 1, fn: (a, ctx) => minverse(a, ctx) },
  { name: 'MUNIT', min: 1, max: 1, fn: (a, ctx) => {
    const n = firstNumber(a, ctx);
    if (isError(n)) return n;
    if (n < 1) return err('NUM');
    const rows: Value[][] = [];
    for (let i = 0; i < n; i++) {
      const row: Value[] = [];
      for (let j = 0; j < n; j++) row.push(i === j ? 1 : 0);
      rows.push(row);
    }
    return makeArray(rows);
  } },

  /* --- conditional aggregates --- */
  { name: 'SUMIF', min: 2, max: 3, fn: (a, ctx) => conditionalSum(a, ctx) },
  { name: 'SUMIFS', min: 3, max: Infinity, fn: (a, ctx) => conditionalSumMulti(a, ctx) },
  { name: 'COUNTIF', min: 2, max: 2, fn: (a, ctx) => conditionalCount(a, ctx) },
  { name: 'COUNTIFS', min: 2, max: Infinity, fn: (a, ctx) => conditionalCount(a, ctx) },
  { name: 'AVERAGEIF', min: 2, max: 3, fn: (a, ctx) => conditionalAverage(a, ctx) },
  { name: 'AVERAGEIFS', min: 3, max: Infinity, fn: (a, ctx) => conditionalAverageMulti(a, ctx) },
  { name: 'MAXIFS', min: 3, max: Infinity, fn: (a, ctx) => conditionalExtremum(a, ctx, true) },
  { name: 'MINIFS', min: 3, max: Infinity, fn: (a, ctx) => conditionalExtremum(a, ctx, false) },

  /* --- pairs --- */
  { name: 'SUMPRODUCT', min: 1, max: Infinity, fn: (a, ctx) => sumproduct(a, ctx) },
  { name: 'SUMX2MY2', min: 2, max: 2, fn: (a, ctx) => paired(a, ctx, (x, y) => x * x - y * y) },
  { name: 'SUMX2PY2', min: 2, max: 2, fn: (a, ctx) => paired(a, ctx, (x, y) => x * x + y * y) },
  { name: 'SUMXMY2', min: 2, max: 2, fn: (a, ctx) => paired(a, ctx, (x, y) => (x - y) * (x - y)) },

  { name: 'SUBTOTAL', min: 2, max: 2, fn: subtotalImpl },
  { name: 'AGGREGATE', min: 3, max: Infinity, fn: aggregateImpl },
];

/* -------------------------------------------------------------- impl pieces */

function roundImpl(args: Lifted[], ctx: FnCtx, rounder: (x: number) => number): Value {
  const x = firstNumber(args, ctx);
  if (isError(x)) return x;
  const digits = optNumber(args, ctx, 1, 0);
  if (isError(digits)) return digits;
  const f = 10 ** Math.trunc(digits);
  const scaled = x * f;
  // Guard against binary representation noise (2.675 * 100 === 267.49999...).
  const snapped = Math.abs(scaled - Math.round(scaled)) < 1e-9 ? Math.round(scaled) : scaled;
  return rounder(snapped) / f;
}

function truncImpl(args: Lifted[], ctx: FnCtx): Value {
  const x = firstNumber(args, ctx);
  if (isError(x)) return x;
  const digits = optNumber(args, ctx, 1, 0);
  if (isError(digits)) return digits;
  const f = 10 ** Math.trunc(digits);
  return Math.trunc(x * f) / f;
}

function ceilingImpl(args: Lifted[], ctx: FnCtx, mode: 'legacy' | 'math'): Value {
  const x = firstNumber(args, ctx);
  if (isError(x)) return x;
  const sig = optNumber(args, ctx, 1, 1);
  if (isError(sig)) return sig;
  if (sig === 0) return 0;
  if (mode === 'legacy') {
    if (x > 0 && sig < 0) return err('NUM', 'CEILING needs a positive significance');
    return Math.ceil(x / sig) * sig;
  }
  // CEILING.MATH rounds away from zero and ignores the sign of significance.
  const s = Math.abs(sig);
  const r = x >= 0 ? Math.ceil(x / s) * s : Math.floor(x / s) * s;
  return r === 0 ? 0 : r;
}

function floorImpl(args: Lifted[], ctx: FnCtx, mode: 'legacy' | 'math'): Value {
  const x = firstNumber(args, ctx);
  if (isError(x)) return x;
  const sig = optNumber(args, ctx, 1, 1);
  if (isError(sig)) return sig;
  if (sig === 0) return err('DIV/0', 'FLOOR significance cannot be zero');
  if (mode === 'legacy') {
    if (x > 0 && sig < 0) return err('NUM', 'FLOOR needs a positive significance');
    return Math.floor(x / sig) * sig;
  }
  const s = Math.abs(sig);
  const r = x >= 0 ? Math.floor(x / s) * s : Math.ceil(x / s) * s;
  return r === 0 ? 0 : r;
}

function factorial(n: number): number {
  const i = Math.trunc(n);
  if (i > 170) return Infinity;
  let out = 1;
  for (let k = 2; k <= i; k++) out *= k;
  return out;
}

/** Gamma via the Lanczos approximation; handles FACT's non-integer inputs. */
function gamma(z: number): number {
  const g = 7;
  const c = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028,
    771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (z < 0.5) return Math.PI / (Math.sin(Math.PI * z) * gamma(1 - z));
  const zz = z - 1;
  let x = c[0];
  for (let i = 1; i < g + 2; i++) x += c[i] / (zz + i);
  const t = zz + g + 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, zz + 0.5) * Math.exp(-t) * x;
}

function gcd(a: number, b: number): number {
  let x = Math.abs(Math.trunc(a));
  let y = Math.abs(Math.trunc(b));
  while (y) {
    const t = y;
    y = x % y;
    x = t;
  }
  return x;
}

function toBase(n: number, base: number, minLen: number): string {
  const digits = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let v = Math.abs(n);
  if (v === 0) return '0'.repeat(Math.max(1, minLen));
  let out = '';
  while (v > 0) {
    out = digits[Math.trunc(v % base)] + out;
    v = Math.floor(v / base);
  }
  if (minLen > out.length) out = '0'.repeat(minLen - out.length) + out;
  return (n < 0 ? '-' : '') + out;
}

const ROMAN_TABLE: [number, string][] = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
  [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
];

function toRoman(args: Lifted[], ctx: FnCtx): Value {
  const x = firstNumber(args, ctx);
  if (isError(x)) return x;
  let n = Math.trunc(x);
  if (n < 0 || n > 3999) return err('NUM', 'ROMAN supports 0..3999');
  if (n === 0) return '';
  let out = '';
  for (const [v, s] of ROMAN_TABLE) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

function romanVal(ch: string): number {
  switch (ch) {
    case 'I': return 1;
    case 'V': return 5;
    case 'X': return 10;
    case 'L': return 50;
    case 'C': return 100;
    case 'D': return 500;
    case 'M': return 1000;
    default: return 0;
  }
}

/* ------------------------------------------------------------------ matrix */

function numericMatrix(arg: Lifted | undefined, ctx: FnCtx): number[][] | CellError {
  const m = argMatrix(arg, ctx);
  return m.rows.map((row) =>
    row.map((v) => {
      if (isError(v)) return NaN;
      if (typeof v === 'boolean') return v ? 1 : 0;
      return typeof v === 'number' ? v : NaN;
    }),
  );
}

function mmult(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const A = numericMatrix(args[0], ctx);
  if (isError(A)) return A;
  const B = numericMatrix(args[1], ctx);
  if (isError(B)) return B;
  const n = A.length;
  const k = B.length;
  const m = B[0]?.length ?? 0;
  if (n > 0 && A[0]?.length !== k) return err('VALUE', 'MMULT shapes do not line up');
  const out: Value[][] = [];
  for (let i = 0; i < n; i++) {
    const row: Value[] = [];
    for (let j = 0; j < m; j++) {
      let s = 0;
      for (let p = 0; p < k; p++) s += A[i][p] * B[p][j];
      row.push(s);
    }
    out.push(row);
  }
  return makeArray(out);
}

function det(m: number[][]): number | null {
  const n = m.length;
  if (n === 0) return null;
  if (n === 1) return m[0][0];
  if (n === 2) return m[0][0] * m[1][1] - m[0][1] * m[1][0];

  const a = m.map((r) => r.slice());
  let result = 1;
  for (let i = 0; i < n; i++) {
    let pivot = i;
    for (let r = i + 1; r < n; r++) {
      if (Math.abs(a[r][i]) > Math.abs(a[pivot][i])) pivot = r;
    }
    if (Math.abs(a[pivot][i]) < 1e-14) return null;
    if (pivot !== i) {
      const tmp = a[pivot];
      a[pivot] = a[i];
      a[i] = tmp;
      result = -result;
    }
    result *= a[i][i];
    for (let r = i + 1; r < n; r++) {
      const f = a[r][i] / a[i][i];
      for (let c = i; c < n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return result;
}

function minverse(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const m = numericMatrix(args[0], ctx);
  if (isError(m)) return m;
  const n = m.length;
  if (n === 0 || m.some((r) => r.length !== n)) {
    return err('VALUE', 'MINVERSE needs a square matrix');
  }

  // Gauss-Jordan elimination with partial pivoting.
  const a: number[][] = m.map((row, i) => [
    ...row,
    ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)),
  ]);
  for (let i = 0; i < n; i++) {
    let pivot = i;
    for (let r = i + 1; r < n; r++) {
      if (Math.abs(a[r][i]) > Math.abs(a[pivot][i])) pivot = r;
    }
    if (Math.abs(a[pivot][i]) < 1e-14) return err('NUM', 'MINVERSE: matrix is singular');
    const tmp = a[pivot];
    a[pivot] = a[i];
    a[i] = tmp;
    const d = a[i][i];
    for (let c = 0; c < 2 * n; c++) a[i][c] /= d;
    for (let r = 0; r < n; r++) {
      if (r === i) continue;
      const f = a[r][i];
      if (f === 0) continue;
      for (let c = 0; c < 2 * n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return makeArray(a.map((row) => row.slice(n)));
}

/* -------------------------------------------------- conditional aggregates */

/** Pair a criteria range with a target range of matching shape. */
function pairRanges(
  criteria: Lifted,
  target: Lifted | undefined,
  ctx: FnCtx,
): { crit: Value[]; vals: Value[] } {
  const crit = argFlat(criteria, ctx);
  const vals = target === undefined ? crit : argFlat(target, ctx);
  return { crit, vals };
}

function conditionalSum(args: Lifted[], ctx: FnCtx): Value {
  const { crit, vals } = pairRanges(args[0], args[2], ctx);
  const pred = criterionPredicate(args[1] as Value);
  let sum = 0;
  let any = false;
  for (let i = 0; i < crit.length; i++) {
    if (!pred(crit[i])) continue;
    const v = vals[i];
    if (typeof v === 'number') {
      sum += v;
      any = true;
    }
  }
  return any ? sum : 0;
}

/** (target, critRange1, crit1, critRange2, crit2, ...) -> predicates. */
function buildCondPairs(
  args: Lifted[],
  ctx: FnCtx,
): { crit: Value[]; pred: Predicate }[] {
  const pairs: { crit: Value[]; pred: Predicate }[] = [];
  for (let i = 1; i + 1 < args.length; i += 2) {
    pairs.push({ crit: argFlat(args[i], ctx), pred: criterionPredicate(args[i + 1] as Value) });
  }
  return pairs;
}

function matchesAll(pairs: { crit: Value[]; pred: Predicate }[], i: number): boolean {
  for (const p of pairs) {
    if (i >= p.crit.length) return false;
    if (!p.pred(p.crit[i])) return false;
  }
  return true;
}

function conditionalSumMulti(args: Lifted[], ctx: FnCtx): Value {
  const vals = argFlat(args[0], ctx);
  const pairs = buildCondPairs(args, ctx);
  let sum = 0;
  let any = false;
  for (let i = 0; i < vals.length; i++) {
    if (!matchesAll(pairs, i)) continue;
    const v = vals[i];
    if (typeof v === 'number') {
      sum += v;
      any = true;
    }
  }
  return any ? sum : 0;
}

function conditionalCount(args: Lifted[], ctx: FnCtx): Value {
  if (args.length === 2) {
    const crit = argFlat(args[0], ctx);
    const pred = criterionPredicate(args[1] as Value);
    return crit.filter((v) => v !== null && pred(v)).length;
  }
  const pairs = buildCondPairs(args, ctx);
  const length = pairs.length ? Math.min(...pairs.map((p) => p.crit.length)) : 0;
  let n = 0;
  for (let i = 0; i < length; i++) if (matchesAll(pairs, i)) n++;
  return n;
}

function conditionalAverage(args: Lifted[], ctx: FnCtx): Value {
  const { crit, vals } = pairRanges(args[0], args[2], ctx);
  const pred = criterionPredicate(args[1] as Value);
  const picked: number[] = [];
  for (let i = 0; i < crit.length; i++) {
    if (pred(crit[i])) {
      const v = vals[i];
      if (typeof v === 'number') picked.push(v);
    }
  }
  if (!picked.length) return err('DIV/0', 'AVERAGEIF found no matching numbers');
  return picked.reduce((s, v) => s + v, 0) / picked.length;
}

function conditionalAverageMulti(args: Lifted[], ctx: FnCtx): Value {
  const vals = argFlat(args[0], ctx);
  const pairs = buildCondPairs(args, ctx);
  const picked: number[] = [];
  for (let i = 0; i < vals.length; i++) {
    if (matchesAll(pairs, i)) {
      const v = vals[i];
      if (typeof v === 'number') picked.push(v);
    }
  }
  if (!picked.length) return err('DIV/0', 'AVERAGEIFS found no matching numbers');
  return picked.reduce((s, v) => s + v, 0) / picked.length;
}

function conditionalExtremum(args: Lifted[], ctx: FnCtx, wantMax: boolean): Value {
  const vals = argFlat(args[0], ctx);
  const pairs = buildCondPairs(args, ctx);
  const picked: number[] = [];
  for (let i = 0; i < vals.length; i++) {
    if (matchesAll(pairs, i)) {
      const v = vals[i];
      if (typeof v === 'number') picked.push(v);
    }
  }
  if (!picked.length) return 0;
  return wantMax ? Math.max(...picked) : Math.min(...picked);
}

/** SUMPRODUCT: multiply corresponding values, coercing non-numerics to 0. */
function sumproduct(args: Lifted[], ctx: FnCtx): Value {
  const arrays = args.map((a) => {
    if (isArray(a)) return (a as ArrayRef).rows.flat();
    if (isRange(a)) return ctx.values(a);
    return [a as Value];
  });
  const len = Math.min(...arrays.map((x) => x.length));
  let total = 0;
  for (let i = 0; i < len; i++) {
    let p = 1;
    for (const arr of arrays) {
      const v = arr[i];
      p *= typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
    }
    total += p;
  }
  return total;
}

function paired(args: Lifted[], ctx: FnCtx, f: (x: number, y: number) => number): Value {
  const xs = numbersOnly([args[0]], ctx);
  const ys = numbersOnly([args[1]], ctx);
  if (xs.length !== ys.length) return err('N/A', 'Ranges must be the same size');
  let s = 0;
  for (let i = 0; i < xs.length; i++) s += f(xs[i], ys[i]);
  return s;
}

/* -------------------------------------------------- SUBTOTAL / AGGREGATE */

/** SUBTOTAL dispatches on Excel's function-number codes. */
function subtotalImpl(args: Lifted[], ctx: FnCtx): Value {
  const code = firstNumber(args, ctx);
  if (isError(code)) return code;
  const fnName = SUBTOTAL_FN[Math.trunc(code)];
  if (!fnName) return err('VALUE', `SUBTOTAL code ${code} is not supported`);
  return applyAggregate(fnName, args.slice(1), ctx);
}

const AGGREGATE_FN: Record<number, string> = {
  1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT',
  7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P',
  12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL',
  16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC',
};

function aggregateImpl(args: Lifted[], ctx: FnCtx): Value {
  const code = firstNumber(args, ctx);
  if (isError(code)) return code;
  const n = Math.trunc(code);
  const fnName = AGGREGATE_FN[n];
  if (!fnName) return err('VALUE', `AGGREGATE code ${n} is not supported`);
  const tail = args.slice(1);
  // Codes 1-5 take no options argument and skip error cells; with an options
  // value present, option 0 skips errors and 6-7 propagate them.
  const optArg = intArg(args, ctx, 1, 0);
  const ignoreErrors = n <= 5 ? true : isError(optArg) || Math.trunc(optArg) === 0;
  if (ignoreErrors) {
    return applyAggregate(fnName, tail.map((a) => stripErrors(a, ctx)), ctx);
  }
  return applyAggregate(fnName, tail, ctx);
}

function stripErrors(arg: Lifted, ctx: FnCtx): Lifted {
  const drop = (row: Value[]): Value[] => row.map((v) => (isError(v) ? null : v));
  if (isArray(arg)) return makeArray((arg as ArrayRef).rows.map(drop));
  if (isRange(arg)) return makeArray(ctx.rowsOf(arg).map(drop));
  return isError(arg) ? null : (arg as Value);
}

/** Shared by SUBTOTAL, AGGREGATE and the STATS family. */
function applyAggregate(name: string, args: Lifted[], ctx: FnCtx): Value {
  const nums = numbersOnly(args, ctx);
  switch (name) {
    case 'SUM':
      return nums.reduce((a, b) => a + b, 0);
    case 'AVERAGE':
      return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : err('DIV/0');
    case 'COUNT':
      return nums.length;
    case 'COUNTA': {
      let counted = 0;
      for (const arg of args) {
        for (const v of flat(arg, ctx)) if (v !== null && v !== '') counted++;
      }
      return counted;
    }
    case 'MAX':
      return nums.length ? Math.max(...nums) : 0;
    case 'MIN':
      return nums.length ? Math.min(...nums) : 0;
    case 'PRODUCT':
      return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'MEDIAN': {
      if (!nums.length) return err('NUM');
      const s = [...nums].sort((p, q) => p - q);
      const m = s.length >> 1;
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case 'STDEV.S':
    case 'STDEV.P':
    case 'STDEVP': {
      const sample = name !== 'STDEV.P' && name !== 'STDEVP';
      if (nums.length < (sample ? 2 : 1)) return err('DIV/0');
      const m = nums.reduce((a, b) => a + b, 0) / nums.length;
      const ss = nums.reduce((a, b) => a + (b - m) ** 2, 0);
      return Math.sqrt(ss / (sample ? nums.length - 1 : nums.length));
    }
    case 'VAR.S':
    case 'VAR.P':
    case 'VARP': {
      const sample = name !== 'VAR.P' && name !== 'VARP';
      if (nums.length < (sample ? 2 : 1)) return err('DIV/0');
      const m = nums.reduce((a, b) => a + b, 0) / nums.length;
      const ss = nums.reduce((a, b) => a + (b - m) ** 2, 0);
      return ss / (sample ? nums.length - 1 : nums.length);
    }
    case 'LARGE':
    case 'SMALL': {
      const k = intArg(args, ctx, 1, 0);
      if (isError(k)) return k;
      if (k < 1 || k > nums.length) return err('NUM');
      const s = [...nums].sort((a, b) => (name === 'LARGE' ? b - a : a - b));
      return s[k - 1];
    }
    case 'PERCENTILE.INC':
    case 'PERCENTILE.EXC':
    case 'QUARTILE.INC':
    case 'QUARTILE.EXC': {
      let p = intArg(args, ctx, 1, 0);
      if (isError(p)) return p;
      const isQuartile = name.startsWith('QUARTILE');
      const q = toNumber(argFlat(args[1], ctx)[0] ?? 1);
      if (isError(q)) return q;
      p = isQuartile ? q / 4 : q;
      if (!nums.length) return err('NUM');
      if (p < 0 || p > 1) return err('NUM');
      const s = [...nums].sort((a, b) => a - b);
      if (name.endsWith('.INC')) return percentileInc(s, p);
      const k = p * (s.length + 1);
      if (k < 1 || k > s.length) return err('NUM');
      const lo = Math.floor(k);
      const hi = Math.ceil(k);
      return lo === hi ? s[lo - 1] : s[lo - 1] + (s[hi - 1] - s[lo - 1]) * (k - lo);
    }
    case 'MODE.SNGL': {
      const counts = new Map<number, number>();
      for (const n of nums) counts.set(n, (counts.get(n) ?? 0) + 1);
      let best: number | null = null;
      let bestN = 1;
      for (const [v, c] of counts) {
        if (c > bestN) {
          best = v;
          bestN = c;
        }
      }
      return best === null ? err('N/A') : best;
    }
    default:
      return err('VALUE', `${name} is not implemented`);
  }
}

function percentileInc(sorted: number[], p: number): number {
  const n = sorted.length;
  if (n === 1) return sorted[0];
  const idx = p * (n - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}
