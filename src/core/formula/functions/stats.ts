/**
 * Descriptive statistics, distributions and regression.
 *
 * All functions ignore text and blanks inside ranges (Excel semantics) and
 * propagate errors found in the numeric inputs.
 */

import { err, isError, type CellError, type Value } from '../../types';
import { toNumber } from '../../coerce';
import {
  flat,
  makeArray,
  numbersOnly,
  type ArrayRef,
  type FnCtx,
  type FnDef,
  type Lifted,
} from '../fntypes';
import { firstError, numOut } from '../helpers';

/* ------------------------------------------------------------- collections */

function nums(args: Lifted[], ctx: FnCtx): number[] | Value {
  const e = firstError(args, ctx);
  if (e) return e;
  return numbersOnly(args, ctx);
}

function ok(v: number[] | Value): v is Value {
  return isError(v);
}

function mean(xs: number[]): number {
  return xs.reduce((a, b) => a + b, 0) / xs.length;
}

function devSq(xs: number[]): number {
  const m = mean(xs);
  return xs.reduce((a, b) => a + (b - m) ** 2, 0);
}

/** Sample variance (n-1) unless `population` is true. */
function variance(xs: number[], population = false): number {
  if (xs.length < (population ? 1 : 2)) return NaN;
  return devSq(xs) / (population ? xs.length : xs.length - 1);
}

function quantile(sorted: number[], p: number): number {
  if (!sorted.length) return NaN;
  if (sorted.length === 1) return sorted[0];
  const idx = (sorted.length - 1) * p;
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  return lo === hi ? sorted[lo] : sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}

/** Inverse of a strictly monotonic CDF via bisection. Excel uses this shape. */
function invertCdf(cdf: (x: number) => number, target: number, lo = -40, hi = 40): number {
  let a = lo;
  let b = hi;
  for (let i = 0; i < 200; i++) {
    const mid = (a + b) / 2;
    if (cdf(mid) < target) a = mid;
    else b = mid;
  }
  return (a + b) / 2;
}

/* ------------------------------------------------------------ special funcs */

/** Abramowitz & Stegun 7.1.26 error function. */
function erf(x: number): number {
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  const t = 1 / (1 + 0.3275911 * ax);
  const y =
    1 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-ax * ax);
  return sign * y;
}

function erfInv(p: number): number {
  return invertCdf((x) => 0.5 * (1 + erf(x / Math.SQRT2)), p, -8, 8);
}

/** Standard normal CDF (mean 0, sd 1). */
function normCdf(x: number, mean = 0, sd = 1): number {
  return 0.5 * (1 + erf((x - mean) / (sd * Math.SQRT2)));
}

/** Standard normal PDF. */
function normPdf(x: number, mean = 0, sd = 1): number {
  const z = (x - mean) / sd;
  return Math.exp(-0.5 * z * z) / (sd * Math.sqrt(2 * Math.PI));
}

function gammaFn(z: number): number {
  const g = 7;
  const c = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028,
    771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (z < 0.5) return Math.PI / (Math.sin(Math.PI * z) * gammaFn(1 - z));
  const zz = z - 1;
  let x = c[0];
  for (let i = 1; i < g + 2; i++) x += c[i] / (zz + i);
  const t = zz + g + 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, zz + 0.5) * Math.exp(-t) * x;
}

/** Natural log of the gamma function; stable for large arguments. */
function lnGamma(z: number): number {
  return Math.log(gammaFn(z));
}

function lowerIncompleteGamma(s: number, x: number): number {
  if (x < 0) return NaN;
  if (x === 0) return 0;
  if (x < s + 1) {
    // series expansion
    let sum = 1 / s;
    let term = sum;
    for (let n = 1; n < 500; n++) {
      term *= x / (s + n);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * 1e-15) break;
    }
    return sum * Math.exp(-x + s * Math.log(x) - lnGamma(s));
  }
  // continued fraction (Lentz)
  const tiny = 1e-300;
  let b = x + 1 - s;
  let c = 1 / tiny;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 500; i++) {
    const an = -i * (i - s);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-15) break;
  }
  const q = Math.exp(-x + s * Math.log(x) - lnGamma(s));
  return 1 - q * h;
}

function gammaincP(s: number, x: number): number {
  if (x <= 0) return 0;
  if (x < s + 1) return lowerIncompleteGamma(s, x);
  return 1 - upperTail(s, x);
}

function upperTail(s: number, x: number): number {
  // P(s,x) complement via continued fraction
  const tiny = 1e-300;
  let b = x + 1 - s;
  let c = 1 / tiny;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 500; i++) {
    const an = -i * (i - s);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-15) break;
  }
  return Math.exp(-x + s * Math.log(x) - lnGamma(s)) * h;
}

function incompleteBeta(a: number, b: number, x: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const lbeta = lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x);
  const front = Math.exp(lbeta);
  if (x < (a + 1) / (a + b + 2)) {
    let sum = 1 / a;
    let term = sum;
    for (let n = 1; n < 500; n++) {
      term *= ((n * (b - n) * x) / ((a + 2 * n - 1) * (a + 2 * n)));
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * 1e-15) break;
    }
    return front * sum;
  }
  return 1 - incompleteBeta(b, a, 1 - x);
}

/** Student-t CDF via the incomplete beta function. */
function tCdf(t: number, df: number): number {
  const x = df / (df + t * t);
  const p = 0.5 * incompleteBeta(df / 2, 0.5, x);
  return t > 0 ? 1 - p : p;
}

/* ---------------------------------------------------------------- library */

export const statsFns: FnDef[] = [
  {
    name: 'AVERAGE', min: 1,
    fn: (a, ctx) => {
      const xs = nums(a, ctx);
      if (ok(xs)) return xs;
      if (!xs.length) return err('DIV/0', 'AVERAGE of an empty set');
      return numOut(mean(xs));
    },
  },
  { name: 'AVERAGEA', min: 1, fn: (a, ctx) => {
    const vals = collect(a, ctx);
    if (!vals.length) return err('DIV/0');
    return numOut(collectNums(a, ctx).reduce((s, v) => s + v, 0) / vals.length);
  } },
  { name: 'COUNT', min: 1, fn: (a, ctx) => numbersOnly(a, ctx).length },
  { name: 'COUNTA', min: 1, fn: (a, ctx) => {
    let n = 0;
    for (const arg of a) {
      for (const v of flat(arg, ctx)) if (v !== null && v !== '') n++;
    }
    return n;
  } },
  { name: 'COUNTBLANK', min: 1, fn: (a, ctx) => {
    let n = 0;
    for (const arg of a) {
      for (const v of flat(arg, ctx)) if (v === null || v === '') n++;
    }
    return n;
  } },
  { name: 'MAX', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); return ok(xs) ? xs : (xs.length ? Math.max(...xs) : 0); } },
  { name: 'MIN', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); return ok(xs) ? xs : (xs.length ? Math.min(...xs) : 0); } },
  { name: 'MAXA', min: 1, fn: (a, ctx) => extremeA(a, ctx, true) },
  { name: 'MINA', min: 1, fn: (a, ctx) => extremeA(a, ctx, false) },
  { name: 'MEDIAN', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; if (!xs.length) return err('NUM'); const s = [...xs].sort((p, q) => p - q); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; } },
  { name: 'MODE', min: 1, fn: (a, ctx) => modeOf(a, ctx) },
  { name: 'MODE.SNGL', min: 1, fn: (a, ctx) => modeOf(a, ctx) },
  { name: 'MODE.MULT', min: 1, fn: (a, ctx) => modeMulti(a, ctx) },

  { name: 'STDEV', min: 1, fn: (a, ctx) => spread(a, ctx, true, false) },
  { name: 'STDEV.S', min: 1, fn: (a, ctx) => spread(a, ctx, true, false) },
  { name: 'STDEVA', min: 1, fn: (a, ctx) => spreadA(a, ctx, true) },
  { name: 'STDEV.P', min: 1, fn: (a, ctx) => spread(a, ctx, false, false) },
  { name: 'STDEVP', min: 1, fn: (a, ctx) => spread(a, ctx, false, false) },
  { name: 'VAR', min: 1, fn: (a, ctx) => varianceOf(a, ctx, true) },
  { name: 'VAR.S', min: 1, fn: (a, ctx) => varianceOf(a, ctx, true) },
  { name: 'VARA', min: 1, fn: (a, ctx) => varianceA(a, ctx, true) },
  { name: 'VAR.P', min: 1, fn: (a, ctx) => varianceOf(a, ctx, false) },
  { name: 'VARP', min: 1, fn: (a, ctx) => varianceOf(a, ctx, false) },

  { name: 'DEVSQ', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); return ok(xs) ? xs : devSq(xs); } },
  { name: 'AVEDEV', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; if (!xs.length) return err('NUM'); const m = mean(xs); return numOut(xs.reduce((s, x) => s + Math.abs(x - m), 0) / xs.length); } },
  { name: 'GEOMEAN', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; if (!xs.length || xs.some((x) => x <= 0)) return err('NUM', 'GEOMEAN needs positive numbers'); return numOut(Math.exp(xs.reduce((s, x) => s + Math.log(x), 0) / xs.length)); } },
  { name: 'HARMEAN', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; if (!xs.length || xs.some((x) => x <= 0)) return err('NUM'); return numOut(xs.length / xs.reduce((s, x) => s + 1 / x, 0)); } },
  { name: 'SKEW', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; const n = xs.length; if (n < 3) return err('DIV/0', 'SKEW needs 3 values'); const m = mean(xs); const s = Math.sqrt(devSq(xs) / (n - 1)); return numOut((n / ((n - 1) * (n - 2))) * xs.reduce((t, x) => t + Math.pow((x - m) / s, 3), 0)); } },
  { name: 'SKEW.P', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; const n = xs.length; if (n < 2) return err('DIV/0'); const m = mean(xs); const s = Math.sqrt(devSq(xs) / n); return numOut(xs.reduce((t, x) => t + Math.pow((x - m) / s, 3), 0) / n); } },
  { name: 'KURT', min: 1, fn: (a, ctx) => { const xs = nums(a, ctx); if (ok(xs)) return xs; const n = xs.length; if (n < 4) return err('DIV/0', 'KURT needs 4 values'); const m = mean(xs); const s = Math.sqrt(devSq(xs) / (n - 1)); const t = xs.reduce((acc, x) => acc + Math.pow((x - m) / s, 4), 0); return numOut((n * (n + 1) / ((n - 1) * (n - 2) * (n - 3))) * t - (3 * (n - 1) ** 2 / ((n - 2) * (n - 3)))); } },

  { name: 'LARGE', min: 2, max: 2, fn: (a, ctx) => nth(a, ctx, 'large') },
  { name: 'SMALL', min: 2, max: 2, fn: (a, ctx) => nth(a, ctx, 'small') },
  { name: 'RANK', min: 2, max: 3, fn: (a, ctx) => rankImpl(a, ctx, false) },
  { name: 'RANK.EQ', min: 2, max: 3, fn: (a, ctx) => rankImpl(a, ctx, false) },
  { name: 'RANK.AVG', min: 2, max: 3, fn: (a, ctx) => rankImpl(a, ctx, true) },
  { name: 'PERCENTILE', min: 2, max: 2, fn: (a, ctx) => pctImpl(a, ctx, true) },
  { name: 'PERCENTILE.INC', min: 2, max: 2, fn: (a, ctx) => pctImpl(a, ctx, true) },
  { name: 'PERCENTILE.EXC', min: 2, max: 2, fn: (a, ctx) => pctImpl(a, ctx, false) },
  { name: 'QUARTILE', min: 2, max: 2, fn: (a, ctx) => quartileImpl(a, ctx) },
  { name: 'QUARTILE.INC', min: 2, max: 2, fn: (a, ctx) => quartileImpl(a, ctx) },
  { name: 'QUARTILE.EXC', min: 2, max: 2, fn: (a, ctx) => quartileImpl(a, ctx, true) },
  { name: 'PERCENTRANK', min: 2, max: 3, fn: (a, ctx) => percentRank(a, ctx, true) },
  { name: 'PERCENTRANK.INC', min: 2, max: 3, fn: (a, ctx) => percentRank(a, ctx, true) },
  { name: 'PERCENTRANK.EXC', min: 2, max: 3, fn: (a, ctx) => percentRank(a, ctx, false) },

  /* --- regression / correlation --- */
  { name: 'CORREL', min: 2, max: 2, fn: (a, ctx) => correl(a, ctx) },
  { name: 'PEARSON', min: 2, max: 2, fn: (a, ctx) => correl(a, ctx) },
  { name: 'RSQ', min: 2, max: 2, fn: (a, ctx) => { const r = correl(a, ctx); return isError(r) ? r : r * r; } },
  { name: 'COVAR', min: 2, max: 2, fn: (a, ctx) => { const p = pairedNums(a, ctx); if (isError(p)) return p; return numOut(devSq(p.xs) + devSq(p.ys) - devSq(p.xs.map((x, i) => x + p.ys[i]))); } },
  { name: 'COVARIANCE.P', min: 2, max: 2, fn: (a, ctx) => { const p = pairedNums(a, ctx); if (isError(p)) return p; return numOut(devSq(p.xs.map((x, i) => x + p.ys[i])) / 2 / ((p.xs.length) * (p.xs.length - 1))); } },
  { name: 'SLOPE', min: 2, max: 2, fn: (a, ctx) => { const p = pairedNums(a, ctx); if (isError(p)) return p;
    // Excel passes (known_y, known_x), so the slope is d(y)/d(x).
    const xs = p.ys;
    const ys = p.xs;
    const den = devSq(xs);
    if (den === 0) return err('DIV/0');
    return numOut(xs.reduce((s, x, i) => s + (x - mean(xs)) * (ys[i] - mean(ys)), 0) / den);
  } },
  { name: 'INTERCEPT', min: 2, max: 2, fn: (a, ctx) => { const p = pairedNums(a, ctx); if (isError(p)) return p;
    const xs = p.ys;
    const ys = p.xs;
    const den = devSq(xs);
    if (den === 0) return err('DIV/0');
    const m = xs.reduce((s, x, i) => s + (x - mean(xs)) * (ys[i] - mean(ys)), 0) / den;
    return numOut(mean(ys) - m * mean(xs));
  } },
  { name: 'FORECAST', min: 3, max: 3, fn: (a, ctx) => forecast(a, ctx) },
  { name: 'FORECAST.LINEAR', min: 3, max: 3, fn: (a, ctx) => forecast(a, ctx) },
  { name: 'TREND', min: 2, max: 4, fn: (a, ctx) => trend(a, ctx) },
  { name: 'GROWTH', min: 2, max: 4, fn: (a, ctx) => growth(a, ctx) },
  { name: 'LINEST', min: 1, max: 4, fn: (a, ctx) => linest(a, ctx) },

  /* --- distributions --- */
  { name: 'NORM.DIST', min: 4, max: 4, fn: (a, ctx) => normDist(a, ctx) },
  { name: 'NORMDIST', min: 4, max: 4, fn: (a, ctx) => normDist(a, ctx) },
  { name: 'NORM.INV', min: 3, max: 3, fn: (a, ctx) => normInv(a, ctx) },
  { name: 'NORMINV', min: 3, max: 3, fn: (a, ctx) => normInv(a, ctx) },
  { name: 'NORM.S.DIST', min: 1, max: 2, fn: (a, ctx) => { const z = numArg(a, ctx, 0); if (isError(z)) return z; const cum = a.length > 1 ? truthyArg(a[1]) : true; return numOut(normCdf(z, 0, 1)); } },
  { name: 'NORMSDIST', min: 1, max: 1, fn: (a, ctx) => { const z = numArg(a, ctx, 0); return isError(z) ? z : numOut(normCdf(z, 0, 1)); } },
  { name: 'NORM.S.INV', min: 1, max: 1, fn: (a, ctx) => { const p = numArg(a, ctx, 0); if (isError(p)) return p; if (p <= 0 || p >= 1) return err('NUM'); return numOut(erfInv(p)); } },
  { name: 'NORMSINV', min: 1, max: 1, fn: (a, ctx) => { const p = numArg(a, ctx, 0); if (isError(p)) return p; if (p <= 0 || p >= 1) return err('NUM'); return numOut(erfInv(p)); } },
  { name: 'NORM.S.DIST.Z', min: 2, max: 2, fn: (a, ctx) => { const z = numArg(a, ctx, 0); return isError(z) ? z : numOut(normCdf(z, 0, 1)); } },

  { name: 'T.DIST', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const df = numArg(a, ctx, 1); if (isError(df)) return df; const cum = truthyArg(a[2]); return df < 1 ? err('NUM') : numOut(cum ? tCdf(x, df) : tCdf(x, df) - tCdf(x, -df)); } },
  { name: 'T.DIST.2T', min: 2, max: 2, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const df = numArg(a, ctx, 1); if (isError(df)) return df; if (df < 1 || x < 0) return err('NUM'); return numOut(2 * (1 - tCdf(Math.abs(x), df))); } },
  { name: 'T.DIST.RT', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const df = numArg(a, ctx, 1); if (isError(df)) return df; return numOut(1 - tCdf(x, df)); } },
  { name: 'TDIST', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const df = numArg(a, ctx, 1); if (isError(df)) return df; const two = truthyArg(a[2]); const p = two ? 2 * (1 - tCdf(Math.abs(x), df)) : 1 - tCdf(x, df); return numOut(p); } },
  { name: 'T.INV', min: 2, max: 2, fn: (a, ctx) => { const p = numArg(a, ctx, 0); if (isError(p)) return p; const df = numArg(a, ctx, 1); if (isError(df)) return df; if (p <= 0 || p >= 1 || df < 1) return err('NUM'); return numOut(invertCdf((x) => tCdf(x, df), p)); } },
  { name: 'T.INV.2T', min: 2, max: 2, fn: (a, ctx) => { const p = numArg(a, ctx, 0); if (isError(p)) return p; const df = numArg(a, ctx, 1); if (isError(df)) return df; if (p <= 0 || p > 1 || df < 1) return err('NUM'); const alpha = p / 2; return numOut(invertCdf((x) => tCdf(x, df), 1 - alpha, -1e4, 1e4)); } },

  { name: 'CHISQ.DIST', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const df = numArg(a, ctx, 1); if (isError(df)) return df; const cum = truthyArg(a[2]); const v = df / 2; if (!cum) return numOut(Math.pow(x, v - 1) * Math.exp(-x / 2) / (Math.pow(2, v) * gammaFn(v))); return numOut(gammaincP(v, x / 2)); } },
  { name: 'CHISQ.DIST.RT', min: 2, max: 2, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const df = numArg(a, ctx, 1); if (isError(df)) return df; return numOut(upperTail(df / 2, x / 2)); } },
  { name: 'CHISQ.INV', min: 2, max: 2, fn: (a, ctx) => { const p = numArg(a, ctx, 0); if (isError(p)) return p; const df = numArg(a, ctx, 1); if (isError(df)) return df; return numOut(2 * invertCdf((x) => gammaincP(df / 2, x), p, 0, df * 20 + 100)); } },
  { name: 'CHISQ.INV.RT', min: 2, max: 2, fn: (a, ctx) => { const p = numArg(a, ctx, 0); if (isError(p)) return p; const df = numArg(a, ctx, 1); if (isError(df)) return df; return numOut(2 * invertCdf((x) => upperTail(df / 2, x), p, 0, df * 20 + 100)); } },

  { name: 'BINOM.DIST', min: 4, max: 4, fn: (a, ctx) => binom(a, ctx) },
  { name: 'BINOMDIST', min: 4, max: 4, fn: (a, ctx) => binom(a, ctx) },
  { name: 'POISSON.DIST', min: 3, max: 3, fn: (a, ctx) => poisson(a, ctx) },
  { name: 'POISSON', min: 2, max: 2, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const meanV = numArg(a, ctx, 1); if (isError(meanV)) return meanV; if (x < 0 || x !== Math.trunc(x)) return err('NUM'); return numOut(Math.exp(-meanV) * Math.pow(meanV, x) / gammaFn(x + 1)); } },
  { name: 'EXPON.DIST', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const lambda = numArg(a, ctx, 1); if (isError(lambda)) return lambda; const cum = truthyArg(a[2]); if (x < 0 || lambda <= 0) return err('NUM'); return numOut(cum ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x)); } },
  { name: 'EXPONDIST', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const lambda = numArg(a, ctx, 1); if (isError(lambda)) return lambda; const cum = truthyArg(a[2]); if (x < 0 || lambda <= 0) return err('NUM'); return numOut(cum ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x)); } },

  { name: 'F.DIST.RT', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const d1 = numArg(a, ctx, 1); if (isError(d1)) return d1; const d2 = numArg(a, ctx, 2); if (isError(d2)) return d2; const p = incompleteBeta(d1 / 2, d2 / 2, (d1 * x) / (d1 * x + d2)); return numOut(1 - p); } },

  { name: 'Z.TEST', min: 2, max: 3, fn: (a, ctx) => ztest(a, ctx) },
  { name: 'ZTEST', min: 2, max: 3, fn: (a, ctx) => ztest(a, ctx) },
  { name: 'T.TEST', min: 4, max: 4, fn: (a, ctx) => ttest(a, ctx) },
  { name: 'TTEST', min: 4, max: 4, fn: (a, ctx) => ttest(a, ctx) },
  { name: 'CHISQ.TEST', min: 2, max: 2, fn: (a, ctx) => chisqTest(a, ctx) },
  { name: 'F.TEST', min: 2, max: 2, fn: (a, ctx) => { const p = pairedNums(a, ctx); if (isError(p)) return p; const v1 = variance(p.ys, true); const v2 = variance(p.xs, true); if (v2 === 0) return err('DIV/0'); return numOut(v1 / v2); } },

  { name: 'CONFIDENCE.NORM', min: 3, max: 3, fn: (a, ctx) => confidence(a, ctx, true) },
  { name: 'CONFIDENCE', min: 3, max: 3, fn: (a, ctx) => confidence(a, ctx, true) },
  { name: 'CONFIDENCE.T', min: 3, max: 3, fn: (a, ctx) => confidence(a, ctx, false) },
  { name: 'STANDARDIZE', min: 3, max: 3, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; const m = numArg(a, ctx, 1); if (isError(m)) return m; const s = numArg(a, ctx, 2); if (isError(s)) return s; if (s <= 0) return err('NUM'); return numOut((x - m) / s); } },
  { name: 'FISHER', min: 1, max: 1, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; if (x <= -1 || x >= 1) return err('NUM', 'FISHER needs -1 < x < 1'); return numOut(Math.atanh(x)); } },
  { name: 'FISHERINV', min: 1, max: 1, fn: (a, ctx) => { const x = numArg(a, ctx, 0); if (isError(x)) return x; if (x <= -1 || x >= 1) return err('NUM'); return numOut(Math.tanh(x)); } },
  { name: 'PROB', min: 3, max: 4, fn: prob },

  { name: 'FREQUENCY', min: 2, max: 2, fn: frequency },
  { name: 'QUARTILE.EXC.A', min: 2, max: 2, fn: (a, ctx) => quartileImpl(a, ctx, true) },
];

/* ------------------------------------------------------------------ pieces */

function isRangeLike(v: Lifted): boolean {
  return typeof v === 'object' && v !== null && '__range' in v;
}

function collect(args: Lifted[], ctx: FnCtx): (number | boolean)[] {
  const out: (number | boolean)[] = [];
  for (const arg of args) {
    const vals = isRangeLike(arg) ? ctx.values(arg as never) : [arg as Value];
    for (const v of vals) {
      if (v === null || v === '') continue;
      if (typeof v === 'number') out.push(v);
      else if (typeof v === 'boolean') out.push(v);
    }
  }
  return out;
}

/** Collect as plain numbers, treating booleans as 1/0. */
function collectNums(args: Lifted[], ctx: FnCtx): number[] {
  return collect(args, ctx).map((v) => (typeof v === 'boolean' ? (v ? 1 : 0) : v));
}

function numArg(args: Lifted[], ctx: FnCtx, i: number): number | CellError {
  const v = args[i];
  if (v === undefined) return err('VALUE', 'Missing argument');
  const s = isRangeLike(v) ? ctx.values(v as never)[0] : (v as Value);
  if (isError(s)) return s;
  return toNumber(s);
}

function truthyArg(v: Lifted | undefined): boolean {
  if (v === undefined) return true;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  return true;
}

function extremeA(args: Lifted[], ctx: FnCtx, wantMax: boolean): Value {
  const nums = collectNums(args, ctx);
  if (!nums.length) return 0;
  return wantMax ? Math.max(...nums) : Math.min(...nums);
}

function spread(args: Lifted[], ctx: FnCtx, sample: boolean, _pop: boolean): Value {
  const xs = nums(args, ctx);
  if (ok(xs)) return xs;
  const v = variance(xs, !sample);
  if (!Number.isFinite(v)) return err('DIV/0', 'STDEV needs at least 2 values');
  return numOut(Math.sqrt(v));
}

function spreadA(args: Lifted[], ctx: FnCtx, sample: boolean): Value {
  const vals = collect(args, ctx);
  const xs = vals.map((v) => (typeof v === 'boolean' ? (v ? 1 : 0) : v));
  const v = variance(xs, !sample);
  if (!Number.isFinite(v)) return err('DIV/0');
  return numOut(Math.sqrt(v));
}

function varianceOf(args: Lifted[], ctx: FnCtx, sample: boolean): Value {
  const xs = nums(args, ctx);
  if (ok(xs)) return xs;
  const v = variance(xs, !sample);
  if (!Number.isFinite(v)) return err('DIV/0', 'VAR needs at least 2 values');
  return numOut(v);
}

function varianceA(args: Lifted[], ctx: FnCtx, sample: boolean): Value {
  const xs = collectNums(args, ctx);
  const v = variance(xs, !sample);
  if (!Number.isFinite(v)) return err('DIV/0');
  return numOut(v);
}

function modeOf(args: Lifted[], ctx: FnCtx): Value {
  const xs = nums(args, ctx);
  if (ok(xs)) return xs;
  if (!xs.length) return err('N/A', 'MODE needs numbers');
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best: number | null = null;
  let bestN = 1;
  for (const [v, c] of counts) if (c > bestN) { best = v; bestN = c; }
  return best === null ? err('N/A', 'MODE found no repeated value') : best;
}

function modeMulti(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const xs = nums(args, ctx);
  if (ok(xs)) return xs;
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  let bestN = 1;
  for (const c of counts.values()) bestN = Math.max(bestN, c);
  const rows = [[...counts.entries()].filter(([, c]) => c === bestN).map(([v]) => v)];
  return makeArray(rows);
}

function nth(args: Lifted[], ctx: FnCtx, dir: 'large' | 'small'): Value {
  const xs = nums(args, ctx);
  if (ok(xs)) return xs;
  const k = numArg(args, ctx, 1);
  if (isError(k)) return k;
  const n = Math.trunc(k);
  if (n < 1 || n > xs.length) return err('NUM', `${dir.toUpperCase()} needs 1 <= k <= ${xs.length}`);
  const s = [...xs].sort((p, q) => (dir === 'large' ? q - p : p - q));
  return s[n - 1];
}

function rankImpl(args: Lifted[], ctx: FnCtx, average: boolean): Value {
  const x = numArg(args, ctx, 0);
  if (isError(x)) return x;
  const xs = numbersOnly([args[1]], ctx);
  if (!xs.length) return err('N/A');
  const desc = args.length < 3 ? false : truthyArg(args[2]);
  const s = desc ? [...xs].sort((p, q) => q - p) : [...xs].sort((p, q) => p - q);
  const hits: number[] = [];
  s.forEach((v, i) => {
    if (v === x) hits.push(i + 1);
  });
  if (!hits.length) return err('N/A', 'RANK value not found');
  if (!average) return hits[0];
  return (Math.min(...hits) + Math.max(...hits)) / 2;
}

function pctImpl(args: Lifted[], ctx: FnCtx, inclusive: boolean): Value {
  const xs = numbersOnly([args[0]], ctx);
  if (!xs.length) return err('NUM');
  const p = numArg(args, ctx, 1);
  if (isError(p)) return p;
  if (p < 0 || p > 1) return err('NUM', 'Percentile must be between 0 and 1');
  const s = [...xs].sort((a, b) => a - b);
  if (inclusive) return numOut(quantile(s, p));
  // EXC variant: p must lie strictly inside (0, 1) with the k/(n+1) grid.
  const k = p * (s.length + 1);
  if (k < 1 || k > s.length) return err('NUM');
  const lo = Math.floor(k);
  const hi = Math.ceil(k);
  if (lo === hi) return s[lo - 1];
  return numOut(s[lo - 1] + (s[hi - 1] - s[lo - 1]) * (k - lo));
}

function quartileImpl(args: Lifted[], ctx: FnCtx, exc = false): Value {
  const xs = numbersOnly([args[0]], ctx);
  if (!xs.length) return err('NUM');
  const q = numArg(args, ctx, 1);
  if (isError(q)) return q;
  if (q < 0 || q > 4) return err('NUM');
  const p = q / 4;
  const s = [...xs].sort((a, b) => a - b);
  if (!exc) return numOut(quantile(s, p));
  const k = p * (s.length + 1);
  if (k < 1 || k > s.length) return err('NUM');
  const lo = Math.floor(k);
  const hi = Math.ceil(k);
  return lo === hi ? s[lo - 1] : numOut(s[lo - 1] + (s[hi - 1] - s[lo - 1]) * (k - lo));
}

function percentRank(args: Lifted[], ctx: FnCtx, inclusive: boolean): Value {
  const xs = numbersOnly([args[0]], ctx);
  if (xs.length < 2) return err('N/A');
  const x = numArg(args, ctx, 1);
  if (isError(x)) return x;
  const s = [...xs].sort((a, b) => a - b);
  const digits = args.length > 2 ? numArg(args, ctx, 2) : 3;
  const sig = isError(digits) ? 3 : Math.trunc(digits);
  const below = s.filter((v) => v < x).length;
  const exact = s.findIndex((v) => v === x);
  if (exact < 0) return err('N/A', 'PERCENTRANK needs a value present in the array');
  const raw = inclusive ? below / (s.length - 1) : (below + 1) / (s.length + 1);
  return numOut(Number(raw.toFixed(Math.max(0, sig))));
}

function pairedNums(
  args: Lifted[],
  ctx: FnCtx,
): { xs: number[]; ys: number[] } | CellError {
  const first = numbersOnly([args[0]], ctx);
  const second = numbersOnly([args[1]], ctx);
  if (first.length !== second.length) {
    return err('N/A', 'Paired ranges must be the same size');
  }
  if (first.length < 2) return err('DIV/0', 'At least 2 paired values are required');
  return { xs: first, ys: second };
}

function correl(args: Lifted[], ctx: FnCtx): number | CellError {
  const p = pairedNums(args, ctx);
  if (isError(p)) return p;
  const mx = mean(p.xs);
  const my = mean(p.ys);
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < p.xs.length; i++) {
    const dx = p.xs[i] - mx;
    const dy = p.ys[i] - my;
    sxy += dx * dy;
    sxx += dx * dx;
    syy += dy * dy;
  }
  if (sxx === 0 || syy === 0) return err('DIV/0', 'CORREL needs variance in both inputs');
  return numOut(sxy / Math.sqrt(sxx * syy));
}

function forecast(args: Lifted[], ctx: FnCtx): number | CellError {
  const x = numArg(args, ctx, 0);
  if (isError(x)) return x;
  const ys = numbersOnly([args[1]], ctx);
  const xs = numbersOnly([args[2]], ctx);
  if (ys.length !== xs.length || ys.length < 2) return err('N/A', 'FORECAST needs matching arrays of 2+ values');
  const mx = mean(xs);
  const my = mean(ys);
  const den = devSq(xs);
  if (den === 0) return err('DIV/0');
  const m = xs.reduce((s, xv, i) => s + (xv - mx) * (ys[i] - my), 0) / den;
  return numOut(my + m * (x - mx));
}

function trend(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const ys = numbersOnly([args[0]], ctx);
  const xs = numbersOnly([args[1]], ctx);
  if (ys.length !== xs.length || ys.length < 2) return err('N/A', 'TREND needs matching arrays of 2+ values');
  const newX = args[2] !== undefined ? numbersOnly([args[2]], ctx) : xs;
  if (!newX.length) return err('N/A');
  const mx = mean(xs);
  const my = mean(ys);
  const den = devSq(xs);
  if (den === 0) return err('DIV/0');
  const m = xs.reduce((s, xv, i) => s + (xv - mx) * (ys[i] - my), 0) / den;
  const b = my - m * mx;
  return makeArray(newX.map((xv) => [m * xv + b]));
}

function growth(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const ys = numbersOnly([args[0]], ctx);
  if (ys.some((y) => y <= 0)) return err('NUM', 'GROWTH needs positive values');
  const xs = args[1] !== undefined ? numbersOnly([args[1]], ctx) : ys.map((_, i) => i + 1);
  if (ys.length !== xs.length || ys.length < 2) return err('N/A');
  const lny = ys.map((y) => Math.log(y));
  const newX = args[2] !== undefined ? numbersOnly([args[2]], ctx) : xs;
  const mx = mean(xs);
  const my = mean(lny);
  const den = devSq(xs);
  if (den === 0) return err('DIV/0');
  const m = xs.reduce((s, xv, i) => s + (xv - mx) * (lny[i] - my), 0) / den;
  const b = my - m * mx;
  return makeArray(newX.map((xv) => [Math.exp(m * xv + b)]));
}

/** LINEST returns {slope, intercept} as a 1x2 array like Excel. */
function linest(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const ys = numbersOnly([args[0]], ctx);
  const xs = args[1] !== undefined ? numbersOnly([args[1]], ctx) : ys.map((_, i) => i + 1);
  if (ys.length !== xs.length || ys.length < 2) return err('N/A');
  const mx = mean(xs);
  const my = mean(ys);
  const den = devSq(xs);
  if (den === 0) return err('DIV/0');
  const m = xs.reduce((s, xv, i) => s + (xv - mx) * (ys[i] - my), 0) / den;
  return makeArray([[m, my - m * mx]]);
}

function normDist(args: Lifted[], ctx: FnCtx): Value {
  const x = numArg(args, ctx, 0);
  if (isError(x)) return x;
  const mean = numArg(args, ctx, 1);
  if (isError(mean)) return mean;
  const sd = numArg(args, ctx, 2);
  if (isError(sd)) return sd;
  const cum = truthyArg(args[3]);
  if (sd <= 0) return err('NUM', 'Standard deviation must be positive');
  return numOut(cum ? normCdf(x, mean, sd) : normPdf(x, mean, sd));
}

function normInv(args: Lifted[], ctx: FnCtx): Value {
  const p = numArg(args, ctx, 0);
  if (isError(p)) return p;
  const mean = numArg(args, ctx, 1);
  if (isError(mean)) return mean;
  const sd = numArg(args, ctx, 2);
  if (isError(sd)) return sd;
  if (p <= 0 || p >= 1) return err('NUM', 'Probability must be between 0 and 1');
  if (sd <= 0) return err('NUM');
  return numOut(mean + sd * erfInv(p));
}

function binom(args: Lifted[], ctx: FnCtx): Value {
  const k = numArg(args, ctx, 0);
  if (isError(k)) return k;
  const n = numArg(args, ctx, 1);
  if (isError(n)) return n;
  const p = numArg(args, ctx, 2);
  if (isError(p)) return p;
  const cum = truthyArg(args[3]);
  if (k < 0 || n < 0 || k > n || p < 0 || p > 1) return err('NUM');
  const choose = chooseFn(n, k);
  if (cum) {
    let s = 0;
    for (let i = 0; i <= Math.trunc(k); i++) s += chooseFn(n, i) * Math.pow(p, i) * Math.pow(1 - p, n - i);
    return numOut(s);
  }
  return numOut(choose * Math.pow(p, k) * Math.pow(1 - p, n - k));
}

function chooseFn(n: number, k: number): number {
  const kk = Math.min(k, n - k);
  let out = 1;
  for (let i = 0; i < kk; i++) out = (out * (n - i)) / (i + 1);
  return out;
}

function poisson(args: Lifted[], ctx: FnCtx): Value {
  const x = numArg(args, ctx, 0);
  if (isError(x)) return x;
  const meanV = numArg(args, ctx, 1);
  if (isError(meanV)) return meanV;
  const cum = truthyArg(args[2]);
  if (x < 0 || x !== Math.trunc(x)) return err('NUM');
  if (cum) {
    let s = 0;
    for (let i = 0; i <= x; i++) s += Math.exp(-meanV) * Math.pow(meanV, i) / gammaFn(i + 1);
    return numOut(s);
  }
  return numOut(Math.exp(-meanV) * Math.pow(meanV, x) / gammaFn(x + 1));
}

function ztest(args: Lifted[], ctx: FnCtx): Value {
  const xs = numbersOnly([args[0]], ctx);
  if (xs.length < 2) return err('N/A', 'Z.TEST needs at least 2 values');
  const x = numArg(args, ctx, 1);
  if (isError(x)) return x;
  const sigma = args.length > 2 ? numArg(args, ctx, 2) : Math.sqrt(variance(xs, true));
  if (isError(sigma)) return sigma;
  if (!sigma) return err('DIV/0', 'Standard deviation cannot be zero');
  return numOut(1 - normCdf(x, mean(xs), sigma));
}

function ttest(args: Lifted[], ctx: FnCtx): Value {
  const xs = numbersOnly([args[0]], ctx);
  const ys = numbersOnly([args[1]], ctx);
  const tailsArg = numArg(args, ctx, 2);
  const typeArg = numArg(args, ctx, 3);
  if (isError(tailsArg) || isError(typeArg)) return err('VALUE', 'T.TEST needs numeric tails and type');
  const tails = Math.trunc(tailsArg);
  const type = Math.trunc(typeArg);
  if (isError(tails) || isError(type)) return err('VALUE');
  if (!xs.length || !ys.length) return err('N/A');
  let df: number;
  let t: number;
  if (type === 1) {
    if (xs.length + ys.length - 2 < 1) return err('DIV/0');
    const sp = Math.sqrt(
      ((xs.length - 1) * variance(xs, true) + (ys.length - 1) * variance(ys, true)) /
        (xs.length + ys.length - 2),
    );
    if (!Number.isFinite(sp) || sp === 0) return err('DIV/0');
    t = (mean(xs) - mean(ys)) / (sp * Math.sqrt(1 / xs.length + 1 / ys.length));
    df = xs.length + ys.length - 2;
  } else if (type === 2) {
    if (xs.length !== ys.length || xs.length < 2) return err('N/A', 'Paired test needs equal arrays');
    const d = xs.map((v, i) => v - ys[i]);
    const sd = Math.sqrt(variance(d, true));
    if (!sd) return err('DIV/0');
    t = mean(d) / (sd / Math.sqrt(d.length));
    df = d.length - 1;
  } else if (type === 3) {
    // Welch / two-sample unequal variance
    const v1 = variance(xs, true);
    const v2 = variance(ys, true);
    if (!Number.isFinite(v1) || !Number.isFinite(v2)) return err('DIV/0');
    const se = Math.sqrt(v1 / xs.length + v2 / ys.length);
    if (!se) return err('DIV/0');
    t = (mean(xs) - mean(ys)) / se;
    df = Math.pow(v1 / xs.length + v2 / ys.length, 2) /
      (Math.pow(v1 / xs.length, 2) / (xs.length - 1) + Math.pow(v2 / ys.length, 2) / (ys.length - 1));
  } else {
    return err('NUM', 'T.TEST type must be 1, 2 or 3');
  }
  const p = tCdf(t, df);
  return numOut(tails === 1 ? 2 * (1 - p) : 1 - p);
}

function chisqTest(args: Lifted[], ctx: FnCtx): Value {
  const obs = numbersOnly([args[0]], ctx);
  const exp = numbersOnly([args[1]], ctx);
  if (obs.length !== exp.length || obs.length === 0) return err('N/A');
  let stat = 0;
  for (let i = 0; i < obs.length; i++) {
    if (exp[i] === 0) return err('DIV/0');
    stat += (obs[i] - exp[i]) ** 2 / exp[i];
  }
  return numOut(1 - gammaincP(obs.length / 2, stat / 2));
}

function confidence(args: Lifted[], ctx: FnCtx, normal: boolean): Value {
  const alpha = numArg(args, ctx, 0);
  if (isError(alpha)) return alpha;
  const sd = numArg(args, ctx, 1);
  if (isError(sd)) return sd;
  const n = numArg(args, ctx, 2);
  if (isError(n)) return n;
  if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) return err('NUM');
  if (normal) {
    const z = Math.SQRT2 * erfInv(1 - alpha / 2);
    return numOut(z * sd / Math.sqrt(n));
  }
  if (n < 2) return err('NUM');
  const t = Math.SQRT2 * erfInv(1 - alpha / 2);
  const approx = invertCdf((x) => tCdf(x, n - 1), 1 - alpha / 2, -1e4, 1e4);
  return numOut(approx * sd / Math.sqrt(n));
}

/** PROB(x_range, prob_range, lower, [upper]). */
function prob(args: Lifted[], ctx: FnCtx): Value {
  const xs = numbersOnly([args[0]], ctx);
  const ps = numbersOnly([args[1]], ctx);
  if (xs.length !== ps.length) return err('N/A', 'PROB ranges must be the same size');
  const lo = numArg(args, ctx, 2);
  if (isError(lo)) return lo;
  const hi = args.length > 3 ? numArg(args, ctx, 3) : null;
  let total = 0;
  for (let i = 0; i < xs.length; i++) {
    const inRange = hi === null || isError(hi)
      ? xs[i] === lo
      : xs[i] >= lo && xs[i] <= hi;
    if (inRange) total += ps[i];
  }
  return numOut(total);
}

/** FREQUENCY(data, bins) -> vertical array of counts (len(bins)+1 rows). */
function frequency(args: Lifted[], ctx: FnCtx): Value | ArrayRef {
  const data = numbersOnly([args[0]], ctx);
  const bins = numbersOnly([args[1]], ctx).slice().sort((a, b) => a - b);
  const counts = new Array(bins.length + 1).fill(0);
  for (const v of data) {
    let placed = false;
    for (let i = 0; i < bins.length; i++) {
      if (v <= bins[i]) {
        counts[i]++;
        placed = true;
        break;
      }
    }
    if (!placed) counts[bins.length]++;
  }
  return makeArray(counts.map((c) => [c]));
}
