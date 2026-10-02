/**
 * The FINANCIAL family: PMT, PV, FV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC,
 * NPV, IRR, MIRR, XNPV and XIRR.
 *
 * Sign convention is Excel's, not the textbook's: cash you receive is positive
 * and cash you pay out is negative, so PMT(0.1/12, 24, 10000) is -505.54. PV,
 * FV and PMT are one balance identity rearranged, and the IPMT family is
 * derived from that identity rather than from an amortisation table, so the
 * zero-rate and beginning-of-period cases fall out for free.
 *
 * NPV/IRR/MIRR/XNPV/XIRR instead take a raw cash-flow series whose FIRST
 * element sits at time 0 and discount from the second element onwards.
 *
 * Nothing here throws: an argument that is not a number yields #VALUE! and a
 * degenerate or unsolvable one yields #NUM!.
 */

import type { FnDef, FnCtx, Lifted } from '../fntypes';
import { err, isError, type CellError, type Value } from '../../types';
import { argFlat, firstError, numberAt, numOut, optNumber } from '../helpers';

/* ------------------------------------------------------------------ inputs */

/**
 * The preamble every entry point shares: propagate the first error anywhere in
 * the arguments, then coerce the documented arity to numbers. `required`
 * leading arguments are mandatory, `optional` gives the trailing fallbacks.
 */
function read(
  args: Lifted[],
  ctx: FnCtx,
  required: number,
  optional: number[],
): number[] | CellError {
  const e = firstError(args, ctx);
  if (isError(e)) return e;
  const out: number[] = [];
  for (let i = 0; i < required; i++) {
    const v = numberAt(args, ctx, i);
    if (typeof v !== 'number') return v;
    out.push(v);
  }
  for (let j = 0; j < optional.length; j++) {
    const v = optNumber(args, ctx, required + j, optional[j]);
    if (typeof v !== 'number') return v;
    out.push(v);
  }
  return out;
}

/** Numeric cells of a cash-flow range, skipping blanks and text as Excel does. */
function flowNums(vals: Value[]): number[] {
  const out: number[] = [];
  for (const v of vals) if (typeof v === 'number') out.push(v);
  return out;
}

/** Excel's `type` switch: any non-zero value means payments at period start. */
function typeFlag(type: number): number {
  return type !== 0 ? 1 : 0;
}

/** A series only has a meaningful IRR/XIRR/NPV if its signs differ. */
function mixedSigns(values: number[]): boolean {
  let pos = false;
  let neg = false;
  for (const v of values) {
    if (v > 0) pos = true;
    else if (v < 0) neg = true;
  }
  return pos && neg;
}

/** Sum of magnitudes: the natural scale for a solver's residual tolerance. */
function magnitude(values: number[]): number {
  let sum = 0;
  for (const v of values) sum += Math.abs(v);
  return sum;
}

/* -------------------------------------------------------------- primitives */

/** ((1+rate)^nper - 1)/rate, taking the rate = 0 limit of nper. */
function annFactor(rate: number, nper: number): number {
  if (rate === 0) return nper;
  return (Math.pow(1 + rate, nper) - 1) / rate;
}

/** PV/FV/PMT, all one identity; rate === 0 means no compounding at all. */
function pmtOf(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  return -(pv * Math.pow(1 + rate, nper) + fv) / (annFactor(rate, nper) * (1 + rate * type));
}

function fvOf(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + pmt * nper);
  return -(pv * Math.pow(1 + rate, nper) + pmt * (1 + rate * type) * annFactor(rate, nper));
}

function nperOf(rate: number, pmt: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return pmt === 0 ? NaN : -(pv + fv) / pmt;
  const adj = (pmt * (1 + rate * type)) / rate;
  const ratio = (adj - fv) / (pv + adj);
  if (!Number.isFinite(ratio) || ratio <= 0) return NaN;
  return Math.log(ratio) / Math.log(1 + rate);
}

/**
 * Excel's IPMT: the interest slice of the payment in period `per` (1-based),
 * negative because paying interest is money leaving the account.
 */
function ipmtOf(rate: number, per: number, nper: number, pv: number, fv: number, type: number): number {
  if (per === 1 && type === 1) return 0;
  const pmt = pmtOf(rate, nper, pv, fv, type);
  if (type === 1) return fvOf(rate, per - 2, pmt, pv, 1) * rate;
  return fvOf(rate, per - 1, pmt, pv, 0) * rate;
}

/** PPMT is simply the rest of the payment. */
function ppmtOf(rate: number, per: number, nper: number, pv: number, fv: number, type: number): number {
  return pmtOf(rate, nper, pv, fv, type) - ipmtOf(rate, per, nper, pv, fv, type);
}

/* ----------------------------------------------------------- root finding */

const SOLVE_ITER = 128;
const SOLVE_TOL = 1e-10;
const RESID_TOL = 1e-9;

/** Bisection over a bracketed interval; null when the ends share a sign. */
function bisect(f: (x: number) => number, lo: number, hi: number): number | null {
  let flo = f(lo);
  const fhi = f(hi);
  if (!Number.isFinite(flo) || !Number.isFinite(fhi)) return null;
  if (flo === 0) return lo;
  if (fhi === 0) return hi;
  if (flo * fhi > 0) return null;
  let a = lo;
  let b = hi;
  for (let i = 0; i < SOLVE_ITER; i++) {
    const m = (a + b) / 2;
    const fm = f(m);
    if (!Number.isFinite(fm)) return null;
    if (fm === 0 || Math.abs(b - a) <= SOLVE_TOL * Math.max(1, Math.abs(m))) return m;
    if (flo * fm < 0) b = m;
    else {
      a = m;
      flo = fm;
    }
  }
  return (a + b) / 2;
}

/** Brackets tried in order once Newton has stalled, widest guess first. */
const BRACKETS: readonly (readonly [number, number])[] = [
  [-0.9999, 1e6],
  [-0.9999, 1e3],
  [-0.99, 100],
  [-0.9, 10],
  [-0.5, 2],
];

/**
 * Newton with a central-difference derivative, falling back to bisection.
 * `scale` normalises the residual so huge cash-flow series still converge.
 */
function solveRoot(f: (x: number) => number, guess: number, scale: number): number | null {
  const limit = RESID_TOL * Math.max(1, Math.abs(scale));
  let x = Number.isFinite(guess) ? guess : 0.1;
  for (let i = 0; i < SOLVE_ITER; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) break;
    if (Math.abs(y) <= limit) return x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-6);
    const slope = (f(x + h) - f(x - h)) / (2 * h);
    if (!Number.isFinite(slope) || slope === 0) break;
    const next = x - y / slope;
    if (!Number.isFinite(next) || next <= -1) break;
    if (Math.abs(next - x) <= SOLVE_TOL * Math.max(1, Math.abs(next))) return next;
    x = next;
  }
  for (const [lo, hi] of BRACKETS) {
    const root = bisect(f, lo, hi);
    if (root !== null) return root;
  }
  return null;
}

/** Excel discounts each flow by whole years from the first date (365-day). */
function xnpvOf(rate: number, values: number[], dates: number[], d0: number): number {
  let total = 0;
  for (let i = 0; i < values.length; i++) {
    total += values[i] / Math.pow(1 + rate, (dates[i] - d0) / 365);
  }
  return total;
}

/* ------------------------------------------------------------- the library */

export const financialFns: FnDef[] = [
  {
    name: 'PMT',
    min: 3,
    max: 5,
    fn: (args, ctx) => {
      const a = read(args, ctx, 3, [0, 0]);
      if (isError(a)) return a;
      const [rate, nper, pv, fv, type] = a;
      return numOut(pmtOf(rate, nper, pv, fv, typeFlag(type)));
    },
  },
  {
    name: 'PV',
    min: 3,
    max: 5,
    fn: (args, ctx) => {
      const a = read(args, ctx, 3, [0, 0]);
      if (isError(a)) return a;
      const [rate, nper, pmt, fv, type] = a;
      const t = typeFlag(type);
      // rate = 0 degenerates to the limit of the general formula, -(fv + pmt*nper).
      if (rate === 0) return numOut(-(fv + pmt * nper));
      const bal = fv + pmt * (1 + rate * t) * annFactor(rate, nper);
      return numOut(-bal / Math.pow(1 + rate, nper));
    },
  },
  {
    name: 'FV',
    min: 3,
    max: 5,
    fn: (args, ctx) => {
      const a = read(args, ctx, 3, [0, 0]);
      if (isError(a)) return a;
      const [rate, nper, pmt, pv, type] = a;
      return numOut(fvOf(rate, nper, pmt, pv, typeFlag(type)));
    },
  },
  {
    name: 'NPER',
    min: 3,
    max: 5,
    fn: (args, ctx) => {
      const a = read(args, ctx, 3, [0, 0]);
      if (isError(a)) return a;
      const [rate, pmt, pv, fv, type] = a;
      return numOut(nperOf(rate, pmt, pv, fv, typeFlag(type)));
    },
  },
  {
    name: 'RATE',
    min: 3,
    max: 6,
    volatile: false,
    fn: (args, ctx) => {
      const a = read(args, ctx, 3, [0, 0, 0.1]);
      if (isError(a)) return a;
      const [nper, pmt, pv, fv, type, guess] = a;
      const t = typeFlag(type);
      // Balance after `nper` periods, which is zero exactly at the right rate.
      const f = (r: number): number => {
        if (r <= -1) return NaN;
        if (r === 0) return pv + pmt * nper + fv;
        return pv * Math.pow(1 + r, nper) + pmt * (1 + r * t) * annFactor(r, nper) + fv;
      };
      const root = solveRoot(f, guess, magnitude([pv, pmt * nper, fv]));
      if (root === null) return err('NUM', 'RATE could not find a solution');
      return numOut(root);
    },
  },
  {
    name: 'IPMT',
    min: 4,
    max: 6,
    fn: (args, ctx) => {
      const a = read(args, ctx, 4, [0, 0]);
      if (isError(a)) return a;
      const [rate, per, nper, pv, fv, type] = a;
      if (per < 1 || per > nper) return err('NUM', 'IPMT period is out of range');
      return numOut(ipmtOf(rate, per, nper, pv, fv, typeFlag(type)));
    },
  },
  {
    name: 'PPMT',
    min: 4,
    max: 6,
    fn: (args, ctx) => {
      const a = read(args, ctx, 4, [0, 0]);
      if (isError(a)) return a;
      const [rate, per, nper, pv, fv, type] = a;
      if (per < 1 || per > nper) return err('NUM', 'PPMT period is out of range');
      return numOut(ppmtOf(rate, per, nper, pv, fv, typeFlag(type)));
    },
  },
  {
    name: 'CUMIPMT',
    min: 5,
    max: 6,
    fn: (args, ctx) => {
      const a = read(args, ctx, 5, [0]);
      if (isError(a)) return a;
      const [rate, nper, pv, start, end, type] = a;
      const from = Math.floor(start);
      const to = Math.floor(end);
      if (from < 1 || from > to || to > nper) return err('NUM', 'CUMIPMT period range is out of range');
      const t = typeFlag(type);
      let total = 0;
      for (let per = from; per <= to; per++) total += ipmtOf(rate, per, nper, pv, 0, t);
      return numOut(total);
    },
  },
  {
    name: 'CUMPRINC',
    min: 5,
    max: 6,
    fn: (args, ctx) => {
      const a = read(args, ctx, 5, [0]);
      if (isError(a)) return a;
      const [rate, nper, pv, start, end, type] = a;
      const from = Math.floor(start);
      const to = Math.floor(end);
      if (from < 1 || from > to || to > nper) return err('NUM', 'CUMPRINC period range is out of range');
      const t = typeFlag(type);
      const pmt = pmtOf(rate, nper, pv, 0, t);
      let total = 0;
      for (let per = from; per <= to; per++) total += pmt - ipmtOf(rate, per, nper, pv, 0, t);
      return numOut(total);
    },
  },
  {
    name: 'NPV',
    min: 2,
    max: 254,
    fn: (args, ctx) => {
      const rate = numberAt(args, ctx, 0);
      if (typeof rate !== 'number') return rate;
      const flows: number[] = [];
      for (let i = 1; i < args.length; i++) {
        for (const v of argFlat(args[i], ctx)) if (typeof v === 'number') flows.push(v);
      }
      let total = 0;
      if (rate === 0) {
        for (const v of flows) total += v;
      } else {
        // The first cash flow is discounted one period, exactly like Excel.
        for (let i = 0; i < flows.length; i++) total += flows[i] / Math.pow(1 + rate, i + 1);
      }
      return numOut(total);
    },
  },
  {
    name: 'IRR',
    min: 1,
    max: 2,
    fn: (args, ctx) => {
      const guess = optNumber(args, ctx, 1, 0.1);
      if (typeof guess !== 'number') return guess;
      const flows = flowNums(argFlat(args[0], ctx));
      if (flows.length < 2) return err('NUM', 'IRR needs at least one period');
      if (!mixedSigns(flows)) return err('NUM', 'IRR needs a positive and a negative value');
      // The first flow is at time 0, so it is never discounted.
      const f = (r: number): number => {
        if (r <= -1) return NaN;
        let sum = 0;
        let p = 1;
        for (const v of flows) {
          sum += v * p;
          p /= 1 + r;
        }
        return sum;
      };
      const root = solveRoot(f, guess, magnitude(flows));
      if (root === null) return err('NUM', 'IRR did not converge');
      return numOut(root);
    },
  },
  {
    name: 'MIRR',
    min: 3,
    max: 3,
    fn: (args, ctx) => {
      const a = read(args, ctx, 3, []);
      if (isError(a)) return a;
      const [, rateIn, rateOut] = a;
      const flows = flowNums(argFlat(args[0], ctx));
      if (flows.length < 2) return err('NUM', 'MIRR needs at least one period');
      if (rateIn <= -1 || rateOut <= -1) return err('NUM', 'MIRR rates must be greater than -1');
      const nper = flows.length - 1;
      // Positives compound to the end at the reinvestment rate; negatives are
      // discounted back to today at the financing rate.
      let grown = 0;
      let financed = 0;
      for (let i = 0; i < flows.length; i++) {
        const v = flows[i];
        if (v > 0) grown += v * Math.pow(1 + rateOut, nper - i);
        else if (v < 0) financed += v / Math.pow(1 + rateIn, i);
      }
      if (grown <= 0 || financed >= 0) return err('NUM', 'MIRR needs a positive and a negative value');
      return numOut(Math.pow(grown / -financed, 1 / nper) - 1);
    },
  },
  {
    name: 'XNPV',
    min: 3,
    max: 3,
    fn: (args, ctx) => {
      const rate = numberAt(args, ctx, 0);
      if (typeof rate !== 'number') return rate;
      const values = flowNums(argFlat(args[1], ctx));
      const dates = flowNums(argFlat(args[2], ctx));
      if (values.length !== dates.length) return err('NUM', 'XNPV needs one date per value');
      if (values.length === 0) return err('NUM', 'XNPV needs at least one value');
      if (rate <= -1) return err('NUM', 'XNPV rate must be greater than -1');
      if (!mixedSigns(values)) return err('NUM', 'XNPV needs a positive and a negative value');
      return numOut(xnpvOf(rate, values, dates, dates[0]));
    },
  },
  {
    name: 'XIRR',
    min: 2,
    max: 3,
    fn: (args, ctx) => {
      const guess = optNumber(args, ctx, 2, 0.1);
      if (typeof guess !== 'number') return guess;
      const values = flowNums(argFlat(args[0], ctx));
      const dates = flowNums(argFlat(args[1], ctx));
      if (values.length !== dates.length) return err('NUM', 'XIRR needs one date per value');
      if (values.length === 0) return err('NUM', 'XIRR needs at least one value');
      if (!mixedSigns(values)) return err('NUM', 'XIRR needs a positive and a negative value');
      const d0 = dates[0];
      const f = (r: number): number => (r <= -1 ? NaN : xnpvOf(r, values, dates, d0));
      const root = solveRoot(f, guess, magnitude(values));
      if (root === null) return err('NUM', 'XIRR did not converge');
      return numOut(root);
    },
  },
];
