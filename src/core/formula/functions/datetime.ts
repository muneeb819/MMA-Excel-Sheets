/**
 * Excel date and time functions.
 *
 * A date/time value is an Excel serial number: the integer counts whole days
 * since 1899-12-31 and the fraction is the time of day, so 1.5 is "noon on
 * 1900-01-01" and 45292.645833 is 2024-01-01 15:30. `serialToDate` returns a
 * Date whose *UTC* fields carry the serial, so every part extractor below reads
 * the `getUTC*` accessors, and `msToSerial` is the exact inverse of that same
 * mapping — including the phantom 02-29-1900 of Excel's 1900 date system, which
 * is why both epoch constants are needed. TODAY()/NOW() read the machine's
 * local zone, like Excel does; everything else is pure serial arithmetic.
 */

import { err, isError, type CellError, type Value } from '../../types';
import { parseDateish, serialToDate, toNumber } from '../../coerce';
import { flat, type FnCtx, type FnDef, type Lifted } from '../fntypes';
import { firstError, firstScalar, numOut, numberAt, optNumber, optText } from '../helpers';

/* ------------------------------------------------------------ serial helpers */

const DAY_MS = 86_400_000;
const SEC_PER_DAY = 86_400;
/** `serialToDate` counts from 1899-12-30 for serial 61+ (Excel's leap-year bug)... */
const EPOCH_1900 = Date.UTC(1899, 11, 30);
/** ...and from 1899-12-31 below that, so serial 1 is 1900-01-01. */
const EPOCH_PLAIN = Date.UTC(1899, 11, 31);
const AFTER_LEAP_BUG = Date.UTC(1900, 2, 1);

function isLeap(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

/** Days in a (possibly out-of-range) zero-based month. */
function daysInMonth(year: number, month: number): number {
  if (month === 1) return isLeap(year) ? 29 : 28;
  return month === 3 || month === 5 || month === 8 || month === 10 ? 30 : 31;
}

/** UTC milliseconds for y/m/d h:m:s, with month/day overflow like Excel's DATE. */
function utcMs(y: number, m: number, d: number, h = 0, mi = 0, s = 0): number {
  const t = new Date(0);
  t.setUTCFullYear(y, m, d);
  t.setUTCHours(h, mi, s, 0);
  return t.getTime();
}

/** Exact inverse of `serialToDate` for whole days. */
function msToSerial(ms: number): number {
  return (ms - (ms >= AFTER_LEAP_BUG ? EPOCH_1900 : EPOCH_PLAIN)) / DAY_MS;
}

/** Midnight serial of a UTC date. */
function dateSerial(y: number, m: number, d: number): number {
  return msToSerial(utcMs(y, m, d));
}

/** Time-of-day fraction of a serial. */
function fracOf(serial: number): number {
  return serial - Math.floor(serial);
}

/** Monday = 0 ... Sunday = 6. */
function dowMon(serial: number): number {
  return (serialToDate(serial).getUTCDay() + 6) % 7;
}

function isWeekend(serial: number): boolean {
  const w = serialToDate(serial).getUTCDay();
  return w === 0 || w === 6;
}

/** ISO 8601 week number, Monday based: week 1 owns the first Thursday. */
function isoWeekNum(serial: number): number {
  const d = serialToDate(serial);
  const start = utcMs(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  const thursday = start + (3 - dowMon(serial)) * DAY_MS;
  // The Thursday, not the date, decides the ISO year: 2021-01-01 belongs to 2020.
  const jan4 = utcMs(new Date(thursday).getUTCFullYear(), 0, 4);
  const firstThursday = jan4 + (3 - ((new Date(jan4).getUTCDay() + 6) % 7)) * DAY_MS;
  return 1 + Math.round((thursday - firstThursday) / (7 * DAY_MS));
}

/* ------------------------------------------------------------- arg plumbing */

type FnBody = (args: Lifted[], ctx: FnCtx) => Value;

/** Wrap a body with the shared prologue: argument errors always win. */
function def(name: string, min: number, max: number, body: FnBody): FnDef {
  return {
    name,
    min,
    max,
    fn: (args, ctx) => {
      const e = firstError(args, ctx);
      return e === null ? body(args, ctx) : e;
    },
  };
}

/** Read `count` required numeric arguments, or the first failure. */
function numArgs(args: Lifted[], ctx: FnCtx, count: number): number[] | Value {
  const out: number[] = [];
  for (let i = 0; i < count; i++) {
    const n = numberAt(args, ctx, i);
    if (isError(n)) return n;
    if (typeof n !== 'number' || !Number.isFinite(n)) return err('VALUE', 'Expected a number');
    out.push(n);
  }
  return out;
}

/** Optional numeric argument with a fallback. */
function optNum(args: Lifted[], ctx: FnCtx, i: number, fallback: number): number | CellError {
  const v = optNumber(args, ctx, i, fallback);
  if (isError(v)) return v;
  return typeof v === 'number' && Number.isFinite(v) ? v : err('VALUE', 'Expected a number');
}

/** Serials below 1900 have no Excel representation. */
function pre1900(v: number, name: string): Value | null {
  return v < 0 ? err('NUM', `${name} needs a date on or after 1900`) : null;
}

/** Read required date serials, rejecting anything before 1900. */
function dateArgs(args: Lifted[], ctx: FnCtx, count: number, name: string): number[] | Value {
  const nums = numArgs(args, ctx, count);
  if (!Array.isArray(nums)) return nums;
  for (const v of nums) {
    const bad = pre1900(v, name);
    if (bad !== null) return bad;
  }
  return nums;
}

/** Holidays arrive as serials (or date text) in a range; blanks are ignored. */
function holidaySet(args: Lifted[], ctx: FnCtx, i: number): Set<number> {
  const out = new Set<number>();
  const arg = args[i];
  if (arg === undefined) return out;
  for (const v of flat(arg, ctx)) {
    if (typeof v === 'number') {
      if (Number.isFinite(v)) out.add(Math.floor(v));
    } else if (typeof v === 'string') {
      const p = parseDateish(v);
      if (p !== null && Number.isFinite(p)) out.add(Math.floor(p));
    }
  }
  return out;
}

/* -------------------------------------------------------- part function gen */

/** YEAR/MONTH/DAY... — one optional serial argument, defaulting to 0. */
function datePart(name: string, pick: (d: Date) => number): FnDef {
  return def(name, 0, 1, (args, ctx) => {
    const s = optNum(args, ctx, 0, 0);
    if (isError(s)) return s;
    const bad = pre1900(s, name);
    return bad === null ? pick(serialToDate(s)) : bad;
  });
}

/** HOUR/MINUTE/SECOND — the fraction is a real time even below serial 1. */
function timePart(name: string, pick: (d: Date) => number): FnDef {
  return def(name, 0, 1, (args, ctx) => {
    const s = optNum(args, ctx, 0, 0);
    return isError(s) ? s : pick(serialToDate(s));
  });
}

/* --------------------------------------------------------- shared algorithms */

/** Shift by whole months, keeping the day (clamped) or taking the month end. */
function shiftMonths(serial: number, months: number, monthEnd: boolean): Value {
  const d = serialToDate(serial);
  const total = d.getUTCMonth() + months;
  const year = d.getUTCFullYear() + Math.floor(total / 12);
  const month = ((total % 12) + 12) % 12;
  const last = daysInMonth(year, month);
  return numOut(dateSerial(year, month, monthEnd ? last : Math.min(d.getUTCDate(), last)));
}

function fullYears(from: Date, to: Date): number {
  let y = to.getUTCFullYear() - from.getUTCFullYear();
  if (
    to.getUTCMonth() < from.getUTCMonth() ||
    (to.getUTCMonth() === from.getUTCMonth() && to.getUTCDate() < from.getUTCDate())
  ) {
    y--;
  }
  return y;
}

function fullMonths(from: Date, to: Date): number {
  const m = (to.getUTCFullYear() - from.getUTCFullYear()) * 12 + (to.getUTCMonth() - from.getUTCMonth());
  return to.getUTCDate() < from.getUTCDate() ? m - 1 : m;
}

/** DATEDIF "MD": day-of-month difference, borrowing from the previous month. */
function dayOfMonthDiff(from: Date, to: Date): number {
  const d = to.getUTCDate() - from.getUTCDate();
  return d < 0 ? d + daysInMonth(to.getUTCFullYear(), to.getUTCMonth() - 1) : d;
}

/** DATEDIF "YM": whole months, years and days ignored. */
function monthDiffIgnoringYears(from: Date, to: Date): number {
  const m = (((to.getUTCMonth() - from.getUTCMonth()) % 12) + 12) % 12;
  return to.getUTCDate() < from.getUTCDate() ? (m + 11) % 12 : m;
}

/** DATEDIF "YD": whole days with the end date re-based on the start's year. */
function yearDayDiff(from: Date, to: Date): number {
  const y = from.getUTCFullYear();
  let day = to.getUTCDate();
  if (to.getUTCMonth() === 1 && day === 29 && !isLeap(y)) day = 28;
  const anchor = utcMs(y, to.getUTCMonth(), day);
  const start = utcMs(y, from.getUTCMonth(), from.getUTCDate());
  const diff = Math.round((anchor - start) / DAY_MS);
  return diff < 0 ? diff + (isLeap(y) ? 366 : 365) : diff;
}

const DATEDIF_UNITS: Record<string, (from: Date, to: Date) => number> = {
  Y: fullYears,
  M: fullMonths,
  MD: dayOfMonthDiff,
  YM: monthDiffIgnoringYears,
  YD: yearDayDiff,
};

/** 30/360 day count; `european` drops the US end-of-month exceptions. */
function days360(from: number, to: number, european: boolean): number {
  const f = serialToDate(from);
  const t = serialToDate(to);
  let d1 = f.getUTCDate();
  let d2 = t.getUTCDate();
  if (european) {
    d1 = Math.min(d1, 30);
    d2 = Math.min(d2, 30);
  } else {
    if (d1 === 31) d1 = 30;
    if (d2 === 31 && d1 === 30) d2 = 30;
  }
  return (
    360 * (t.getUTCFullYear() - f.getUTCFullYear()) +
    30 * (t.getUTCMonth() - f.getUTCMonth()) +
    (d2 - d1)
  );
}

/** Actual/actual: whole years count as 1, the rest over its own year length. */
function actualActual(from: number, to: number): number {
  const start = serialToDate(from);
  const whole = fullYears(start, serialToDate(to));
  if (whole === 0) {
    return (to - from) / (isLeap(start.getUTCFullYear()) ? 366 : 365);
  }
  // Re-base the start on its own anniversary, clamping 29 Feb onto 28 Feb.
  const month = start.getUTCMonth();
  const year = start.getUTCFullYear() + whole;
  const anchor = dateSerial(year, month, Math.min(start.getUTCDate(), daysInMonth(year, month)));
  const rest = (to - anchor) / (isLeap(serialToDate(anchor).getUTCFullYear()) ? 366 : 365);
  return whole + rest;
}

/** Whole days in [from, to] falling on `target` (0 = Monday). */
function countWeekday(from: number, to: number, target: number): number {
  const first = from + ((target - dowMon(from) + 7) % 7);
  return first > to ? 0 : Math.floor((to - first) / 7) + 1;
}

/* ----------------------------------------------------------------- functions */

export const dateFns: FnDef[] = [
  def('DATE', 3, 3, (args, ctx) => {
    const n = numArgs(args, ctx, 3);
    if (!Array.isArray(n)) return n;
    let year = Math.trunc(n[0]);
    if (year >= 0 && year <= 99) year += year < 30 ? 2000 : 1900;
    // Month and day overflow roll into the surrounding months for free.
    const ms = utcMs(year, Math.trunc(n[1]) - 1, Math.trunc(n[2]));
    if (!Number.isFinite(ms)) return err('NUM', 'DATE is out of the supported range');
    const serial = msToSerial(ms);
    if (!Number.isFinite(serial) || serial < 0) {
      return err('NUM', 'DATE cannot produce a date before 1900');
    }
    return numOut(serial);
  }),

  def('DATEVALUE', 1, 1, (args, ctx) => {
    const v = firstScalar(args, ctx);
    if (typeof v === 'string') {
      const p = parseDateish(v);
      return p === null || !Number.isFinite(p)
        ? err('VALUE', `DATEVALUE cannot read "${v}" as a date`)
        : Math.floor(p);
    }
    // A bare number already is a serial, so it is taken as-is.
    const n = toNumber(v);
    return typeof n === 'number' ? numOut(n) : n;
  }),

  def('TIME', 3, 3, (args, ctx) => {
    const n = numArgs(args, ctx, 3);
    if (!Array.isArray(n)) return n;
    // Hours past midnight carry into whole days instead of wrapping away.
    const secs = Math.trunc(n[0]) * 3600 + Math.trunc(n[1]) * 60 + Math.trunc(n[2]);
    return numOut(secs / SEC_PER_DAY);
  }),

  def('TIMEVALUE', 1, 1, (args, ctx) => {
    const v = firstScalar(args, ctx);
    if (typeof v === 'string') {
      const p = parseDateish(v);
      if (p === null || !Number.isFinite(p)) {
        return err('VALUE', `TIMEVALUE cannot read "${v}" as a time`);
      }
      // "2024-01-01 13:30" parses to a full serial; keep only the clock.
      return numOut(fracOf(p));
    }
    const n = toNumber(v);
    return typeof n === 'number' ? numOut(fracOf(n)) : n;
  }),

  {
    name: 'TODAY', min: 0, max: 0, volatile: true,
    fn: () => {
      const n = new Date();
      return numOut(dateSerial(n.getFullYear(), n.getMonth(), n.getDate()));
    },
  },
  {
    name: 'NOW', min: 0, max: 0, volatile: true,
    fn: () => {
      const n = new Date();
      const day = dateSerial(n.getFullYear(), n.getMonth(), n.getDate());
      const secs =
        n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds() + n.getMilliseconds() / 1000;
      return numOut(day + secs / SEC_PER_DAY);
    },
  },

  datePart('YEAR', (d) => d.getUTCFullYear()),
  datePart('MONTH', (d) => d.getUTCMonth() + 1),
  datePart('DAY', (d) => d.getUTCDate()),
  timePart('HOUR', (d) => d.getUTCHours()),
  timePart('MINUTE', (d) => d.getUTCMinutes()),
  timePart('SECOND', (d) => d.getUTCSeconds()),

  def('WEEKDAY', 0, 2, (args, ctx) => {
    const s = optNum(args, ctx, 0, 0);
    if (isError(s)) return s;
    const t = optNum(args, ctx, 1, 1);
    if (isError(t)) return t;
    const bad = pre1900(s, 'WEEKDAY');
    if (bad !== null) return bad;
    const type = Math.trunc(t);
    const dow = serialToDate(Math.trunc(s)).getUTCDay();
    if (type === 1) return dow + 1; // 1 = Sunday
    if (type === 2 || (type >= 11 && type <= 17)) return ((dow + 6) % 7) + 1; // 1 = Monday
    if (type === 3) return (dow + 6) % 7; // 0 = Monday
    return err('NUM', `WEEKDAY type ${type} is invalid`);
  }),

  def('ISOWEEKNUM', 0, 1, (args, ctx) => {
    const s = optNum(args, ctx, 0, 0);
    if (isError(s)) return s;
    const bad = pre1900(s, 'ISOWEEKNUM');
    return bad === null ? isoWeekNum(Math.trunc(s)) : bad;
  }),

  def('WEEKNUM', 0, 2, (args, ctx) => {
    const s = optNum(args, ctx, 0, 0);
    if (isError(s)) return s;
    const t = optNum(args, ctx, 1, 1);
    if (isError(t)) return t;
    const bad = pre1900(s, 'WEEKNUM');
    if (bad !== null) return bad;
    const type = Math.trunc(t);
    if (type !== 1 && type !== 2 && type !== 21) return err('NUM', `WEEKNUM type ${type} is invalid`);
    const day = Math.trunc(s);
    if (type === 21) return isoWeekNum(day);
    const jan1 = dateSerial(serialToDate(day).getUTCFullYear(), 0, 1);
    const jan1Dow = serialToDate(jan1).getUTCDay();
    const shift = type === 1 ? jan1Dow : (jan1Dow + 6) % 7;
    return Math.floor((day - jan1 + shift) / 7) + 1;
  }),

  def('EDATE', 2, 2, (args, ctx) => {
    const n = dateArgs(args, ctx, 2, 'EDATE');
    return Array.isArray(n) ? shiftMonths(n[0], Math.trunc(n[1]), false) : n;
  }),

  def('EOMONTH', 2, 2, (args, ctx) => {
    const n = dateArgs(args, ctx, 2, 'EOMONTH');
    return Array.isArray(n) ? shiftMonths(n[0], Math.trunc(n[1]), true) : n;
  }),

  def('DAYS', 2, 2, (args, ctx) => {
    const n = numArgs(args, ctx, 2);
    return Array.isArray(n) ? numOut(n[0] - n[1]) : n;
  }),

  def('DATEDIF', 3, 3, (args, ctx) => {
    const n = dateArgs(args, ctx, 2, 'DATEDIF');
    if (!Array.isArray(n)) return n;
    if (n[0] > n[1]) return err('NUM', 'DATEDIF start date is after the end date');
    const unit = optText(args, ctx, 2, '').trim().toUpperCase();
    if (unit === 'D') return numOut(Math.trunc(n[1]) - Math.trunc(n[0]));
    const calc = DATEDIF_UNITS[unit];
    if (!calc) return err('NUM', `DATEDIF unit "${unit}" is not recognised`);
    return numOut(calc(serialToDate(Math.trunc(n[0])), serialToDate(Math.trunc(n[1]))));
  }),

  def('YEARFRAC', 2, 3, (args, ctx) => {
    const basis = optNum(args, ctx, 2, 0);
    if (isError(basis)) return basis;
    const b = Math.trunc(basis);
    if (b < 0 || b > 4) return err('NUM', `YEARFRAC basis ${b} is invalid`);
    const n = dateArgs(args, ctx, 2, 'YEARFRAC');
    if (!Array.isArray(n)) return n;
    const from = Math.trunc(Math.min(n[0], n[1]));
    const to = Math.trunc(Math.max(n[0], n[1]));
    const days = to - from;
    const frac =
      b === 0 || b === 4 ? days360(from, to, b === 4) / 360
      : b === 1 ? actualActual(from, to)
      : b === 2 ? days / 360
      : days / 365;
    return numOut((n[0] <= n[1] ? 1 : -1) * frac);
  }),

  def('WORKDAY', 2, 3, (args, ctx) => {
    const n = dateArgs(args, ctx, 2, 'WORKDAY');
    if (!Array.isArray(n)) return n;
    const holidays = holidaySet(args, ctx, 2);
    const count = Math.trunc(n[1]);
    let cur = Math.trunc(n[0]);
    if (count === 0) return numOut(cur);
    const step = count > 0 ? 1 : -1;
    let left = count;
    for (let guard = 0; left !== 0 && guard < 1_000_000; guard++) {
      cur += step;
      if (!isWeekend(cur) && !holidays.has(cur)) left -= step;
    }
    return left === 0 ? numOut(cur) : err('NUM', 'WORKDAY found no free working day');
  }),

  def('NETWORKDAYS', 2, 3, (args, ctx) => {
    const n = dateArgs(args, ctx, 2, 'NETWORKDAYS');
    if (!Array.isArray(n)) return n;
    const holidays = holidaySet(args, ctx, 2);
    const from = Math.trunc(Math.min(n[0], n[1]));
    const to = Math.trunc(Math.max(n[0], n[1]));
    let count = to - from + 1 - countWeekday(from, to, 5) - countWeekday(from, to, 6);
    for (const h of holidays) {
      if (h >= from && h <= to && !isWeekend(h)) count--;
    }
    return numOut((n[0] <= n[1] ? 1 : -1) * count);
  }),
];
