/**
 * Value coercion, Excel serial dates, and the number-format engine.
 *
 * These rules are what make the app *feel* like Excel: `50%` typed into a cell
 * becomes 0.5, dates entered as text become serials, and formats like
 * `#,##0.00;[Red](#,##0.00)` render exactly as expected.
 */

import { err, isError, type CellError, type Value } from './types';

const ERR_TEXT: Record<string, string> = {
  NULL: '#NULL!',
  'DIV/0': '#DIV/0!',
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  NUM: '#NUM!',
  'N/A': '#N/A',
  CYCLE: '#CIRCULAR!',
  CALC: '#CALC!',
  SPILL: '#SPILL!',
  PARSE: '#ERROR!',
  GETTING_DATA: '#GETTING_DATA',
};

export function errorText(e: { __error: string; message?: string }): string {
  return ERR_TEXT[e.__error] ?? '#ERROR!';
}

/* ------------------------------------------------------------- serial dates */

const DAY = 86_400_000;
/** Excel wrongly treats 1900 as a leap year; serial 60 is the phantom 02-29. */
export function serialToDate(serial: number): Date {
  const n = Math.floor(serial);
  const frac = serial - n;
  const base = n >= 61 ? Date.UTC(1899, 11, 30) : Date.UTC(1899, 11, 31);
  const d = new Date(base + n * DAY + Math.round(frac * DAY));
  return d;
}

export function dateToSerial(d: Date): number {
  const utc = Date.UTC(
    d.getUTCFullYear(),
    d.getUTCMonth(),
    d.getUTCDate(),
    d.getUTCHours(),
    d.getUTCMinutes(),
    d.getUTCSeconds(),
    d.getUTCMilliseconds(),
  );
  // Days since 1899-12-31, then +1 from 1900-03-01 onwards because Excel
  // believes 1900 was a leap year (serial 60 is the phantom 29 Feb 1900).
  const days = Math.floor((utc - Date.UTC(1899, 11, 31)) / DAY);
  const dayFrac =
    (d.getUTCHours() * 3600 + d.getUTCMinutes() * 60 + d.getUTCSeconds()) / 86400 +
    d.getUTCMilliseconds() / 86_400_000;
  return (days >= 60 ? days + 1 : days) + dayFrac;
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Recognise the date/text shapes Excel accepts, returning a serial or null. */
export function parseDateish(input: string): number | null {
  const s = input.trim();
  if (!s) return null;

  // ISO date  yyyy-mm-dd  /  yyyy/mm/dd
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s);
  if (m) return mk(+m[1], +m[2] - 1, +m[3]);

  // US style m/d/y or m/d/yy
  m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2}|\d{4})$/.exec(s);
  if (m) {
    let year = +m[3];
    if (year < 100) year += year < 30 ? 2000 : 1900;
    return mk(year, +m[1] - 1, +m[2]);
  }

  // d-mmm-yyyy / mmm d, yyyy
  m = /^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{2,4})$/.exec(s);
  if (m) {
    const mo = monthIdx(m[2]);
    if (mo >= 0) {
      let year = +m[3];
      if (year < 100) year += year < 30 ? 2000 : 1900;
      return mk(year, mo, +m[1]);
    }
  }
  m = /^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/.exec(s);
  if (m) {
    const mo = monthIdx(m[1]);
    if (mo >= 0) return mk(+m[3], mo, +m[2]);
  }

  // datetime  yyyy-mm-dd[h]mm[:ss] / m/d/yyyy hh:mm
  m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})[T\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (m) {
    return mk(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], m[6] ? +m[6] : 0);
  }
  m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (m) return mk(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], m[6] ? +m[6] : 0);

  // ISO timestamps with a timezone offset -> treat as instant
  m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})[T ](\d{1,2}):(\d{2})(?::(\d{2}))?(Z|[+-]\d{2}:?\d{2})$/.exec(s);
  if (m) {
    const off = m[7] === 'Z' ? 0 : tzMinutes(m[7]);
    const base = mk(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], m[6] ? +m[6] : 0);
    return base - off / 1440;
  }

  // bare clock time
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (m) return (+m[1] * 3600 + +m[2] * 60 + (m[3] ? +m[3] : 0)) / 86400;

  return null;
}

function tzMinutes(z: string): number {
  const sign = z[0] === '-' ? -1 : 1;
  const t = z.slice(1).replace(':', '');
  return sign * (+t.slice(0, 2) * 60 + +t.slice(2, 4));
}

function monthIdx(name: string): number {
  const n = name.slice(0, 3).toLowerCase();
  return MONTHS.findIndex((m) => m.toLowerCase().startsWith(n));
}

function mk(
  y: number,
  mo: number,
  d: number,
  h = 0,
  mi = 0,
  s = 0,
): number {
  if (mo < 0 || mo > 11 || d < 1 || d > 31) return NaN;
  const dt = new Date(Date.UTC(y, mo, d, h, mi, s));
  if (dt.getUTCMonth() !== mo || dt.getUTCDate() !== d) return NaN;
  return dateToSerial(dt);
}

/**
 * Heuristic used to decide whether a number-format string is a date format.
 *
 * Date formats never contain the digit placeholders `#` or `0` once quoted
 * text, escapes and `[...]` tokens are removed, whereas numeric formats always
 * contain at least one of them.
 */
export function isDateFormat(fmt: string): boolean {
  const stripped = stripLiterals(fmt);
  if (/[#0]/.test(stripped)) return false;
  return /y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}/i.test(stripped);
}

function stripLiterals(s: string): string {
  return s
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '');
}

/* --------------------------------------------------------------- coercion */

const PERCENT = /^([+-]?[\d,]*\.?\d+)\s*%$/;
const CURRENCY = /^([$€£¥])\s*([+-]?[\d,]*\.?\d+)$|^([+-]?[\d,]*\.?\d+)\s*([$€£¥])$/;

/**
 * Parse what the user typed into a literal value.
 * Returns `undefined` when the text should be stored as a plain string.
 */
export function parseInput(text: string): number | string | boolean | null {
  const s = text.trim();
  if (s === '') return null;
  if (/^(TRUE|FALSE)$/i.test(s)) return /^true$/i.test(s);

  const pct = PERCENT.exec(s);
  if (pct) return parseFloat(pct[1].replace(/,/g, '')) / 100;

  const cur = CURRENCY.exec(s);
  if (cur) return parseFloat((cur[2] ?? cur[3]).replace(/,/g, ''));

  if (/^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(s)) {
    return parseFloat(s.replace(/,/g, ''));
  }
  if (/^[+-]?\d*\.?\d+([eE][+-]?\d+)?$/.test(s)) {
    const n = Number(s);
    if (!Number.isNaN(n)) return n;
  }
  const d = parseDateish(s);
  if (d !== null && !Number.isNaN(d)) return d;
  return text;
}

/** Excel's text-to-number coercion, used by arithmetic and by VALUE(). */
export function toNumber(v: Value): number | CellError {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const s = v.trim();
  if (s === '') return 0;
  const pct = PERCENT.exec(s);
  if (pct) return parseFloat(pct[1].replace(/,/g, '')) / 100;
  const d = parseDateish(s);
  if (d !== null && !Number.isNaN(d)) return d;
  const n = Number(s.replace(/,/g, ''));
  if (Number.isNaN(n)) return err('VALUE', `Cannot convert "${v}" to a number`);
  return n;
}

export function toText(v: Value): string {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return errorText(v);
  return formatGeneral(v);
}

export function toBool(v: Value): boolean | CellError {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null) return false;
  if (isError(v)) return v;
  const s = v.trim().toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  const n = toNumber(v);
  if (typeof n === 'number') return n !== 0;
  return err('VALUE', `"${v}" is not a logical value`);
}

/* --------------------------------------------------------------- rendering */

/** Excel "General": up to 11 significant digits, scientific when unwieldy. */
export function formatGeneral(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const abs = Math.abs(n);
  if (abs >= 1e11 || abs < 1e-9) return trimSci(n);
  const rounded = roundTo(n, 10);
  let s = String(rounded);
  if (s.includes('e')) s = trimSci(rounded);
  return s;
}

function roundTo(n: number, digits: number): number {
  if (n === 0) return 0;
  const f = 10 ** Math.min(digits, 20);
  return Math.round(n * f) / f;
}

function trimSci(n: number): string {
  const [m, e] = n.toExponential(8).split('e');
  const mantissa = m.replace(/\.?0+$/, '');
  const exp = parseInt(e, 10);
  return `${mantissa}E${exp >= 0 ? '+' : '-'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

/** Render a value using an Excel number-format code. */
export function formatValue(v: Value, fmt?: string): string {
  if (isError(v)) return errorText(v);
  if (!fmt || fmt.toLowerCase() === 'general') {
    if (v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return formatGeneral(v);
    return v;
  }

  const sections = splitSections(fmt);

  if (typeof v === 'string') {
    const textSection = sections[2]?.text ?? (sections[3] ? sections[3].numeric : undefined);
    if (textSection === undefined) return v;
    return textSection.replace(/@/g, () => v);
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';

  const n = v ?? 0;

  // The negative section is applied to the absolute value, like Excel.
  if (n < 0 && sections.length > 1) {
    return renderNumeric(Math.abs(n), sections[1].numeric);
  }
  if (n === 0 && sections.length > 2) {
    return renderNumeric(0, sections[2].numeric);
  }
  return renderNumeric(n, sections[0].numeric);
}

interface Section {
  numeric: string;
  text: string | null;
}

/** Split on unquoted, unbracketed `;` into positive/negative/zero/text. */
function splitSections(fmt: string): Section[] {
  const out: Section[] = [];
  let cur = '';
  let inQuote = false;
  let inBracket = false;

  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') inQuote = !inQuote;
    else if (ch === '\\') {
      cur += ch + (fmt[i + 1] ?? '');
      i++;
      continue;
    } else if (ch === '[') inBracket = true;
    else if (ch === ']') inBracket = false;

    if (ch === ';' && !inQuote && !inBracket) {
      out.push({ numeric: cur, text: null });
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push({ numeric: cur, text: null });

  // A fourth section holds the text format; remember it on the third entry.
  if (out.length >= 4) {
    out[2].text = out[3].numeric;
  }
  return out;
}

const COLOR_TOKEN = /\[(Red|Black|Blue|Green|White|Cyan|Magenta|Yellow)\]/gi;

/**
 * Split a numeric format into literal text and the digit placeholders.
 *
 * `#,##0`  -> prefix '', int '000', dec ''
 * `$#,##0.00` -> prefix '$', int '000', dec '00', suffix ''
 * `#,##0,` -> trailingComma 1 (divides by a thousand)
 */
interface Core {
  prefix: string;
  int: string;
  dec: string;
  trailingComma: number;
  grouped: boolean;
  suffix: string;
}

function splitCore(pattern: string): Core {
  const n = pattern.length;
  let i = 0;
  let prefix = '';

  while (i < n && !/[#0?]/.test(pattern[i])) {
    const ch = pattern[i];
    if (ch === '"') {
      const close = pattern.indexOf('"', i + 1);
      if (close < 0) {
        prefix += pattern.slice(i + 1);
        break;
      }
      prefix += pattern.slice(i + 1, close);
      i = close + 1;
    } else if (ch === '\\') {
      prefix += pattern[i + 1] ?? '';
      i += 2;
    } else if (ch === '[') {
      const close = pattern.indexOf(']', i);
      i = close < 0 ? n : close + 1;
    } else {
      prefix += ch;
      i++;
    }
  }

  let intPat = '';
  while (i < n && /[#0?,]/.test(pattern[i])) intPat += pattern[i++];

  let trailingComma = 0;
  while (intPat.endsWith(',')) {
    trailingComma++;
    intPat = intPat.slice(0, -1);
  }
  const grouped = intPat.includes(',');
  intPat = intPat.replace(/,/g, '');

  let decPat = '';
  if (pattern[i] === '.') {
    i++;
    while (i < n && /[#0?]/.test(pattern[i])) decPat += pattern[i++];
  }

  let suffix = '';
  while (i < n) {
    const ch = pattern[i];
    if (ch === '"') {
      const close = pattern.indexOf('"', i + 1);
      if (close < 0) {
        suffix += pattern.slice(i + 1);
        break;
      }
      suffix += pattern.slice(i + 1, close);
      i = close + 1;
    } else if (ch === '\\') {
      suffix += pattern[i + 1] ?? '';
      i += 2;
    } else if (ch === '[') {
      const close = pattern.indexOf(']', i);
      i = close < 0 ? n : close + 1;
    } else {
      suffix += ch;
      i++;
    }
  }

  return { prefix, int: intPat, dec: decPat, trailingComma, grouped, suffix };
}

function renderNumeric(value: number, pattern: string): string {
  if (pattern.trim() === '') return '';
  if (isDateFormat(pattern)) return renderDate(value, pattern);

  let colorName: string | null = null;
  const work = pattern.replace(COLOR_TOKEN, (_m, c: string) => {
    colorName = c;
    return '';
  });
  if (work.trim() === '') return '';

  let n = value;

  // Each `%` multiplies by 100; trailing commas scale by thousands.
  const percents = (work.match(/%/g) ?? []).length;
  if (percents > 0) n *= 100 ** percents;

  const core = splitCore(work);
  if (core.trailingComma > 0) n /= 1000 ** core.trailingComma;

  if (!Number.isFinite(n)) return '#NUM!';

  // '#' adds optional places; '0' forces them.
  const maxDecimals = Math.min(core.dec.length, 20);
  const minDecimals = (core.dec.match(/0/g) ?? []).length;
  const minIntDigits = (core.int.match(/0/g) ?? []).length;

  const negative = n < 0;
  const abs = Math.abs(n);
  const fixed = abs.toFixed(maxDecimals);
  let [intPart = '', fracPart = ''] = fixed.split('.');

  // Trim optional fraction places that came out as zeros.
  if (fracPart.length > minDecimals) {
    fracPart = fracPart.replace(new RegExp(`0{${fracPart.length - minDecimals},$`), '');
  }

  if (minIntDigits > 0) intPart = intPart.padStart(minIntDigits, '0');
  if (core.grouped) intPart = groupThousands(intPart);

  const body =
    fracPart.length > 0 ? `${intPart}.${fracPart}` : intPart;

  let out = `${negative ? '-' : ''}${core.prefix}${body}${core.suffix}`;
  if (colorName) out = colorOut(colorName, out);
  return out;
}

function groupThousands(digits: string): string {
  if (digits.length <= 3) return digits;
  const head = digits.length % 3;
  const parts = head > 0 ? [digits.slice(0, head)] : [];
  for (let i = head; i < digits.length; i += 3) parts.push(digits.slice(i, i + 3));
  return parts.join(',');
}

function colorOut(name: string, body: string): string {
  switch (name.toLowerCase()) {
    case 'red': return `\u001b[31m${body}\u001b[39m`;
    case 'blue': return `\u001b[34m${body}\u001b[39m`;
    case 'green': return `\u001b[32m${body}\u001b[39m`;
    case 'cyan': return `\u001b[36m${body}\u001b[39m`;
    case 'magenta': return `\u001b[35m${body}\u001b[39m`;
    case 'white': return `\u001b[97m${body}\u001b[39m`;
    default: return body;
  }
}
/* -------------------------------------------------------------- date output */

function renderDate(value: number, pattern: string): string {
  const has12h = /am\/pm|a\/p/i.test(pattern);
  const d = serialToDate(value);
  const H = d.getUTCHours();
  const h12 = H % 12 === 0 ? 12 : H % 12;

  let out = '';
  let i = 0;
  while (i < pattern.length) {
    const rest = pattern.slice(i);
    const m = /^(am\/pm|a\/p)/i.exec(rest);
    if (m) {
      out += H < 12 ? (m[1].length === 5 ? 'AM' : 'A') : (m[1].length === 5 ? 'PM' : 'P');
      i += m[1].length;
      continue;
    }
    const q = /^(\\.)/.exec(rest);
    if (q) {
      out += q[1];
      i += 2;
      continue;
    }
    const dq = /^"([^"]*)"/.exec(rest);
    if (dq) {
      out += dq[1];
      i += dq[0].length;
      continue;
    }
    const br = /^\[([^\]]*)\]/.exec(rest);
    if (br) {
      out += br[1].toLowerCase() === 'h' ? String(H).padStart(2, '0') : '';
      i += br[0].length;
      continue;
    }
    const y = /^(yyyy|yyy|yy|y)/.exec(rest);
    if (y) {
      const yy = d.getUTCFullYear();
      out += y[1].length >= 3 ? String(yy).padStart(4, '0') : String(yy % 100).padStart(2, '0');
      i += y[1].length;
      continue;
    }
    const mo = /^(mmmmm|mmmm|mmm|mm|m)/.exec(rest);
    if (mo) {
      const M = d.getUTCMonth();
      const len = mo[1].length;
      if (len === 1) out += String(M + 1);
      else if (len === 2) out += String(M + 1).padStart(2, '0');
      else if (len === 3) out += MONTHS[M].slice(0, 3);
      else if (len === 4) out += MONTHS[M];
      else out += MONTHS[M][0];
      i += mo[1].length;
      continue;
    }
    const dd = /^(dddd|ddd|dd|d)/.exec(rest);
    if (dd) {
      const D = d.getUTCDate();
      const w = d.getUTCDay();
      const len = dd[1].length;
      if (len === 1) out += String(D);
      else if (len === 2) out += String(D).padStart(2, '0');
      else if (len === 3) out += DAYS[w].slice(0, 3);
      else out += DAYS[w];
      i += dd[1].length;
      continue;
    }
    const hh = /^(hh|h)/.exec(rest);
    if (hh) {
      const val = has12h ? h12 : H;
      out += hh[1].length === 2 ? String(val).padStart(2, '0') : String(val);
      i += hh[1].length;
      continue;
    }
    const ss = /^(ss|s)/.exec(rest);
    if (ss) {
      out += ss[1].length === 2 ? String(d.getUTCSeconds()).padStart(2, '0') : String(d.getUTCSeconds());
      i += ss[1].length;
      continue;
    }
    if (/^\.0+/.test(rest)) {
      const zeros = /^\.(0+)/.exec(rest)![1].length;
      out += '.' + String(d.getUTCMilliseconds()).padStart(3, '0').slice(0, zeros);
      i += 1 + zeros;
      continue;
    }
    if (rest[0] === '0') {
      const frac = Math.round((value % 1) * 86400);
      const hh2 = String(Math.floor(frac / 3600)).padStart(2, '0');
      const mm2 = String(Math.floor((frac % 3600) / 60)).padStart(2, '0');
      out += `${hh2}:${mm2}`;
      i++;
      continue;
    }
    out += rest[0];
    i++;
  }
  return out;
}

/** True when a plain number should right-align by default. */
export function defaultAlign(v: Value): 'left' | 'right' | 'center' {
  if (typeof v === 'number') return 'right';
  if (typeof v === 'boolean') return 'center';
  return 'left';
}
