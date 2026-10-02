/**
 * A1-notation helpers and range packing.
 *
 * Cell coordinates are 0-based internally (r0c0 === A1). Everything the user
 * sees goes through `colName` / `a1` / `parseA1` in the opposite direction.
 */

import { MAX_COLS, type Range } from './types';

/** Pack a coordinate into a single integer so cells fit in a sparse Map. */
export function pack(r: number, c: number): number {
  return r * MAX_COLS + c;
}
export function unpackRow(k: number): number {
  return Math.floor(k / MAX_COLS);
}
export function unpackCol(k: number): number {
  return k % MAX_COLS;
}

/** 0 -> A, 25 -> Z, 26 -> AA */
export function colName(c: number): string {
  let n = c + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 'A' -> 0, 'AA' -> 26 */
export function colIndex(name: string): number {
  let n = 0;
  const up = name.toUpperCase();
  for (let i = 0; i < up.length; i++) {
    const ch = up.charCodeAt(i);
    if (ch < 65 || ch > 90) return -1;
    n = n * 26 + (ch - 64);
  }
  return n - 1;
}

/** A1 -> {r,c} */
export function parseA1(a1: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(a1.trim());
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = parseInt(m[2], 10) - 1;
  if (c < 0 || r < 0) return null;
  return { r, c };
}

export function a1(r: number, c: number, absolute = false): string {
  return `${absolute ? '$' : ''}${colName(c)}${absolute ? '$' : ''}${r + 1}`;
}

/** Single cell to absolute A1, e.g. `$B$7` */
export function a1abs(r: number, c: number): string {
  return a1(r, c, true);
}

/** Range to `A1:C9`, collapsing to one cell when r1===r2 && c1===c2. */
export function rangeToA1(r: Range, absolute = false): string {
  const a = a1(r.r1, r.c1, absolute);
  if (r.r1 === r.r2 && r.c1 === r.c2) return a;
  const b = a1(r.r2, r.c2, absolute);
  return `${a}:${b}`;
}

export function normRange(r: Range): Range {
  return {
    r1: Math.min(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2),
    c2: Math.max(r.c1, r.c2),
  };
}

export function rangeContains(r: Range, rr: number, cc: number): boolean {
  const n = normRange(r);
  return rr >= n.r1 && rr <= n.r2 && cc >= n.c1 && cc <= n.c2;
}

export function rangeEquals(a: Range, b: Range): boolean {
  const x = normRange(a);
  const y = normRange(b);
  return x.r1 === y.r1 && x.r2 === y.r2 && x.c1 === y.c1 && x.c2 === y.c2;
}

export function rangeCells(r: Range): number {
  const n = normRange(r);
  return (n.r2 - n.r1 + 1) * (n.c2 - n.c1 + 1);
}

export function* iterRange(r: Range): Generator<{ r: number; c: number }> {
  const n = normRange(r);
  for (let rr = n.r1; rr <= n.r2; rr++) {
    for (let cc = n.c1; cc <= n.c2; cc++) yield { r: rr, c: cc };
  }
}

export function union(a: Range, b: Range): Range {
  const x = normRange(a);
  const y = normRange(b);
  return {
    r1: Math.min(x.r1, y.r1),
    c1: Math.min(x.c1, y.c1),
    r2: Math.max(x.r2, y.r2),
    c2: Math.max(x.c2, y.c2),
  };
}

/**
 * Split `Sheet1!$A$1` into its parts. Sheet name may itself be quoted
 * ('My Sheet'!A1:B2) and may contain `!` when escaped with ''.
 */
export function splitRef(
  ref: string,
): { sheet: string | null; a1: string } | null {
  const s = ref.trim();
  if (!s) return null;
  const bang = findUnquotedBang(s);
  if (bang < 0) return { sheet: null, a1: s };
  const sheetRaw = s.slice(0, bang);
  let sheet = sheetRaw.trim();
  if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, a1: s.slice(bang + 1) };
}

function findUnquotedBang(s: string): number {
  let inQuote = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'") {
      if (inQuote && s[i + 1] === "'") {
        i++;
        continue;
      }
      inQuote = !inQuote;
    } else if (ch === '!' && !inQuote) return i;
  }
  return -1;
}

/** Quote a sheet name for use in a formula when required. */
export function quoteSheet(name: string): string {
  const bare = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name);
  return bare ? name : `'${name.replace(/'/g, "''")}'`;
}

/**
 * Parse a reference expression such as `A1`, `$A$1:$B$9`, `Sheet1!A1:B2`.
 * Returns null when the text is not a reference.
 */
export function parseRangeExpr(expr: string): (Range & { sheet: string | null }) | null {
  const sp = splitRef(expr);
  if (!sp) return null;
  const parts = sp.a1.split(':');
  if (parts.length === 1) {
    const p = parseA1(parts[0]);
    if (!p) return null;
    return { r1: p.r, c1: p.c, r2: p.r, c2: p.c, sheet: sp.sheet };
  }
  if (parts.length !== 2) return null;
  const a = parseA1(parts[0]);
  const b = parseA1(parts[1]);
  if (!a || !b) return null;
  return { r1: a.r, c1: a.c, r2: b.r, c2: b.c, sheet: sp.sheet };
}

/** Human readable range size, e.g. `12R x 3C` (used by the status bar). */
export function describeRange(r: Range): string {
  const n = normRange(r);
  return `${n.r2 - n.r1 + 1}R x ${n.c2 - n.c1 + 1}C`;
}

/** Quote a literal string for display in the formula bar / in cells. */
export function quoteString(s: string): string {
  return `"${s.replace(/"/g, '""')}"`;
}

/**
 * Shift the relative parts of every cell reference in a formula.
 *
 * Used by the fill handle and by copy/paste, so `A1+1` copied down one row
 * becomes `A2+1` while `$A$1` and `A$1` stay put. String literals are left
 * alone, and a reference pushed off the grid becomes `#REF!`.
 */
export function translateFormula(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula;

  let out = '';
  let i = 0;
  const n = formula.length;

  while (i < n) {
    const ch = formula[i];

    // String literals are copied verbatim.
    if (ch === '"') {
      let j = i + 1;
      while (j < n) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      out += formula.slice(i, Math.min(j + 1, n));
      i = j + 1;
      continue;
    }

    // Quoted sheet names followed by `!` are passed through.
    if (ch === "'") {
      let j = i + 1;
      while (j < n && formula[j] !== "'") j++;
      if (formula[j + 1] === '!') {
        out += formula.slice(i, j + 2);
        i = j + 2;
        continue;
      }
      out += ch;
      i++;
      continue;
    }

    // Bare sheet names: consume `Name!` so the reference after it is shifted.
    if (/[A-Za-z_\\]/.test(ch)) {
      let j = i;
      while (j < n && /[A-Za-z0-9_.\\$]/.test(formula[j])) j++;
      const word = formula.slice(i, j);
      const afterBang = formula[j] === '!' ? formula.slice(j + 1) : '';
      if (formula[j] === '!' && /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?![A-Za-z0-9_])/.test(afterBang)) {
        // Keep the sheet prefix and shift the reference that follows it.
        out += word + '!';
        i = j + 1;
        continue;
      }
      // Otherwise `word` may itself be a cell reference such as A1 or $B$7.
      if (/^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(word)) {
        out += shiftRefToken(word, dr, dc);
        i = j;
        continue;
      }
      out += word;
      i = j;
      continue;
    }

    const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?![A-Za-z0-9_])/.exec(formula.slice(i));
    if (m) {
      out += shiftRefToken(m[0], dr, dc);
      i += m[0].length;
      continue;
    }

    out += ch;
    i++;
  }

  return out;
}

function shiftRefToken(token: string, dr: number, dc: number): string {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/.exec(token)!;
  const colAbs = m[1] === '$';
  const rowAbs = m[3] === '$';
  const c = colIndex(m[2]);
  const r = parseInt(m[4], 10) - 1;
  const nc = colAbs ? c : c + dc;
  const nr = rowAbs ? r : r + dr;
  if (nc < 0 || nr < 0) return '#REF!';
  return `${colAbs ? '$' : ''}${colName(nc)}${rowAbs ? '$' : ''}${nr + 1}`;
}
