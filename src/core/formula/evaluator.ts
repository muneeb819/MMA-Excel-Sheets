/**
 * The evaluator: walks an AST against a sheet, dispatching calls into the
 * function registry.
 *
 * Cell references are returned as `RangeRef` so functions can operate on whole
 * blocks; arithmetic collapses them to a single value (legacy Excel behaviour).
 */

import { err, isError, type Sheet, type Value } from '../types';
import { toNumber, toText } from '../coerce';
import { errFromLiteral, type BinOp, type Node } from './parser';
import {
  arity,
  isArray,
  isRange,
  makeArray,
  type ArrayRef,
  type FnCtx,
  type FnRegistry,
  type Lifted,
  type RangeRef,
} from './fntypes';
import { ParseError, parseFormula } from './parser';

export type EvalResult = Lifted;

/** Supplies cell values to the evaluator. */
export interface Resolver {
  /** resolve a sheet by (possibly null = current) name */
  sheet(name: string | null, current: Sheet): Sheet | null;
  cellValue(sheet: Sheet, r: number, c: number): Value;
  /** a defined name -> A1 reference text (may include a sheet prefix) */
  named(name: string): string | null;
  /** notified about every cell/range the formula reads */
  onRef?: (sheet: Sheet | null, r1: number, c1: number, r2: number, c2: number) => void;
}

export function evaluate(
  node: Node,
  sheet: Sheet,
  res: Resolver,
  fns: FnRegistry,
): EvalResult {
  const ctx = makeCtx(sheet, res, fns, '');
  return evalNode(node, sheet, res, fns, ctx);
}

function makeCtx(sheet: Sheet, res: Resolver, fns: FnRegistry, fnName: string): FnCtx {
  const ctx: FnCtx = {
    sheet,
    fnName,
    values(ref) {
      const s = ref.sheet;
      if (!s) return [];
      const out: Value[] = [];
      for (let r = ref.r1; r <= ref.r2; r++) {
        for (let c = ref.c1; c <= ref.c2; c++) out.push(res.cellValue(s, r, c));
      }
      return out;
    },
    rowsOf(ref) {
      const s = ref.sheet;
      if (!s) return [];
      const out: Value[][] = [];
      for (let r = ref.r1; r <= ref.r2; r++) {
        const row: Value[] = [];
        for (let c = ref.c1; c <= ref.c2; c++) row.push(res.cellValue(s, r, c));
        out.push(row);
      }
      return out;
    },
    shape(ref) {
      return { rows: ref.r2 - ref.r1 + 1, cols: ref.c2 - ref.c1 + 1 };
    },
    cell(s, r, c) {
      const target = s ?? sheet;
      return res.cellValue(target, r, c);
    },
    valid(s) {
      return s !== null;
    },
    evalNode(n) {
      return evalNode(n, sheet, res, fns, ctx);
    },
  };
  return ctx;
}

function evalNode(
  node: Node,
  sheet: Sheet,
  res: Resolver,
  fns: FnRegistry,
  ctx: FnCtx,
): EvalResult {
  switch (node.k) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return node.v;

    case 'ref': {
      const s = resolveSheet(node.sheet, sheet, res);
      if (!s) return err('REF', `Unknown sheet "${node.sheet ?? sheet.name}"`);
      res.onRef?.(s, node.r, node.c, node.r, node.c);
      return { __range: true, sheet: s, r1: node.r, c1: node.c, r2: node.r, c2: node.c };
    }

    case 'range': {
      const s = resolveSheet(node.sheet, sheet, res);
      if (!s) return err('REF', `Unknown sheet "${node.sheet ?? sheet.name}"`);
      res.onRef?.(s, node.r1, node.c1, node.r2, node.c2);
      return { __range: true, sheet: s, r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2 };
    }

    case 'name': {
      // A bare name may be a defined range, or TRUE/FALSE written loosely.
      const up = node.name.toUpperCase();
      if (up === 'TRUE') return true;
      if (up === 'FALSE') return false;
      const ref = res.named(node.name) ?? res.named(up);
      if (ref) {
        const expanded = expandNamedRef(ref, sheet, res);
        if (expanded) return expanded;
      }
      return err('NAME', `"${node.name}" is not a recognised name`);
    }

    case 'unary': {
      const v = scalarize(evalNode(node.a, sheet, res, fns, ctx), ctx);
      if (isError(v)) return v;
      if (node.op === '+') return v;
      const n = toNumber(v);
      if (isError(n)) return n;
      return -n;
    }

    case 'post': {
      const v = scalarize(evalNode(node.a, sheet, res, fns, ctx), ctx);
      if (isError(v)) return v;
      const n = toNumber(v);
      if (isError(n)) return n;
      return n / 100;
    }

    case 'bin': {
      const l = evalNode(node.l, sheet, res, fns, ctx);
      const r = evalNode(node.r, sheet, res, fns, ctx);
      return applyBinary(node.op, l, r, ctx);
    }

    case 'arr': {
      const rows: Value[][] = [];
      for (const row of node.rows) {
        const vals: Value[] = [];
        for (const item of row) vals.push(scalarize(evalNode(item, sheet, res, fns, ctx), ctx));
        rows.push(vals);
      }
      return makeArray(rows);
    }

    case 'call':
      return evalCall(node, sheet, res, fns, ctx);

    default:
      return err('PARSE');
  }
}

function resolveSheet(name: string | null, cur: Sheet, res: Resolver): Sheet | null {
  if (name === null) return cur;
  return res.sheet(name, cur);
}

function expandNamedRef(text: string, cur: Sheet, res: Resolver): EvalResult | null {
  const bang = text.lastIndexOf('!');
  const sheetName = bang >= 0 ? text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : null;
  const body = bang >= 0 ? text.slice(bang + 1) : text;
  const sheet = resolveSheet(sheetName, cur, res);
  if (!sheet) return null;
  const parts = body.split(':');
  const one = parseCell(parts[0]);
  if (!one) return null;
  if (parts.length === 1) {
    res.onRef?.(sheet, one.r, one.c, one.r, one.c);
    return { __range: true, sheet, r1: one.r, c1: one.c, r2: one.r, c2: one.c };
  }
  const two = parseCell(parts[1]);
  if (!two) return null;
  res.onRef?.(sheet, one.r, one.c, two.r, two.c);
  return {
    __range: true,
    sheet,
    r1: Math.min(one.r, two.r),
    c1: Math.min(one.c, two.c),
    r2: Math.max(one.r, two.r),
    c2: Math.max(one.c, two.c),
  };
}

function parseCell(text: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(text.trim());
  if (!m) return null;
  let c = 0;
  const col = m[1].toUpperCase();
  for (let i = 0; i < col.length; i++) c = c * 26 + (col.charCodeAt(i) - 64);
  return { r: parseInt(m[2], 10) - 1, c: c - 1 };
}

/** Collapse a range/array to a single value (implicit intersection). */
function scalarize(v: EvalResult, ctx: FnCtx): Value {
  if (isRange(v)) {
    const vals = ctx.values(v);
    return vals.length ? vals[0] : null;
  }
  if (isArray(v)) {
    const vals = v.rows.flat();
    return vals.length ? vals[0] : null;
  }
  return v;
}

/* ------------------------------------------------------------- arithmetic */

function applyBinary(op: BinOp, l: EvalResult, r: EvalResult, ctx: FnCtx): EvalResult {
  if (op === '&') {
    const lv = scalarize(l, ctx);
    if (isError(lv)) return lv;
    const rv = scalarize(r, ctx);
    if (isError(rv)) return rv;
    return toText(lv) + toText(rv);
  }

  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    return compare(op, l, r, ctx);
  }

  const lv = scalarize(l, ctx);
  if (isError(lv)) return lv;
  const rv = scalarize(r, ctx);
  if (isError(rv)) return rv;

  const a = toNumber(lv);
  if (isError(a)) return a;
  const b = toNumber(rv);
  if (isError(b)) return b;

  switch (op) {
    case '+': return a + b;
    case '-': return a - b;
    case '*': return a * b;
    case '/':
      if (b === 0) return err('DIV/0', 'Division by zero');
      return a / b;
    case '^': {
      const out = Math.pow(a, b);
      return Number.isFinite(out) ? out : err('NUM', 'Power produced a non-finite result');
    }
    default:
      return err('PARSE');
  }
}

function compare(op: BinOp, l: EvalResult, r: EvalResult, ctx: FnCtx): Value {
  const lv = scalarize(l, ctx);
  const rv = scalarize(r, ctx);
  if (isError(lv)) return lv;
  if (isError(rv)) return rv;

  // blank compares as 0 or "" depending on the other side
  const lx: Value = lv === null ? (typeof rv === 'number' ? 0 : typeof rv === 'boolean' ? false : '') : lv;
  const rx: Value = rv === null ? (typeof lv === 'number' ? 0 : typeof lv === 'boolean' ? false : '') : rv;

  if (typeof lx === 'string' && typeof rx === 'string') {
    const cmp = lx.localeCompare(rx, undefined, { sensitivity: 'base' });
    return testCmp(op, cmp);
  }
  if (typeof lx === 'boolean' || typeof rx === 'boolean') {
    const a = typeof lx === 'boolean' ? (lx ? 1 : 0) : typeof lx === 'number' ? lx : null;
    const b = typeof rx === 'boolean' ? (rx ? 1 : 0) : typeof rx === 'number' ? rx : null;
    if (a === null || b === null) return err('VALUE', 'Cannot compare text with a logical value');
    return testCmp(op, a === b ? 0 : a < b ? -1 : 1);
  }
  const a = toNumber(lx);
  if (isError(a)) return a;
  const b = toNumber(rx);
  if (isError(b)) return b;
  return testCmp(op, a === b ? 0 : a < b ? -1 : 1);
}

function testCmp(op: BinOp, cmp: number): boolean {
  switch (op) {
    case '=': return cmp === 0;
    case '<>': return cmp !== 0;
    case '<': return cmp < 0;
    case '>': return cmp > 0;
    case '<=': return cmp <= 0;
    case '>=': return cmp >= 0;
    default: return false;
  }
}

/* ------------------------------------------------------------- call + lazy */

function evalCall(
  node: Extract<Node, { k: 'call' }>,
  sheet: Sheet,
  res: Resolver,
  fns: FnRegistry,
  ctx: FnCtx,
): EvalResult {
  const def = fns.get(node.name);
  if (!def) return err('NAME', `${node.name} is not a function`);
  const inner = makeCtx(sheet, res, fns, def.name);

  if (def.lazy) {
    const lazyImpl = def.lazyFn;
    if (!lazyImpl) return err('VALUE', `${def.name} has no lazy implementation`);
    try {
      return lazyImpl(node.args, inner);
    } catch (e) {
      return err('VALUE', (e as Error).message);
    }
  }

  const impl = def.fn;
  if (!impl) return err('VALUE', `${def.name} has no implementation`);
  const lifted: Lifted[] = [];
  for (const a of node.args) lifted.push(evalNode(a, sheet, res, fns, inner));

  const { min, max } = arity(def);
  if (lifted.length < min || lifted.length > max) {
    return err('VALUE', `${def.name} expects ${min === max ? min : `${min} to ${max}`} arguments, got ${lifted.length}`);
  }

  try {
    const out = impl(lifted, inner);
    if (isArray(out)) return out;
    return out;
  } catch (e) {
    if (e instanceof ParseError) return err('PARSE', e.message);
    return err('VALUE', (e as Error).message);
  }
}

/**
 * Evaluate a formula body, collapsing ranges and arrays to a single value.
 * Use `evalFormulaRaw` when the caller needs to spill an array result.
 */
export function evalFormula(
  formula: string,
  sheet: Sheet,
  res: Resolver,
  fns: FnRegistry,
): Value {
  const out = evalFormulaRaw(formula, sheet, res, fns);
  return scalarize(out, makeCtx(sheet, res, fns, ''));
}

/** Parse + evaluate without collapsing an array result. Never throws. */
export function evalFormulaRaw(
  formula: string,
  sheet: Sheet,
  res: Resolver,
  fns: FnRegistry,
): EvalResult {
  let ast: Node;
  try {
    ast = parseFormula(formula);
  } catch (e) {
    return err('PARSE', (e as Error).message);
  }
  try {
    return evaluate(ast, sheet, res, fns);
  } catch (e) {
    if (e instanceof ParseError) return err('PARSE', e.message);
    return err('VALUE', (e as Error).message);
  }
}

export type { ArrayRef, RangeRef };
export { errFromLiteral };
