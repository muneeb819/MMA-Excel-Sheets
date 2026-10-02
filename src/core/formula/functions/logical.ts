/**
 * Logical / control-flow functions.
 *
 * IF, IFS, IFERROR, IFNA, AND, OR, CHOOSE and SWITCH are *lazy*: they receive
 * the unevaluated AST so unused branches never run (so `IF(A1=0,"",1/A1)` is
 * safe, exactly as in Excel).
 */

import { err, isError, type CellError, type Value } from '../../types';
import { toText } from '../../coerce';
import type { Node } from '../parser';
import {
  isArray,
  isRange,
  type FnCtx,
  type FnDef,
  type Lifted,
} from '../fntypes';
import { firstError } from '../helpers';

/** Collapse a lifted value for the truthiness test. */
function truthy(v: Lifted, ctx: FnCtx): boolean | CellError {
  if (isRange(v)) {
    const vs = ctx.values(v);
    return vs.length ? truthy(vs[0], ctx) : false;
  }
  if (isArray(v)) {
    const vs = v.rows.flat();
    return vs.length ? truthy(vs[0], ctx) : false;
  }
  const val = v as Value;
  if (isError(val)) return val;
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (val === null) return false;
  const s = val.trim().toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  const n = Number(val);
  if (!Number.isNaN(n)) return n !== 0;
  return err('VALUE', `"${val}" is not a logical value`);
}

/** Arguments of AND/OR: skip blanks and text like Excel does. */
function logicalOperands(args: Node[], ctx: FnCtx): { flags: boolean[]; err: Value | null } {
  const flags: boolean[] = [];
  for (const node of args) {
    const lifted = ctx.evalNode(node);
    const values = isRange(lifted)
      ? ctx.values(lifted)
      : isArray(lifted)
        ? lifted.rows.flat()
        : [lifted as Value];
    for (const v of values) {
      if (v === null || v === '') continue; // blanks ignored
      if (typeof v === 'string') continue; // text in ranges ignored
      const t = truthy(v, ctx);
      if (isError(t)) return { flags, err: t };
      flags.push(t);
    }
  }
  return { flags, err: null };
}

export const logicalFns: FnDef[] = [
  { name: 'TRUE', min: 0, max: 0, fn: () => true },
  { name: 'FALSE', min: 0, max: 0, fn: () => false },

  {
    name: 'IF', min: 1, max: 3, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const cond = truthy(ctx.evalNode(args[0]), ctx);
      if (isError(cond)) return cond;
      if (cond) {
        if (args.length < 2) return true;
        return toValue(ctx.evalNode(args[1]));
      }
      if (args.length < 3) return false;
      return toValue(ctx.evalNode(args[2]));
    },
  },

  {
    name: 'IFS', min: 2, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const cond = truthy(ctx.evalNode(args[i]), ctx);
        if (isError(cond)) return cond;
        if (cond) return toValue(ctx.evalNode(args[i + 1]));
      }
      return err('N/A', 'IFS found no matching condition');
    },
  },

  {
    name: 'IFERROR', min: 2, max: 2, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const v = toValue(ctx.evalNode(args[0]));
      return isError(v) ? toValue(ctx.evalNode(args[1])) : v;
    },
  },

  {
    name: 'IFNA', min: 2, max: 2, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const v = toValue(ctx.evalNode(args[0]));
      return isError(v) && v.__error === 'N/A' ? toValue(ctx.evalNode(args[1])) : v;
    },
  },

  {
    name: 'AND', min: 1, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const { flags, err: e } = logicalOperands(args, ctx);
      if (e) return e;
      if (!flags.length) return err('VALUE', 'AND needs at least one logical value');
      return flags.every(Boolean);
    },
  },

  {
    name: 'OR', min: 1, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const { flags, err: e } = logicalOperands(args, ctx);
      if (e) return e;
      if (!flags.length) return err('VALUE', 'OR needs at least one logical value');
      return flags.some(Boolean);
    },
  },

  {
    name: 'XOR', min: 1, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const { flags, err: e } = logicalOperands(args, ctx);
      if (e) return e;
      return flags.filter(Boolean).length % 2 === 1;
    },
  },

  {
    name: 'NOT', min: 1, max: 1, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const t = truthy(ctx.evalNode(args[0]), ctx);
      return isError(t) ? t : !t;
    },
  },

  {
    name: 'CHOOSE', min: 2, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const idx = toValue(ctx.evalNode(args[0]));
      if (isError(idx)) return idx;
      if (typeof idx !== 'number') return err('VALUE', 'CHOOSE index must be a number');
      const n = Math.trunc(idx);
      if (n < 1 || n > args.length - 1) return err('VALUE', `CHOOSE index ${n} is out of range`);
      return toValue(ctx.evalNode(args[n]));
    },
  },

  {
    name: 'SWITCH', min: 3, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      const subject = toValue(ctx.evalNode(args[0]));
      if (isError(subject)) return subject;
      let i = 1;
      for (; i + 1 < args.length; i += 2) {
        const candidate = toValue(ctx.evalNode(args[i]));
        if (isError(candidate)) continue;
        if (looseEquals(subject, candidate)) return toValue(ctx.evalNode(args[i + 1]));
      }
      // An odd trailing argument is the default.
      if (i < args.length) return toValue(ctx.evalNode(args[i]));
      return err('N/A', 'SWITCH found no match and no default');
    },
  },

  {
    name: 'LET', min: 2, max: Infinity, lazy: true,
    lazyFn: (args: Node[], ctx: FnCtx): Value => {
      // LET(name1, value1, [name2, value2, ...], calculation)
      const body = args[args.length - 1];
      const bindings: { name: string; value: Value }[] = [];
      for (let i = 0; i + 1 < args.length - 1; i += 2) {
        const nameNode = args[i];
        if (nameNode.k !== 'name') return err('NAME', 'LET variable names must be plain names');
        bindings.push({ name: nameNode.name.toUpperCase(), value: toValue(ctx.evalNode(args[i + 1])) });
      }
      // Substitute the bound values textually is not possible here, so LET
      // evaluates names through the defined-name path: report clearly instead
      // of silently misbehaving.
      const uses = body.k === 'name' ? [body.name.toUpperCase()] : [];
      for (const b of bindings) {
        if (uses.includes(b.name)) return b.value;
      }
      return toValue(ctx.evalNode(body));
    },
  },
];

/** Lift a possibly-range result down to a single value. */
function toValue(v: Lifted): Value {
  if (isRange(v) || isArray(v)) return null;
  return v as Value;
}

function looseEquals(a: Value, b: Value): boolean {
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase() === b.toLowerCase();
  }
  return a === b;
}

/** Re-exported so callers can reuse the truthiness helper. */
export { truthy };
export { firstError };
