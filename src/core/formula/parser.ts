/**
 * Formula tokenizer and recursive-descent parser.
 *
 * Produces an AST that `evaluator.ts` walks. Precedence and associativity match
 * Excel (so `-2^2` is 4, not -4, and `^` is left-associative).
 */

import { err, type CellError } from '../types';

/* -------------------------------------------------------------------- AST */

export type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; v: CellError }
  | { k: 'ref'; sheet: string | null; r: number; c: number }
  | { k: 'range'; sheet: string | null; r1: number; c1: number; r2: number; c2: number }
  | { k: 'name'; name: string }
  | { k: 'call'; name: string; args: Node[] }
  | { k: 'bin'; op: BinOp; l: Node; r: Node }
  | { k: 'unary'; op: '-' | '+'; a: Node }
  | { k: 'post'; op: '%'; a: Node }
  | { k: 'arr'; rows: Node[][] };

export type BinOp =
  | '+' | '-' | '*' | '/' | '^' | '&'
  | '=' | '<>' | '<' | '>' | '<=' | '>=';

/* ------------------------------------------------------------------ lexer */

type TokKind =
  | 'num' | 'str' | 'bool' | 'err' | 'ident' | 'sheet'
  | 'op' | '(' | ')' | '{' | '}' | ',' | ';' | ':' | 'eof';

interface Tok {
  t: TokKind;
  /** token text, or decoded literal for num/str/bool */
  v: string;
  num?: number;
  pos: number;
}

const ERR_LITERALS = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A',
]);

export function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  const n = src.length;

  while (i < n) {
    const ch = src[i];

    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++;
      continue;
    }

    // Quoted sheet name: 'My Sheet'!A1
    if (ch === "'") {
      let j = i + 1;
      let name = '';
      while (j < n) {
        if (src[j] === "'") {
          if (src[j + 1] === "'") {
            name += "'";
            j += 2;
            continue;
          }
          break;
        }
        name += src[j];
        j++;
      }
      if (src[j] === "'" && src[j + 1] === '!') {
        out.push({ t: 'sheet', v: name, pos: i });
        i = j + 2;
        continue;
      }
      // otherwise fall through and treat as an error
      out.push({ t: 'err', v: '#VALUE!', pos: i });
      i++;
      continue;
    }

    // Error literals
    if (ch === '#') {
      const rest = src.slice(i).toUpperCase();
      const hit = [...ERR_LITERALS].find((e) => rest.startsWith(e));
      if (hit) {
        out.push({ t: 'err', v: hit, pos: i });
        i += hit.length;
        continue;
      }
      out.push({ t: 'err', v: '#ERROR!', pos: i });
      i++;
      continue;
    }

    // String literal
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j];
        j++;
      }
      out.push({ t: 'str', v: s, pos: i });
      i = j + 1;
      continue;
    }

    // Number
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      let j = i;
      while (j < n && /[0-9]/.test(src[j])) j++;
      if (src[j] === '.') {
        j++;
        while (j < n && /[0-9]/.test(src[j])) j++;
      }
      if (src[j] === 'e' || src[j] === 'E') {
        const save = j;
        j++;
        if (src[j] === '+' || src[j] === '-') j++;
        if (/[0-9]/.test(src[j] ?? '')) {
          while (j < n && /[0-9]/.test(src[j])) j++;
        } else j = save;
      }
      const text = src.slice(i, j);
      out.push({ t: 'num', v: text, num: parseFloat(text), pos: i });
      i = j;
      continue;
    }

    // Identifier: function names, names, cell refs, boolean literals
    if (/[A-Za-z_\\$]/.test(ch)) {
      let j = i;
      while (j < n && /[A-Za-z0-9_.$\\]/.test(src[j])) j++;
      // allow sheet prefix in names like Sheet1!A1 handled by the parser via '!'
      const text = src.slice(i, j);
      const up = text.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') {
        out.push({ t: 'bool', v: up, pos: i });
      } else {
        out.push({ t: 'ident', v: text, pos: i });
      }
      i = j;
      continue;
    }

    // Multi-character operators
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ t: 'op', v: two, pos: i });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>:%!'.includes(ch)) {
      out.push({ t: 'op', v: ch, pos: i });
      i++;
      continue;
    }
    if (ch === '(') { out.push({ t: '(', v: ch, pos: i }); i++; continue; }
    if (ch === ')') { out.push({ t: ')', v: ch, pos: i }); i++; continue; }
    if (ch === '{') { out.push({ t: '{', v: ch, pos: i }); i++; continue; }
    if (ch === '}') { out.push({ t: '}', v: ch, pos: i }); i++; continue; }
    if (ch === ',') { out.push({ t: ',', v: ch, pos: i }); i++; continue; }
    if (ch === ';') { out.push({ t: ';', v: ch, pos: i }); i++; continue; }
    if (ch === ':') { out.push({ t: ':', v: ch, pos: i }); i++; continue; }
    if (ch === ',') { out.push({ t: ',', v: ch, pos: i }); i++; continue; }

    // Unknown character
    out.push({ t: 'err', v: '#ERROR!', pos: i });
    i++;
  }

  out.push({ t: 'eof', v: '', pos: n });
  return out;
}

/* ----------------------------------------------------------------- parser */

const PREC: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

function parseCellRef(text: string): { r: number; c: number } | null {
  const m = CELL_RE.exec(text);
  if (!m) return null;
  let c = 0;
  const colPart = m[2].toUpperCase();
  for (let i = 0; i < colPart.length; i++) c = c * 26 + (colPart.charCodeAt(i) - 64);
  c -= 1;
  const r = parseInt(m[4], 10) - 1;
  if (r < 0 || c < 0) return null;
  return { r, c };
}

class Parser {
  private toks: Tok[];
  private i = 0;

  constructor(src: string) {
    this.toks = tokenize(src);
  }

  private peek(off = 0): Tok {
    return this.toks[Math.min(this.i + off, this.toks.length - 1)];
  }
  private next(): Tok {
    return this.toks[this.i++];
  }
  private atOp(v: string): boolean {
    const t = this.peek();
    return t.t === 'op' && t.v === v;
  }
  private eat(v: string): boolean {
    if (this.atOp(v)) {
      this.i++;
      return true;
    }
    return false;
  }

  parse(): Node {
    const node = this.expr(0);
    const t = this.peek();
    if (t.t !== 'eof') throw new ParseError(`Unexpected "${t.v}"`, t.pos);
    return node;
  }

  private expr(minPrec: number): Node {
    let left = this.unary();
    for (;;) {
      const t = this.peek();
      if (t.t !== 'op') break;
      const p = PREC[t.v];
      if (p === undefined || p < minPrec) break;
      this.next();
      // '^' is left-associative in Excel
      const right = this.expr(p + 1);
      left = { k: 'bin', op: t.v as BinOp, l: left, r: right };
    }
    return left;
  }

  private unary(): Node {
    if (this.atOp('-')) {
      this.next();
      return { k: 'unary', op: '-', a: this.unary() };
    }
    if (this.atOp('+')) {
      this.next();
      return { k: 'unary', op: '+', a: this.unary() };
    }
    return this.postfix();
  }

  private postfix(): Node {
    let node = this.primary();
    while (this.atOp('%')) {
      this.next();
      node = { k: 'post', op: '%', a: node };
    }
    return node;
  }

  private primary(): Node {
    const t = this.next();

    if (t.t === 'num') return { k: 'num', v: t.num! };
    if (t.t === 'str') return { k: 'str', v: t.v };
    if (t.t === 'bool') return { k: 'bool', v: t.v === 'TRUE' };
    if (t.t === 'err') return { k: 'err', v: errFromLiteral(t.v) };

    if (t.t === '(') {
      const inner = this.expr(0);
      if (this.peek().t === ')') this.next();
      return inner;
    }

    if (t.t === '{') {
      const rows: Node[][] = [];
      let row: Node[] = [];
      for (;;) {
        row.push(this.expr(0));
        const nx = this.peek();
        if (nx.t === ',') { this.next(); continue; }
        if (nx.t === ';') { this.next(); rows.push(row); row = []; continue; }
        if (nx.t === '}') { this.next(); break; }
        throw new ParseError('Malformed array constant', nx.pos);
      }
      rows.push(row);
      return { k: 'arr', rows };
    }

    if (t.t === 'sheet') {
      const a1 = this.expectRefPart();
      return this.finishRef(t.v, a1);
    }

    if (t.t === 'ident') {
      const up = t.v.toUpperCase();

      // Sheet1!A1
      if (this.atOp('!')) {
        this.next();
        const a1 = this.expectRefPart();
        return this.finishRef(t.v, a1);
      }

      // function call
      if (this.peek().t === '(') {
        this.next(); // consume '('
        const args: Node[] = [];
        if (this.peek().t === ')') {
          this.next(); // zero-argument call
        } else {
          for (;;) {
            if (this.peek().t === ',') {
              // an omitted argument, e.g. ROUND(A1,,2)
              this.next();
              args.push({ k: 'str', v: '' });
              continue;
            }
            args.push(this.expr(0));
            const sep = this.peek();
            if (sep.t === ',') {
              this.next();
              continue;
            }
            if (sep.t === ')') {
              this.next(); // consume the closing paren of THIS call
              break;
            }
            throw new ParseError('Malformed function call', sep.pos);
          }
        }
        return { k: 'call', name: up, args };
      }

      // plain cell / range reference
      const first = parseCellRef(t.v);
      if (first) {
        if (this.atOp(':')) {
          this.next();
          const a2 = this.expectRefPart();
          const second = parseCellRef(a2);
          if (second) {
            return {
              k: 'range',
              sheet: null,
              r1: Math.min(first.r, second.r),
              c1: Math.min(first.c, second.c),
              r2: Math.max(first.r, second.r),
              c2: Math.max(first.c, second.c),
            };
          }
        }
        return { k: 'ref', sheet: null, r: first.r, c: first.c };
      }

      return { k: 'name', name: t.v };
    }

    throw new ParseError(`Unexpected "${t.v}"`, t.pos);
  }

  /** After a `!` we must see A1 or B2:C3 (possibly sheet-qualified again). */
  private expectRefPart(): string {
    const t = this.next();
    if (t.t === 'ident') return t.v;
    if (t.t === 'sheet') {
      // nested: Sheet1!Sheet2!A1 is invalid, report cleanly
      throw new ParseError('Invalid reference', t.pos);
    }
    throw new ParseError('Expected a cell reference', t.pos);
  }

  private finishRef(sheet: string, a1: string): Node {
    const parts = a1.split(':');
    const first = parseCellRef(parts[0]);
    if (!first) throw new ParseError(`Invalid reference "${sheet}!${a1}"`, 0);
    if (parts.length === 2) {
      const second = parseCellRef(parts[1]);
      if (!second) throw new ParseError(`Invalid reference "${sheet}!${a1}"`, 0);
      return {
        k: 'range',
        sheet,
        r1: Math.min(first.r, second.r),
        c1: Math.min(first.c, second.c),
        r2: Math.max(first.r, second.r),
        c2: Math.max(first.c, second.c),
      };
    }
    return { k: 'ref', sheet, r: first.r, c: first.c };
  }
}

export class ParseError extends Error {
  constructor(message: string, public pos: number) {
    super(message);
    this.name = 'ParseError';
  }
}

export function errFromLiteral(text: string): CellError {
  const map: Record<string, CellError['__error']> = {
    '#NULL!': 'NULL',
    '#DIV/0!': 'DIV/0',
    '#VALUE!': 'VALUE',
    '#REF!': 'REF',
    '#NAME?': 'NAME',
    '#NUM!': 'NUM',
    '#N/A': 'N/A',
  };
  return err(map[text.toUpperCase()] ?? 'PARSE');
}

/** Parse a formula body (no leading '='). Throws ParseError. */
export function parseFormula(src: string): Node {
  return new Parser(src).parse();
}
