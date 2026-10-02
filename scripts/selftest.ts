/**
 * Engine self-test.
 *
 * Exercises the formula engine end to end so behaviour is verified against
 * known Excel results rather than assumed. Run with `npm run selftest`.
 */

import { Engine, newWorkbook } from '../src/core/engine';
import { isError } from '../src/core/types';
import { formatValue } from '../src/core/coerce';
import { translateFormula } from '../src/core/ref';
import type { Value } from '../src/core/types';

let pass = 0;
let fail = 0;
const failures: string[] = [];

/**
 * Scratch cell for the formula under test. It sits well away from the A1:D4
 * block the fixtures populate so a formula can never reference itself by
 * accident.
 */
const SCRATCH = { r: 9, c: 4 };

/** Evaluate a formula in a scratch cell and compare against `expected`. */
function check(label: string, formula: string, expected: Value, setup?: (e: Engine) => void): void {
  let got: Value;
  try {
    const engine = new Engine(newWorkbook());
    const sheet = engine.active;
    if (setup) setup(engine);
    engine.setInput(sheet, SCRATCH.r, SCRATCH.c, formula.startsWith('=') ? formula : `=${formula}`);
    engine.recalc();
    got = engine.valueAt(sheet, SCRATCH.r, SCRATCH.c);
  } catch (e) {
    fail++;
    failures.push(`${label}\n    ${formula}\n    threw ${(e as Error).message}`);
    return;
  }
  if (same(got, expected)) {
    pass++;
  } else {
    fail++;
    failures.push(`${label}\n    ${formula}\n    expected ${show(expected)}\n    got      ${show(got)}`);
  }
}

function same(a: Value, b: Value): boolean {
  if (isError(a) || isError(b)) {
    return isError(a) && isError(b) && a.__error === b.__error;
  }
  if (typeof a === 'number' && typeof b === 'number') {
    return Math.abs(a - b) < 1e-6 * Math.max(1, Math.abs(b));
  }
  return a === b;
}

/** Plain assertion for helpers that are not formulas. */
function eq<T>(label: string, got: T, expected: T): void {
  if (got === expected) {
    pass++;
  } else {
    fail++;
    failures.push(`${label}\n    expected ${String(expected)}\n    got      ${String(got)}`);
  }
}

function show(v: Value): string {
  if (isError(v)) return `${v.__error}${v.message ? ` (${v.message})` : ''}`;
  if (v === null) return '<blank>';
  return formatValue(v);
}

/* --------------------------------------------------------------- arithmetic */

check('precedence', '1+2*3', 7);
check('power', '2^3', 8);
check('excel unary vs power', '-2^2', 4);
check('power left assoc', '2^3^2', 64);
check('paren', '(1+2)*3', 9);
check('percent postfix', '50%', 0.5);
check('unary minus percent', '-50%', -0.5);
check('concat', '"a"&"b"&1', 'ab1');
check('comparison', '2>1', true);
check('text compare', '"A"="a"', true);
check('div by zero', '1/0', { __error: 'DIV/0' });
check('mod sign', 'MOD(-3,2)', 1);
check('rounding', 'ROUND(2.675,2)', 2.68);
check('floor neg', 'FLOOR(-2.5,1)', -3);
check('int neg', 'INT(-2.5)', -3);
check('abs', 'ABS(-7)', 7);
check('power err', '(-8)^(1/3)', { __error: 'NUM' });

/* ---------------------------------------------------------------- references */

/** Put input data in B/D columns so the scratch cell (A1) never self-references. */
check('cell ref', 'B1+1', 11, (e) => e.setInput(e.active, 0, 1, '10'));
check('range sum', 'SUM(B1:B3)', 6, (e) => {
  e.setInput(e.active, 0, 1, '1');
  e.setInput(e.active, 1, 1, '2');
  e.setInput(e.active, 2, 1, '3');
});
check('cross sheet', 'S2!A1*2', 20, (e) => {
  const s2 = e.addSheet('S2');
  e.setInput(s2, 0, 0, '10');
  e.wb.activeSheet = 0;
  e.invalidateAll();
});
check('recalc chain', 'B1*B1', 100, (e) => e.setInput(e.active, 0, 1, '10'));
check('circular ref', '=F11', { __error: 'CYCLE' }, (e) => e.setInput(e.active, 10, 5, '=E10'));
check('indirect cycle', '=F12', { __error: 'CYCLE' }, (e) => {
  e.setInput(e.active, 11, 5, '=G12');
  e.setInput(e.active, 11, 6, '=F12');
});
check('transitive chain', 'SUM(B1:B3)', 7, (e) => {
  e.setInput(e.active, 0, 1, '1');
  e.setInput(e.active, 1, 1, '=B1*2');
  e.setInput(e.active, 2, 1, '=B2*2');
});

/* ----------------------------------------------------------------- logical */

check('if true', 'IF(1>0,"yes","no")', 'yes');
check('if picks true branch', 'IF(B1=0,"",1/B1)', '', (e) => e.setInput(e.active, 0, 1, '0'));
check('if avoids div0', 'IF(B2=0,"safe",1/B2)', 'safe', (e) => e.setInput(e.active, 1, 1, '0'));
check('ifs', 'IFS(FALSE,1,TRUE,2)', 2);
check('iferror', 'IFERROR(1/0,"caught")', 'caught');
check('ifna', 'IFNA(1/0,"na")', { __error: 'DIV/0' });
check('and or', 'AND(TRUE,1>0)', true);
check('switch', 'SWITCH(3,1,"a",3,"c","z")', 'c');

/* -------------------------------------------------------------- statistics */

check('sum', 'SUM(1,2,3)', 6);
check('average', 'AVERAGE(2,4,6)', 4);
check('average ignores text', 'AVERAGE(1,"x",3)', 2, undefined);
check('count', 'COUNT(1,"x",3)', 2);
check('counta', 'COUNTA(1,"x",3)', 3);
check('counta blank', 'COUNTA(B5:B9)', 1, (e) => {
  e.setInput(e.active, 4, 1, 'x');
  e.setInput(e.active, 5, 1, '');
});
check('median', 'MEDIAN(1,3,2)', 2);
check('mode', 'MODE(1,2,2,3)', 2);
check('stdev.s', 'STDEV.S(2,4,4,4,5,5,7,9)', 2.13809);
check('var.p', 'VAR.P(1,2,3,4)', 1.25);
check('large', 'LARGE({1;2;3;4;5},2)', 4);
check('rank ascending', 'RANK(3,{1;2;3;4})', 3);
check('percentile', 'PERCENTILE({1;2;3;4},0.5)', 2.5);
check('correl', 'CORREL({1;2;3},{2;4;6})', 1);
check('slope', 'SLOPE({2;4;6},{1;2;3})', 2);
check('norm dist', 'NORM.DIST(0,0,1,TRUE)', 0.5);
check('norm inv', 'ROUND(NORM.INV(0.975,0,1),4)', 1.96);
check('binom', 'BINOM.DIST(2,5,0.5,FALSE)', 0.3125);

/* -------------------------------------------------------------------- text */

check('concat fn', 'CONCATENATE("a",1,"b")', 'a1b');
check('textjoin', 'TEXTJOIN("-",TRUE,"a","b","c")', 'a-b-c');
check('left', 'LEFT("hello",2)', 'he');
check('mid', 'MID("hello",2,3)', 'ell');
check('len', 'LEN("hello")', 5);
check('trim', 'TRIM("  a   b  ")', 'a b');
check('proper', 'PROPER("hello world")', 'Hello World');
check('substitute', 'SUBSTITUTE("a-a-a","a","b",2)', 'a-b-a');
check('replace', 'REPLACE("abcdef",2,3,"X")', 'aXef');
check('find case sensitive', 'FIND("B","abc")', { __error: 'VALUE' });
check('search insensitive', 'SEARCH("B","abc")', 2);
check('repeat', 'REPT("ab",3)', 'ababab');
check('value', 'VALUE("1,234.5")', 1234.5);
check('text fmt', 'TEXT(1234.567,"#,##0.00")', '1,234.57');
check('text pct', 'TEXT(0.25,"0%")', '25%');
check('exact', 'EXACT("a","A")', false);
check('t fn', 'T("x")', 'x');
check('t fn on number', 'T(5)', '');
check('regexextract', 'REGEXEXTRACT("abc123","[0-9]+")', '123');

/* ------------------------------------------------------------------ lookup */

const table = (e: Engine): void => {
  e.setInput(e.active, 0, 0, 'Name');
  e.setInput(e.active, 0, 1, 'Dept');
  e.setInput(e.active, 0, 2, 'Age');
  e.setInput(e.active, 1, 0, 'Ana');
  e.setInput(e.active, 1, 1, 'Ops');
  e.setInput(e.active, 1, 2, '31');
  e.setInput(e.active, 2, 0, 'Ben');
  e.setInput(e.active, 2, 1, 'Eng');
  e.setInput(e.active, 2, 2, '45');
  e.setInput(e.active, 3, 0, 'Cara');
  e.setInput(e.active, 3, 1, 'Ops');
  e.setInput(e.active, 3, 2, '27');
};

check('vlookup exact', 'VLOOKUP("Ben",A1:C4,3,FALSE)', 45, table);
check('vlookup case insensitive', 'VLOOKUP("ben",A1:C4,2,FALSE)', 'Eng', table);
check('vlookup approx', 'VLOOKUP("Bz",A2:A4,1)', 'Ben', table);
check('vlookup not found', 'VLOOKUP("Zed",A1:C4,2,FALSE)', { __error: 'N/A' }, table);
check('hlookup', 'HLOOKUP("Age",A1:C4,4,FALSE)', 27, table);
check('xlookup', 'XLOOKUP("Cara",A1:A4,C1:C4)', 27, table);
check('xlookup if not found', 'XLOOKUP("Zed",A1:A4,C1:C4,"none")', 'none', table);
check('match exact', 'MATCH("Eng",B1:B4,0)', 3, table);
check('index', 'INDEX(A1:C4,3,3)', 45, table);
check('index row', 'INDEX(A1:A4,4)', 'Cara', table);
check('offset', 'SUM(OFFSET(C1,0,0,4,1))', 103, table);
check('rows cols', 'ROWS(A1:C3)+COLUMNS(A1:C3)', 6, table);
check('indirect', 'INDIRECT("C2")', 31, table);
check('address', 'ADDRESS(2,3)', '$C$2');
check('choose', 'CHOOSE(2,"a","b","c")', 'b');

/* ------------------------------------------------------------- conditional */

check('sumif', 'SUMIF(B2:B4,"Ops",C2:C4)', 58, table);
check('sumif wildcard', 'SUMIF(A2:A4,"B*",C2:C4)', 45, table);
check('sumifs', 'SUMIFS(C2:C4,B2:B4,"Ops",C2:C4,">28")', 31, table);
check('countif', 'COUNTIF(B2:B4,"Ops")', 2, table);
check('countifs', 'COUNTIFS(C2:C4,">30")', 2, table);
check('averageif', 'AVERAGEIF(B2:B4,"Ops",C2:C4)', 29, table);
check('maxifs', 'MAXIFS(C2:C4,B2:B4,"Ops")', 31, table);
check('sumproduct', 'SUMPRODUCT({1;2;3},{4;5;6})', 32);

/* -------------------------------------------------------------- financial */

check('pmt', 'ROUND(PMT(0.05/12,60,-20000),2)', 377.42);
check('fv', 'ROUND(FV(0.06/12,10,-200,-500,1),2)', 2581.4);
check('pv', 'ROUND(PV(0.08,10,-1000),2)', 6710.08);
check('nper', 'ROUND(NPER(0.05,-1000,10000),4)', 14.2067);
check('rate', 'ROUND(RATE(48,-200,8000),6)', 0.007701);
check('ipmt', 'ROUND(IPMT(0.1/12,1,24,2000),4)', -16.6667);
check('ppmt plus ipmt equals pmt', 'ROUND(PPMT(0.1/12,1,24,2000)+IPMT(0.1/12,1,24,2000),4)', -92.2899);
check('npv', 'ROUND(NPV(0.1,-10000,3000,4200,6800),2)', 1188.44);
check('irr', 'ROUND(IRR({-10000;3000;4200;6800}),6)', 0.163406);
check('irr excel example', 'ROUND(IRR({-70000;12000;15000;18000;21000;26000}),6)', 0.086631);

/* -------------------------------------------------------------------- date */

check('date serial', 'DATE(2024,1,1)', 45292);
check('year', 'YEAR(DATE(2024,3,5))', 2024);
check('month', 'MONTH(DATE(2024,3,5))', 3);
check('day', 'DAY(DATE(2024,3,5))', 5);
check('days', 'DAYS(DATE(2024,1,10),DATE(2024,1,1))', 9);
check('eomonth', 'DAY(EOMONTH(DATE(2024,1,15),1))', 29);
check('edate leap', 'DAY(EDATE(DATE(2024,1,31),1))', 29);
check('datevalue', 'DATEVALUE("2024-01-01")', 45292);
check('weekday', 'WEEKDAY(DATE(2024,1,1))', 2);
check('workday', 'WORKDAY(DATE(2024,1,1),5)', 45299);
check('networkdays', 'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))', 5);
check('datedif y', 'DATEDIF(DATE(2020,1,1),DATE(2024,3,1),"Y")', 4);
check('yearfrac', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),0)', 0.5);

/* -------------------------------------------------------- dynamic arrays */

check('sequence', 'SUM(SEQUENCE(4))', 10);
check('sequence 2d', 'INDEX(SEQUENCE(2,3,1,1),2,3)', 6);
check('filter', 'SUM(FILTER({1;2;3;4},{TRUE;FALSE;TRUE;FALSE}))', 4);
check('unique', 'COUNTA(UNIQUE({1;2;2;3}))', 3);
check('sortby', 'INDEX(SORTBY({10;20;30},{3;1;2}),1)', 20);

/* -------------------------------------------------------- dynamic spill */

{
  // FILTER writes into A1:A3; the second and third results must be readable.
  const engine = new Engine(newWorkbook());
  const sheet = engine.active;
  engine.setInput(sheet, 0, 0, '=FILTER({1;2;3},{TRUE;FALSE;TRUE})');
  engine.recalc();
  eq('spill anchor', engine.valueAt(sheet, 0, 0), 1);
  eq('spill row 2', engine.valueAt(sheet, 1, 0), 3);
  eq('spill row 3 is empty', engine.valueAt(sheet, 2, 0), null);

  // A blocking cell must produce #SPILL! instead of overwriting data.
  const e2 = new Engine(newWorkbook());
  const s2 = e2.active;
  e2.setInput(s2, 1, 0, 'blocker');
  e2.setInput(s2, 0, 0, '=SEQUENCE(3)');
  e2.recalc();
  const anchor = e2.valueAt(s2, 0, 0);
  const blocked = typeof anchor === 'object' && anchor !== null && '__error' in anchor;
  eq('spill blocked reports error', blocked, true);
}

/* ------------------------------------------------------------ ref shifting */

eq('translate down', translateFormula('A1', 1, 0), 'A2');
eq('translate absolute', translateFormula('$A$1', 1, 0), '$A$1');
eq('translate mixed', translateFormula('A$1', 1, 0), 'A$1');
eq('translate row abs', translateFormula('$A1', 1, 0), '$A2');
eq('translate with sum', translateFormula('SUM(A1:B2)+A1', 1, 1), 'SUM(B2:C3)+B2');
eq('translate skips strings', translateFormula('"A1"&A1', 1, 0), '"A1"&A2');
eq('translate sheet', translateFormula('Sheet1!A1', 2, 0), 'Sheet1!A3');
eq('translate off grid', translateFormula('A1', -5, 0), '#REF!');

/* -------------------------------------------------------- copy and fill */

{
  // Copying `=A1*2` down two rows must re-point at A2 and A3.
  const engine = new Engine(newWorkbook());
  const sheet = engine.active;
  engine.setInput(sheet, 0, 0, '10');
  engine.setInput(sheet, 1, 0, '20');
  engine.setInput(sheet, 2, 0, '30');
  engine.setInput(sheet, 0, 1, '=A1*2');
  engine.recalc();
  eq('copy source value', engine.valueAt(sheet, 0, 1), 20);

  engine.copyRange(sheet, { r1: 1, c1: 1, r2: 2, c2: 1 }, { r1: 0, c1: 1, r2: 0, c2: 1 });
  engine.recalc();
  eq('copy row 2 translated', engine.valueAt(sheet, 1, 1), 40);
  eq('copy row 3 translated', engine.valueAt(sheet, 2, 1), 60);
  eq('copy kept formula text', engine.editText(sheet, 2, 1), '=A3*2');

  // An absolute reference must not move.
  const e2 = new Engine(newWorkbook());
  const s2 = e2.active;
  e2.setInput(s2, 0, 0, '7');
  e2.setInput(s2, 0, 1, '=$A$1+1');
  e2.recalc();
  e2.copyRange(s2, { r1: 1, c1: 1, r2: 1, c2: 1 }, { r1: 0, c1: 1, r2: 0, c2: 1 });
  e2.recalc();
  eq('absolute ref survives copy', e2.valueAt(s2, 1, 1), 8);
}

/* ------------------------------------------------------------- formatting */

check('number format', 'TEXT(1234.5,"#,##0.00")', '1,234.50');
check('date format', 'TEXT(DATE(2024,1,15),"yyyy-mm-dd")', '2024-01-15');
check('month name', 'TEXT(DATE(2024,1,15),"mmm")', 'Jan');
check('negative format', 'TEXT(-5,"#,##0;(#,##0)")', '(5)');

/* ------------------------------------------------------------------ report */

console.log(`\n  ${pass} passed, ${fail} failed\n`);
if (failures.length) {
  console.log('Failures:');
  for (const f of failures) console.log(`  - ${f}`);
  console.log('');
}
process.exit(fail === 0 ? 0 : 1);
