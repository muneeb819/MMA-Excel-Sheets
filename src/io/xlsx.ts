/**
 * XLSX import/export via SheetJS.
 *
 * Formulas survive a round trip because the workbook keeps its formula text,
 * and styles map onto the internal format model closely enough that fonts,
 * fills, alignment and number formats come back looking the same.
 */

import * as XLSX from 'xlsx';
import type { CellFormat, Range, Sheet, Value, Workbook } from '../core/types';
import { formatValue, isDateFormat, parseInput, toText } from '../core/coerce';
import { colName, iterRange, normRange, unpackCol, unpackRow } from '../core/ref';
import type { Engine } from '../core/engine';
import { uid } from '../core/engine';

export interface ImportResult {
  sheetNames: string[];
  cellCount: number;
}

/**
 * Replace the engine's workbook contents with the given file.
 * Formulas are preserved as text; everything else is coerced to a literal.
 */
export function importXlsx(engine: Engine, data: ArrayBuffer): ImportResult {
  const wb = XLSX.read(data, { type: 'array', cellStyles: true, cellDates: false, cellNF: true });
  const target = engine.wb;

  // Reuse the existing sheet objects so the UI keeps its references valid.
  while (target.sheets.length > wb.SheetNames.length) target.sheets.pop();
  target.namedRanges = {};

  let cellCount = 0;

  wb.SheetNames.forEach((name, index) => {
    const ws = wb.Sheets[name];
    let sheet = target.sheets[index];
    if (!sheet) {
      sheet = {
        id: uid('sh'),
        name,
        cells: new Map(),
        rows: 100,
        cols: 26,
        colWidths: new Map(),
        rowHeights: new Map(),
        hiddenRows: new Set(),
        hiddenCols: new Set(),
        frozenRows: 0,
        frozenCols: 0,
        charts: [],
        condFormats: [],
        defColWidth: 96,
        defRowHeight: 22,
      };
      target.sheets.push(sheet);
    }
    sheet.name = name;
    sheet.cells = new Map();
    sheet.colWidths = new Map();
    sheet.rowHeights = new Map();
    sheet.charts = [];
    sheet.condFormats = [];
    sheet.hiddenRows = new Set();
    sheet.hiddenCols = new Set();

    const ref = ws['!ref'] ?? 'A1';
    const range = XLSX.utils.decode_range(ref);

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const address = XLSX.utils.encode_cell({ r, c });
        const cell = ws[address] as XLSX.CellObject | undefined;
        if (!cell) continue;
        cellCount++;

        const fmt = readFormat(cell);
        const styleId = fmt ? engine.styleIdFor(fmt) : undefined;

        if (cell.f) {
          const text = `=${cell.f}`;
          engine.setInput(sheet, r, c, text);
          if (styleId !== undefined) engine.setStyleId(sheet, r, c, styleId);
        } else if (cell.t === 'n') {
          engine.setValue(sheet, r, c, Number(cell.v), styleId);
        } else if (cell.t === 'b') {
          engine.setValue(sheet, r, c, Boolean(cell.v), styleId);
        } else if (cell.t === 'e') {
          engine.setValue(sheet, r, c, null, styleId);
        } else {
          const text = cell.v === undefined || cell.v === null ? null : String(cell.v);
          engine.setValue(sheet, r, c, text ?? null, styleId);
        }
      }
    }

    // Column widths (SheetJS reports them in character units).
    const cols = ws['!cols'] as Array<{ wpx?: number; wch?: number }> | undefined;
    if (cols) {
      cols.forEach((col, c) => {
        if (col?.wpx) sheet.colWidths.set(c, Math.max(24, Math.round(col.wpx)));
        else if (col?.wch) sheet.colWidths.set(c, Math.round(col.wch * 7));
      });
    }

    const rows = ws['!rows'] as Array<{ hpx?: number; hpt?: number }> | undefined;
    if (rows) {
      rows.forEach((row, r) => {
        if (row?.hpx) sheet.rowHeights.set(r, Math.max(14, Math.round(row.hpx)));
        else if (row?.hpt) sheet.rowHeights.set(r, Math.round(row.hpt * 1.33));
      });
    }

    const merges = ws['!merges'] as Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> | undefined;
    if (merges?.length) {
      for (const m of merges) {
        const topLeft = sheet.cells.get(m.s.r * 16384 + m.s.c);
        for (let r = m.s.r; r <= m.e.r; r++) {
          for (let c = m.s.c; c <= m.e.c; c++) {
            if (r === m.s.r && c === m.s.c) continue;
            const cell = { ...(topLeft ?? {}) };
            delete cell.formula;
            delete cell.raw;
            sheet.cells.set(r * 16384 + c, cell);
          }
        }
      }
    }
  });

  target.activeSheet = 0;
  target.name = 'Book1';
  engine.trimAll();
  engine.invalidateAll();
  engine.recalc();
  engine.emit();
  return { sheetNames: wb.SheetNames.slice(), cellCount };
}

function readFormat(cell: XLSX.CellObject): CellFormat | undefined {
  const fmt: CellFormat = {};
  const s = cell.s as
    | {
        font?: { bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; color?: { rgb?: string }; name?: string; sz?: number };
        fill?: { fgColor?: { rgb?: string }; patternType?: string };
        alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
        numFmt?: string;
      }
    | undefined;

  if (s?.font) {
    if (s.font.bold) fmt.bold = 1;
    if (s.font.italic) fmt.italic = 1;
    if (s.font.underline) fmt.underline = 1;
    if (s.font.strike) fmt.strike = 1;
    if (s.font.name) fmt.font = s.font.name;
    if (s.font.sz) fmt.size = s.font.sz;
    if (s.font.color?.rgb) fmt.color = rgbToCss(s.font.color.rgb);
  }
  if (s?.fill?.patternType === 'solid' && s.fill.fgColor?.rgb) {
    fmt.bg = rgbToCss(s.fill.fgColor.rgb);
  }
  if (s?.alignment?.horizontal) {
    const h = s.alignment.horizontal;
    if (h === 'left' || h === 'center' || h === 'right') fmt.align = h;
    else if (h === 'justify') fmt.align = 'left';
  }
  if (s?.alignment?.wrapText) fmt.wrap = 1;
  const nf = cell.z ?? s?.numFmt;
  if (typeof nf === 'string' && nf !== '' && nf !== 'General') fmt.numFmt = nf;

  return Object.keys(fmt).length ? fmt : undefined;
}

function rgbToCss(rgb: string): string {
  if (rgb.length === 8) return `#${rgb.slice(2)}`;
  if (rgb.length === 6) return `#${rgb}`;
  return rgb;
}

/** Build an XLSX buffer from the current workbook. */
export function exportXlsx(engine: Engine): ArrayBuffer {
  const wb = XLSX.utils.book_new();

  for (const sheet of engine.wb.sheets) {
    const ws: XLSX.WorkSheet = {};
    let maxR = 0;
    let maxC = 0;
    let count = 0;

    for (const [k, cell] of sheet.cells) {
      const r = unpackRow(k);
      const c = unpackCol(k);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
      const address = XLSX.utils.encode_cell({ r, c });

      const out: XLSX.CellObject = { t: 's', v: '' };
      if (cell.formula !== undefined) {
        out.f = cell.formula;
        const value = engine.valueAt(sheet, r, c);
        if (typeof value === 'number') {
          out.t = 'n';
          out.v = value;
        } else if (typeof value === 'boolean') {
          out.t = 'b';
          out.v = value;
        } else if (typeof value === 'string') {
          out.t = 's';
          out.v = value;
        }
      } else if (cell.raw === undefined) {
        // style-only cell; keep it so formatting survives
        if (cell.styleId === undefined) continue;
        out.t = 'z';
      } else if (typeof cell.raw === 'number') {
        out.t = 'n';
        out.v = cell.raw;
      } else if (typeof cell.raw === 'boolean') {
        out.t = 'b';
        out.v = cell.raw;
      } else {
        out.t = 's';
        out.v = cell.raw ?? '';
      }

      const fmt = engine.formatOf(cell.styleId);
      if (fmt?.numFmt) out.z = fmt.numFmt;
      out.s = toSheetStyle(fmt) as XLSX.CellObject['s'];

      ws[address] = out;
      count++;
    }

    if (count === 0) ws['A1'] = { t: 's', v: '' };
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, maxR), c: Math.max(0, maxC) } });

    const cols: Array<{ wpx?: number }> = [];
    for (const [c, w] of sheet.colWidths) cols[c] = { wpx: w };
    if (cols.length) ws['!cols'] = cols;

    const rows: Array<{ hpx?: number }> = [];
    for (const [r, h] of sheet.rowHeights) rows[r] = { hpx: h };
    if (rows.length) ws['!rows'] = rows;

    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31));
  }

  if (wb.SheetNames.length === 0) {
    const ws = XLSX.utils.aoa_to_sheet([['']]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  }

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true }) as ArrayBuffer;
}

function toSheetStyle(fmt: CellFormat | undefined): Record<string, unknown> | undefined {
  if (!fmt) return undefined;
  const style: Record<string, unknown> = {};
  const font: Record<string, unknown> = {};
  if (fmt.bold) font.bold = true;
  if (fmt.italic) font.italic = true;
  if (fmt.underline) font.underline = true;
  if (fmt.strike) font.strike = true;
  if (fmt.color) font.color = { rgb: cssToRgb(fmt.color) };
  if (fmt.font) font.name = fmt.font;
  if (fmt.size) font.sz = fmt.size;
  if (Object.keys(font).length) style.font = font;

  if (fmt.bg) style.fill = { patternType: 'solid', fgColor: { rgb: cssToRgb(fmt.bg) } };

  const alignment: Record<string, unknown> = {};
  if (fmt.align) alignment.horizontal = fmt.align;
  if (fmt.valign) alignment.vertical = fmt.valign;
  if (fmt.wrap) alignment.wrapText = true;
  if (Object.keys(alignment).length) style.alignment = alignment;

  return Object.keys(style).length ? style : undefined;
}

function cssToRgb(css: string): string {
  const hex = css.replace('#', '');
  if (hex.length === 6) return `FF${hex.toUpperCase()}`;
  if (hex.length === 8) return hex.toUpperCase();
  return `FF${hex.toUpperCase().padEnd(6, '0')}`;
}

/** Values of a sheet as plain rows, ready to be joined into CSV text. */
export function sheetToRows(engine: Engine, sheet: Sheet, range?: Range): string[][] {
  const used = range ?? engine.usedRange(sheet);
  const n = normRange(used);
  const rows: string[][] = [];
  for (let r = n.r1; r <= n.r2; r++) {
    const row: string[] = [];
    for (let c = n.c1; c <= n.c2; c++) row.push(toText(engine.valueAt(sheet, r, c)));
    rows.push(row);
  }
  return rows;
}

/** A1 label for a range, used in status messages. */
export function describe(range: Range): string {
  const n = normRange(range);
  return n.r1 === n.r2 && n.c1 === n.c2
    ? `${colName(n.c1)}${n.r1 + 1}`
    : `${colName(n.c1)}${n.r1 + 1}:${colName(n.c2)}${n.r2 + 1}`;
}

export { formatValue, isDateFormat, parseInput };
export type { Value, Workbook };
