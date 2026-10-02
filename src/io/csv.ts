/**
 * CSV import/export.
 *
 * The parser follows RFC 4180 (quoted fields, embedded newlines and doubled
 * quotes) and infers nothing: values land as typed cells only when they look
 * unambiguous, so `007` stays text the way Excel treats it.
 */

import type { CellFormat, Range } from '../core/types';
import { parseInput } from '../core/coerce';
import { rangeToA1 } from '../core/ref';

export interface CsvGrid {
  rows: string[][];
}

/** Split CSV text into rows of raw strings. */
export function parseCsv(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const n = text.length;

  // Strip a UTF-8 BOM.
  if (text.charCodeAt(0) === 0xfeff) i = 1;

  while (i < n) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\r') {
      i++;
      continue;
    }
    if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i++;
      continue;
    }
    field += ch;
    i++;
  }

  // Flush whatever is left, unless the file ended with a clean newline.
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** Serialise a grid back to CSV text. */
export function toCsv(rows: (string | number | boolean | null)[][], delimiter = ','): string {
  const body = rows
    .map((r) => r.map((v) => escapeField(v === null ? '' : String(v), delimiter)).join(delimiter))
    .join('\r\n');
  return `\uFEFF${body}`;
}

function escapeField(value: string, delimiter: string): string {
  const needsQuotes =
    value.includes(delimiter) ||
    value.includes('"') ||
    value.includes('\n') ||
    value.includes('\r');
  return needsQuotes ? `"${value.replace(/"/g, '""')}"` : value;
}

/** Guess the delimiter from the header line (comma, semicolon or tab). */
export function sniffDelimiter(text: string): string {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? '';
  const candidates = [',', ';', '\t', '|'];
  let best = ',';
  let bestCount = -1;
  for (const d of candidates) {
    let count = 0;
    let inQuotes = false;
    for (let i = 0; i < firstLine.length; i++) {
      const ch = firstLine[i];
      if (ch === '"') inQuotes = !inQuotes;
      else if (ch === d && !inQuotes) count++;
    }
    if (count > bestCount) {
      bestCount = count;
      best = d;
    }
  }
  return best;
}

/** Cell text as it should appear in CSV. */
export function cellToText(v: string | number | boolean | null): string {
  return v === null ? '' : String(v);
}

export interface CsvOptions {
  delimiter?: string;
  /** row where data starts (0-based); 1 means "the first row is a header" */
  startRow?: number;
  /** treat each field as a date when it parses as one */
  parseTypes?: boolean;
  /** format applied to the header row */
  headerFormat?: CellFormat;
  /** infer a date format for date-looking columns */
  detectDates?: boolean;
}

export interface CsvImportResult {
  /** rows of parsed values ready to be written into a sheet */
  rows: (string | number | boolean | null)[][];
  /** date-looking columns, so the caller can apply a number format */
  dateColumns: number[];
  rowCount: number;
  colCount: number;
}

/**
 * Parse CSV into typed values, reporting which columns look like dates.
 */
export function importCsv(text: string, options: CsvOptions = {}): CsvImportResult {
  const delimiter = options.delimiter ?? sniffDelimiter(text);
  const grid = parseCsv(text, delimiter);
  const startRow = options.startRow ?? 0;
  const parseTypes = options.parseTypes ?? true;
  const detectDates = options.detectDates ?? true;

  const body = grid.slice(startRow);
  const colCount = body.reduce((w, r) => Math.max(w, r.length), 0);

  const dateColumns: number[] = [];
  if (detectDates) {
    for (let c = 0; c < colCount; c++) {
      let seen = 0;
      let dates = 0;
      for (const row of body) {
        const raw = (row[c] ?? '').trim();
        if (raw === '') continue;
        seen++;
        if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(raw) || /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/.test(raw)) {
          dates++;
        }
      }
      if (seen > 0 && dates === seen) dateColumns.push(c);
    }
  }

  const rows = body.map((row) => {
    const out: (string | number | boolean | null)[] = [];
    for (let c = 0; c < colCount; c++) {
      const raw = row[c] ?? '';
      if (!parseTypes) {
        out.push(raw === '' ? null : raw);
        continue;
      }
      const parsed = parseInput(raw);
      out.push(parsed);
    }
    return out;
  });

  return { rows, dateColumns, rowCount: rows.length, colCount };
}

/** Range covering the imported data, for status reporting. */
export function importRange(result: CsvImportResult, startRow = 0, startCol = 0): Range {
  return {
    r1: startRow,
    c1: startCol,
    r2: startRow + Math.max(0, result.rowCount - 1),
    c2: startCol + Math.max(0, result.colCount - 1),
  };
}

/** A1 label describing where an import landed, e.g. `A1:D250`. */
export function importLabel(result: CsvImportResult): string {
  return rangeToA1(importRange(result));
}
