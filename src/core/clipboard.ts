/**
 * Clipboard handling.
 *
 * A rich payload is kept in memory for same-window paste, and a text rendering
 * is written to the system clipboard so Excel and Sheets can exchange tabs and
 * newlines. Reading prefers the text form, which round-trips through the CSV
 * parser.
 */

import { rangeToA1 } from './ref';

export interface ClipCell {
  r: number;
  c: number;
  /** formula (with `=`) or literal text */
  text: string;
  styleId?: number;
}

export interface ClipPayload {
  sheetName: string;
  range: { r1: number; c1: number; r2: number; c2: number };
  cells: ClipCell[];
}

/** Custom MIME type so a paste inside SheetCraft keeps its formulas. */
const MIME = 'application/x-sheetcraft-cells';

let inMemory: ClipPayload | null = null;

/** Rows as TSV, which spreadsheets parse cleanly. */
export function toTsv(payload: ClipPayload): string {
  const { range, cells } = payload;
  const byRow = new Map<number, Map<number, ClipCell>>();
  for (const cell of cells) {
    let row = byRow.get(cell.r);
    if (!row) {
      row = new Map();
      byRow.set(cell.r, row);
    }
    row.set(cell.c, cell);
  }
  const out: string[] = [];
  for (let r = range.r1; r <= range.r2; r++) {
    const cols: string[] = [];
    for (let c = range.c1; c <= range.c2; c++) {
      cols.push((byRow.get(r)?.get(c)?.text ?? '').replace(/[\t\r\n]/g, ' '));
    }
    out.push(cols.join('\t'));
  }
  return out.join('\n');
}

/** Parse TSV (or CSV) text back into a payload. */
export function fromTsv(text: string, sheetName = 'Sheet1'): ClipPayload | null {
  const rows = text
    .replace(/^﻿/, '')
    .split(/\r?\n/)
    .filter((line, i, all) => line !== '' || i < all.length - 1);
  if (!rows.length) return null;

  const grid = rows.map((line) => splitLine(line));
  const width = grid.reduce((w, r) => Math.max(w, r.length), 0);

  const cells: ClipCell[] = [];
  grid.forEach((row, r) => {
    row.forEach((value, c) => {
      cells.push({ r, c, text: value });
    });
  });

  return {
    sheetName,
    range: { r1: 0, c1: 0, r2: Math.max(0, grid.length - 1), c2: Math.max(0, width - 1) },
    cells,
  };
}

/** Split a pasted line, honouring tabs inside quoted CSV fields. */
function splitLine(line: string): string[] {
  if (!line.includes('"')) return line.split('\t');
  const out: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else inQuotes = false;
      } else cur += ch;
      continue;
    }
    if (ch === '"') inQuotes = true;
    else if (ch === '\t') {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** Write both the rich payload and a TSV fallback. */
export async function copyToClipboard(payload: ClipPayload): Promise<void> {
  inMemory = payload;
  const text = toTsv(payload);
  try {
    if (typeof navigator !== 'undefined' && navigator.clipboard?.write) {
      const items: Record<string, string> = { 'text/plain': text };
      if (typeof ClipboardItem !== 'undefined' && typeof Blob !== 'undefined') {
        items[MIME] = JSON.stringify(payload);
      }
      const blobs: Record<string, Blob> = {
        'text/plain': new Blob([text], { type: 'text/plain' }),
      };
      if (items[MIME]) {
        blobs[MIME] = new Blob([items[MIME]], { type: MIME });
      }
      await navigator.clipboard.write([new ClipboardItem(blobs)]);
    }
  } catch {
    // Clipboard permission denied (common in file:// contexts); the in-memory
    // payload still allows paste inside the app.
  }
}

/** Current in-app payload, without touching the system clipboard. */
export function localPayload(): ClipPayload | null {
  return inMemory;
}

/** Read the system clipboard, preferring our own MIME type when present. */
export async function readClipboardText(sheetName = 'Sheet1'): Promise<ClipPayload | null> {
  if (inMemory) return inMemory;
  try {
    if (typeof navigator === 'undefined' || !navigator.clipboard?.read) return null;
    const items = await navigator.clipboard.read();
    for (const item of items) {
      try {
        const rich = await item.getType(MIME);
        const parsed = JSON.parse(await rich.text()) as ClipPayload;
        inMemory = parsed;
        return parsed;
      } catch {
        // This item does not carry our MIME type; keep looking.
      }
    }
    for (const item of items) {
      try {
        const plain = await item.getType('text/plain');
        const payload = fromTsv(await plain.text(), sheetName);
        if (payload) inMemory = payload;
        return payload;
      } catch {
        // ignore and try the next item
      }
    }
  } catch {
    // readText may be unavailable; fall through to the null result.
  }
  return null;
}

/** Human-readable description of a payload, for the status bar. */
export function describePayload(payload: ClipPayload | null): string {
  if (!payload) return 'Clipboard empty';
  const label = rangeToA1(payload.range);
  return `Copied ${label} (${payload.cells.length} cell${payload.cells.length === 1 ? '' : 's'})`;
}
