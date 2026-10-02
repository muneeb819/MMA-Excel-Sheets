/**
 * The spreadsheet grid.
 *
 * Rendering is canvas based and virtualised: only the visible rows and columns
 * are drawn, so a sheet with a million rows scrolls at the same speed as a small
 * one. Cell editing happens in an overlaid textarea positioned exactly over the
 * active cell.
 */

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  DEFAULT_COL_W,
  DEFAULT_ROW_H,
  HEADER_H,
  HEADER_W,
  MAX_COLS,
  MAX_ROWS,
  type CellFormat,
  type Range,
  type Sheet,
  type Value,
} from '../core/types';
import { a1, colName, normRange, pack, rangeContains, rangeToA1, translateFormula, unpackCol, unpackRow } from '../core/ref';
import { formatValue } from '../core/coerce';
import { isError } from '../core/types';
import type { Engine } from '../core/engine';
import { copyToClipboard, readClipboardText, type ClipPayload } from '../core/clipboard';

export interface Selection {
  /** the cell the cursor is on */
  cursor: { r: number; c: number };
  /** the selected block */
  range: Range;
}

interface GridProps {
  engine: Engine;
  sheet: Sheet;
  selection: Selection;
  onSelectionChange: (next: Selection) => void;
  /** bumped by the parent to force a repaint */
  revision: number;
  zoom: number;
  /** raised when the user asks to edit a cell from the formula bar */
  editSignal: { r: number; c: number; text: string; mode: 'replace' | 'append' } | null;
  onEditSignalHandled: () => void;
}

interface Metrics {
  colWidths: number[];
  rowHeights: number[];
  /** cumulative pixel offset of each visible row/col */
  colX: Float64Array;
  rowY: Float64Array;
  totalW: number;
  totalH: number;
}

/** Cumulative offsets so scroll -> cell lookups are O(log n). */
function buildMetrics(engine: Engine, sheet: Sheet, rows: number, cols: number, zoom: number): Metrics {
  const colWidths = new Array<number>(cols);
  for (let c = 0; c < cols; c++) colWidths[c] = engine.colWidth(sheet, c) * zoom;

  const rowHeights = new Array<number>(rows);
  for (let r = 0; r < rows; r++) rowHeights[r] = engine.rowHeight(sheet, r) * zoom;

  const colX = new Float64Array(cols + 1);
  for (let c = 0; c < cols; c++) colX[c + 1] = colX[c] + colWidths[c];
  const rowY = new Float64Array(rows + 1);
  for (let r = 0; r < rows; r++) rowY[r + 1] = rowY[r] + rowHeights[r];

  return {
    colWidths,
    rowHeights,
    colX,
    rowY,
    totalW: colX[cols] ?? 0,
    totalH: rowY[rows] ?? 0,
  };
}

/** Largest index whose cumulative offset is <= value. */
function indexAt(offsets: Float64Array, value: number, max: number): number {
  let lo = 0;
  let hi = max;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (offsets[mid] <= value) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

export default function Grid({
  engine,
  sheet,
  selection,
  onSelectionChange,
  revision,
  zoom,
  editSignal,
  onEditSignalHandled,
}: GridProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const hostRef = useRef<HTMLDivElement | null>(null);
  const scrollRef = useRef({ x: 0, y: 0 });
  const sizeRef = useRef({ w: 800, h: 600 });
  const [size, setSize] = useState({ w: 800, h: 600 });

  // Editing state
  const [editing, setEditing] = useState<{ r: number; c: number; text: string } | null>(null);
  const editorRef = useRef<HTMLTextAreaElement | null>(null);

  // Drag state kept in a ref so mouse moves do not re-render.
  const dragRef = useRef<{
    mode: 'cells' | 'rows' | 'cols' | 'fill' | 'move';
    anchor: { r: number; c: number };
    origin: { x: number; y: number };
    moved: boolean;
  } | null>(null);

  const [fillPreview, setFillPreview] = useState<Range | null>(null);
  const [dropPreview, setDropPreview] = useState<{ r: number; c: number } | null>(null);
  const [clip, setClip] = useState<ClipPayload | null>(null);

  const viewRows = Math.min(Math.max(sheet.rows, 200), 20000);
  const viewCols = Math.min(Math.max(sheet.cols, 60), 400);

  const metrics = useMemo(
    () => buildMetrics(engine, sheet, viewRows, viewCols, zoom),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet, viewRows, viewCols, zoom, revision],
  );

  const frozenW = 0;
  const frozenH = 0;
  const bodyW = size.w - HEADER_W;
  const bodyH = size.h - HEADER_H;

  /* ------------------------------------------------------------- rendering */

  const draw = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const dpr = window.devicePixelRatio || 1;
    const { w, h } = sizeRef.current;
    if (canvas.width !== Math.round(w * dpr) || canvas.height !== Math.round(h * dpr)) {
      canvas.width = Math.round(w * dpr);
      canvas.height = Math.round(h * dpr);
      canvas.style.width = `${w}px`;
      canvas.style.height = `${h}px`;
    }
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);

    const scroll = scrollRef.current;
    const m = metrics;
    const sel = normRange(selection.range);

    const firstRow = indexAt(m.rowY, scroll.y, viewRows);
    const lastRow = Math.min(viewRows - 1, indexAt(m.rowY, scroll.y + bodyH, viewRows));
    const firstCol = indexAt(m.colX, scroll.x, viewCols);
    const lastCol = Math.min(viewCols - 1, indexAt(m.colX, scroll.x + bodyW, viewCols));

    const xOf = (c: number): number => HEADER_W + m.colX[c] - scroll.x;
    const yOf = (r: number): number => HEADER_H + m.rowY[r] - scroll.y;

    /* selection wash */
    if (rangeContains(sel, firstRow, firstCol) || true) {
      const x1 = xOf(sel.c1);
      const y1 = yOf(sel.r1);
      const x2 = xOf(sel.c2 + 1);
      const y2 = yOf(sel.r2 + 1);
      ctx.fillStyle = 'rgba(47, 111, 237, 0.10)';
      ctx.fillRect(x1, y1, x2 - x1, y2 - y1);
    }

    /* fill / drop preview */
    if (fillPreview) {
      const x1 = xOf(fillPreview.c1);
      const y1 = yOf(fillPreview.r1);
      const x2 = xOf(fillPreview.c2 + 1);
      const y2 = yOf(fillPreview.r2 + 1);
      ctx.strokeStyle = '#2f6fed';
      ctx.setLineDash([4, 3]);
      ctx.lineWidth = 1.5;
      ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
      ctx.setLineDash([]);
      if (dropPreview) {
        ctx.strokeStyle = '#22a06b';
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(xOf(dropPreview.c), yOf(dropPreview.r));
        ctx.stroke();
      }
    }

    /* cell text */
    ctx.textBaseline = 'middle';
    for (let r = firstRow; r <= lastRow; r++) {
      const y = yOf(r);
      const rowH = m.rowHeights[r];
      if (y + rowH < HEADER_H || y > h) continue;

      for (let c = firstCol; c <= lastCol; c++) {
        const x = xOf(c);
        const colW = m.colWidths[c];
        if (x + colW < HEADER_W || x > w) continue;

        const cell = sheet.cells.get(pack(r, c));
        if (!cell) continue;

        const fmt = engine.formatOf(cell.styleId);
        const value = engine.valueAt(sheet, r, c);

        // background
        if (fmt?.bg) {
          ctx.fillStyle = fmt.bg;
          ctx.fillRect(x + 1, y + 1, colW - 1, rowH - 1);
        }

        const text = formatValue(value, fmt?.numFmt);
        if (text === '') continue;

        ctx.font = fontFor(fmt);
        ctx.fillStyle = isError(value) ? '#c2436b' : fmt?.color ?? '#1f2733';

        const pad = 4;
        const maxW = colW - pad * 2;
        let shown = text;
        if (ctx.measureText(shown).width > maxW) {
          // Trim from the right, like a spreadsheet cell that overflows.
          while (shown.length > 1 && ctx.measureText(`${shown}...`).width > maxW) {
            shown = shown.slice(0, -1);
          }
          shown = `${shown}...`;
        }

        let tx = x + pad;
        if (fmt?.align === 'center') {
          ctx.textAlign = 'center';
          tx = x + colW / 2;
        } else if (fmt?.align === 'right' || (!fmt?.align && typeof value === 'number')) {
          ctx.textAlign = 'right';
          tx = x + colW - pad;
        } else {
          ctx.textAlign = 'left';
        }

        ctx.fillText(shown, tx, y + rowH / 2);
        ctx.textAlign = 'left';
      }
    }

    /* grid lines */
    ctx.strokeStyle = '#dfe3e8';
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let r = firstRow; r <= lastRow + 1; r++) {
      const y = Math.round(yOf(r)) + 0.5;
      ctx.moveTo(HEADER_W, y);
      ctx.lineTo(w, y);
    }
    for (let c = firstCol; c <= lastCol + 1; c++) {
      const x = Math.round(xOf(c)) + 0.5;
      ctx.moveTo(x, HEADER_H);
      ctx.lineTo(x, h);
    }
    ctx.stroke();

    /* highlighted row/column headers */
    const inSel = (r: number, c: number): boolean =>
      rangeContains(sel, r, c) || (r === selection.cursor.r) || (c === selection.cursor.c);

    ctx.fillStyle = '#eef3ff';
    for (let r = firstRow; r <= lastRow; r++) {
      if (r !== selection.cursor.r && !rangeContains(sel, r, sel.c1)) continue;
      ctx.fillRect(0, yOf(r), HEADER_W, m.rowHeights[r]);
    }
    for (let c = firstCol; c <= lastCol; c++) {
      if (c !== selection.cursor.c && !rangeContains(sel, sel.r1, c)) continue;
      ctx.fillRect(xOf(c), 0, m.colWidths[c], HEADER_H);
    }

    /* selection border */
    ctx.strokeStyle = '#2f6fed';
    ctx.lineWidth = 2;
    const sx1 = xOf(sel.c1);
    const sy1 = yOf(sel.r1);
    const sx2 = xOf(sel.c2 + 1);
    const sy2 = yOf(sel.r2 + 1);
    ctx.strokeRect(sx1 + 1, sy1 + 1, sx2 - sx1 - 1, sy2 - sy1 - 1);

    /* fill handle */
    if (selection.range.r1 === selection.range.r2 || selection.range.c1 === selection.range.c2) {
      ctx.fillStyle = '#2f6fed';
      ctx.fillRect(sx2 - 4, sy2 - 4, 6, 6);
    }

    /* headers */
    ctx.fillStyle = '#f7f8fa';
    ctx.fillRect(0, 0, w, HEADER_H);
    ctx.fillRect(0, 0, HEADER_W, h);
    ctx.strokeStyle = '#d7dbe0';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, HEADER_H + 0.5);
    ctx.lineTo(w, HEADER_H + 0.5);
    ctx.moveTo(HEADER_W + 0.5, 0);
    ctx.lineTo(HEADER_W + 0.5, h);
    ctx.stroke();

    ctx.fillStyle = '#5b6572';
    ctx.font = '11px ' + "'Segoe UI', system-ui, sans-serif";
    ctx.textAlign = 'center';
    for (let c = firstCol; c <= lastCol; c++) {
      const x = xOf(c);
      if (x + m.colWidths[c] < HEADER_W || x > w) continue;
      ctx.fillStyle = inSel(selection.cursor.r, c) ? '#1b4bb8' : '#5b6572';
      ctx.fillText(colName(c), x + m.colWidths[c] / 2, HEADER_H / 2);
      ctx.strokeStyle = '#d7dbe0';
      ctx.beginPath();
      ctx.moveTo(Math.round(x + m.colWidths[c]) + 0.5, 4);
      ctx.lineTo(Math.round(x + m.colWidths[c]) + 0.5, HEADER_H - 4);
      ctx.stroke();
    }
    for (let r = firstRow; r <= lastRow; r++) {
      const y = yOf(r);
      if (y + m.rowHeights[r] < HEADER_H || y > h) continue;
      ctx.fillStyle = inSel(r, selection.cursor.c) ? '#1b4bb8' : '#5b6572';
      ctx.fillText(String(r + 1), HEADER_W / 2, y + m.rowHeights[r] / 2);
      ctx.strokeStyle = '#d7dbe0';
      ctx.beginPath();
      ctx.moveTo(4, Math.round(y + m.rowHeights[r]) + 0.5);
      ctx.lineTo(HEADER_W - 4, Math.round(y + m.rowHeights[r]) + 0.5);
      ctx.stroke();
    }

    /* frozen split line */
    if (frozenW || frozenH) {
      ctx.strokeStyle = '#b9c0c8';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(HEADER_W + frozenW, 0);
      ctx.lineTo(HEADER_W + frozenW, h);
      ctx.moveTo(0, HEADER_H + frozenH);
      ctx.lineTo(w, HEADER_H + frozenH);
      ctx.stroke();
    }

    ctx.textAlign = 'left';
  }, [engine, sheet, selection, metrics, viewRows, viewCols, fillPreview, dropPreview, frozenW, frozenH, bodyH, bodyW]);

  useLayoutEffect(() => {
    draw();
  }, [draw, revision]);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const ro = new ResizeObserver(() => {
      const rect = host.getBoundingClientRect();
      sizeRef.current = { w: Math.max(200, rect.width), h: Math.max(200, rect.height) };
      setSize(sizeRef.current);
    });
    ro.observe(host);
    const rect = host.getBoundingClientRect();
    sizeRef.current = { w: Math.max(200, rect.width), h: Math.max(200, rect.height) };
    setSize(sizeRef.current);
    return () => ro.disconnect();
  }, []);

  /* ---------------------------------------------------------- scroll logic */

  const scrollToCell = useCallback(
    (r: number, c: number) => {
      const m = metrics;
      const scroll = scrollRef.current;
      const cw = m.colWidths[c] ?? DEFAULT_COL_W * zoom;
      const rh = m.rowHeights[r] ?? DEFAULT_ROW_H * zoom;
      let changed = false;

      const left = m.colX[c];
      if (left < scroll.x) {
        scroll.x = left;
        changed = true;
      } else if (left + cw > scroll.x + bodyW) {
        scroll.x = left + cw - bodyW;
        changed = true;
      }
      const top = m.rowY[r];
      if (top < scroll.y) {
        scroll.y = top;
        changed = true;
      } else if (top + rh > scroll.y + bodyH) {
        scroll.y = top + rh - bodyH;
        changed = true;
      }
      if (changed) {
        scroll.x = clamp(scroll.x, 0, Math.max(0, m.totalW - bodyW));
        scroll.y = clamp(scroll.y, 0, Math.max(0, m.totalH - bodyH));
        scheduleDraw();
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [metrics, zoom, bodyW, bodyH],
  );

  const drawQueued = useRef(false);
  const scheduleDraw = useCallback(() => {
    if (drawQueued.current) return;
    drawQueued.current = true;
    requestAnimationFrame(() => {
      drawQueued.current = false;
      draw();
    });
  }, [draw]);

  useEffect(() => {
    scrollToCell(selection.cursor.r, selection.cursor.c);
  }, [selection.cursor.r, selection.cursor.c, scrollToCell]);

  /* ------------------------------------------------------------ hit testing */

  const cellAt = useCallback(
    (px: number, py: number): { r: number; c: number } | null => {
      if (px < HEADER_W || py < HEADER_H) return null;
      const scroll = scrollRef.current;
      const c = indexAt(metrics.colX, px - HEADER_W + scroll.x, viewCols);
      const r = indexAt(metrics.rowY, py - HEADER_H + scroll.y, viewRows);
      return { r, c };
    },
    [metrics, viewRows, viewCols],
  );

  const isHandle = (px: number, py: number): boolean => {
    const sel = normRange(selection.range);
    const scroll = scrollRef.current;
    const hx = HEADER_W + metrics.colX[sel.c2 + 1] - scroll.x;
    const hy = HEADER_H + metrics.rowY[sel.r2 + 1] - scroll.y;
    return Math.abs(px - hx) <= 5 && Math.abs(py - hy) <= 5;
  };

  const isHeaderCol = (px: number, py: number): number | null => {
    if (py > HEADER_H || px < HEADER_W) return null;
    const c = indexAt(metrics.colX, px - HEADER_W + scrollRef.current.x, viewCols);
    return c;
  };

  const isHeaderRow = (px: number, py: number): number | null => {
    if (px > HEADER_W || py < HEADER_H) return null;
    const r = indexAt(metrics.rowY, py - HEADER_H + scrollRef.current.y, viewRows);
    return r;
  };

  /* --------------------------------------------------------------- editing */

  const startEdit = useCallback(
    (r: number, c: number, initial: string) => {
      setEditing({ r, c, text: initial });
    },
    [],
  );

  const commitEdit = useCallback(
    (move: 'down' | 'right' | 'none' = 'none') => {
      if (!editing) return;
      const { r, c, text } = editing;
      setEditing(null);
      engine.transact(() => engine.setInput(sheet, r, c, text));
      engine.recalc();

      if (move === 'down') onSelectionChange({ cursor: { r: r + 1, c }, range: { r1: r + 1, c1: c, r2: r + 1, c2: c } });
      else if (move === 'right') onSelectionChange({ cursor: { r, c: c + 1 }, range: { r1: r, c1: c + 1, r2: r, c2: c + 1 } });
    },
    [editing, engine, sheet, onSelectionChange],
  );

  const cancelEdit = useCallback(() => setEditing(null), []);

  useEffect(() => {
    if (editing && editorRef.current) {
      const el = editorRef.current;
      el.focus();
      el.setSelectionRange(el.value.length, el.value.length);
    }
  }, [editing]);

  // Formula-bar driven edits.
  useEffect(() => {
    if (!editSignal) return;
    const { r, c, text, mode } = editSignal;
    startEdit(r, c, mode === 'append' ? `${text}${editSignal.text}` : text);
    onEditSignalHandled();
  }, [editSignal, onEditSignalHandled, startEdit]);

  /* ------------------------------------------------------------- clipboard */

  const doCopy = useCallback(
    async (cut: boolean) => {
      const sel = normRange(selection.range);
      const payload: ClipPayload = {
        sheetName: sheet.name,
        range: sel,
        cells: [],
      };
      for (let r = sel.r1; r <= sel.r2; r++) {
        for (let c = sel.c1; c <= sel.c2; c++) {
          payload.cells.push({
            r,
            c,
            text: engine.editText(sheet, r, c),
            styleId: sheet.cells.get(pack(r, c))?.styleId,
          });
        }
      }
      setClip(payload);
      await copyToClipboard(payload);
      if (cut) {
        engine.transact(() => engine.clearRange(sheet, sel));
        engine.recalc();
      }
    },
    [selection.range, sheet, engine],
  );

  const doPaste = useCallback(
    async (target?: { r: number; c: number }) => {
      const at = target ?? selection.cursor;
      let payload = clip;
      if (!payload) payload = await readClipboardText();
      if (!payload || !payload.cells.length) return;

      const minR = Math.min(...payload.cells.map((x) => x.r));
      const minC = Math.min(...payload.cells.map((x) => x.c));
      const dr = at.r - minR;
      const dc = at.c - minC;

      const height = Math.max(...payload.cells.map((x) => x.r)) - minR + 1;
      const width = Math.max(...payload.cells.map((x) => x.c)) - minC + 1;

      engine.transact(() => {
        for (const cell of payload.cells) {
          const r = cell.r + dr;
          const c = cell.c + dc;
          if (r >= MAX_ROWS || c >= MAX_COLS) continue;
          engine.setInput(sheet, r, c, translateIfNeeded(cell.text, dr, dc));
          if (cell.styleId !== undefined) engine.setStyleId(sheet, r, c, cell.styleId);
        }
      });
      engine.recalc();

      onSelectionChange({
        cursor: at,
        range: { r1: at.r, c1: at.c, r2: at.r + height - 1, c2: at.c + width - 1 },
      });
    },
    [clip, selection.cursor, sheet, engine, onSelectionChange],
  );

  /* -------------------------------------------------------- mouse handlers */

  const onMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (e.button === 2) return;
    const rect = canvasRef.current!.getBoundingClientRect();
    const px = e.clientX - rect.left;
    const py = e.clientY - rect.top;

    if (editing) {
      commitEdit('none');
    }

    // fill handle -> drag a copy
    if (isHandle(px, py)) {
      dragRef.current = { mode: 'fill', anchor: { ...selection.cursor }, origin: { x: px, y: py }, moved: false };
      setFillPreview(normRange(selection.range));
      return;
    }

    const headerRow = isHeaderRow(px, py);
    if (headerRow !== null) {
      dragRef.current = { mode: 'rows', anchor: { r: headerRow, c: 0 }, origin: { x: px, y: py }, moved: false };
      onSelectionChange({ cursor: { r: headerRow, c: selection.cursor.c }, range: { r1: headerRow, c1: 0, r2: headerRow, c2: viewCols - 1 } });
      return;
    }
    const headerCol = isHeaderCol(px, py);
    if (headerCol !== null) {
      dragRef.current = { mode: 'cols', anchor: { r: 0, c: headerCol }, origin: { x: px, y: py }, moved: false };
      onSelectionChange({ cursor: { r: selection.cursor.r, c: headerCol }, range: { r1: 0, c1: headerCol, r2: viewRows - 1, c2: headerCol } });
      return;
    }

    const hit = cellAt(px, py);
    if (!hit) return;

    const shift = e.shiftKey;
    if (shift) {
      const merged = {
        r1: Math.min(selection.cursor.r, hit.r),
        r2: Math.max(selection.cursor.r, hit.r),
        c1: Math.min(selection.cursor.c, hit.c),
        c2: Math.max(selection.cursor.c, hit.c),
      };
      onSelectionChange({ cursor: hit, range: merged });
    } else {
      onSelectionChange({ cursor: hit, range: { r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c } });
    }
    dragRef.current = { mode: 'cells', anchor: hit, origin: { x: px, y: py }, moved: false };
  };

  const onMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const drag = dragRef.current;
    if (!drag) return;
    const rect = canvasRef.current!.getBoundingClientRect();
    const px = e.clientX - rect.left;
    const py = e.clientY - rect.top;

    // auto-scroll when dragging past the edge
    const scroll = scrollRef.current;
    const edge = 24;
    if (px > size.w - edge) scroll.x = Math.min(metrics.totalW - bodyW, scroll.x + 30);
    else if (px < HEADER_W + edge) scroll.x = Math.max(0, scroll.x - 30);
    if (py > size.h - edge) scroll.y = Math.min(metrics.totalH - bodyH, scroll.y + 30);
    else if (py < HEADER_H + edge) scroll.y = Math.max(0, scroll.y - 30);

    const hit = cellAt(clamp(px, HEADER_W, size.w), clamp(py, HEADER_H, size.h)) ?? drag.anchor;
    drag.moved = true;

    if (drag.mode === 'cells') {
      onSelectionChange({
        cursor: selection.cursor,
        range: {
          r1: Math.min(drag.anchor.r, hit.r),
          r2: Math.max(drag.anchor.r, hit.r),
          c1: Math.min(drag.anchor.c, hit.c),
          c2: Math.max(drag.anchor.c, hit.c),
        },
      });
    } else if (drag.mode === 'fill') {
      const sel = normRange(selection.range);
      setFillPreview({
        r1: Math.min(sel.r1, hit.r),
        r2: Math.max(sel.r2, hit.r),
        c1: Math.min(sel.c1, hit.c),
        c2: Math.max(sel.c2, hit.c),
      });
      setDropPreview(hit);
    } else if (drag.mode === 'rows') {
      onSelectionChange({
        cursor: selection.cursor,
        range: { r1: Math.min(drag.anchor.r, hit.r), r2: Math.max(drag.anchor.r, hit.r), c1: 0, c2: viewCols - 1 },
      });
    } else if (drag.mode === 'cols') {
      onSelectionChange({
        cursor: selection.cursor,
        range: { r1: 0, r2: viewRows - 1, c1: Math.min(drag.anchor.c, hit.c), c2: Math.max(drag.anchor.c, hit.c) },
      });
    }
    scheduleDraw();
  };

  const onMouseUp = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const drag = dragRef.current;
    dragRef.current = null;
    if (!drag) return;

    if (drag.mode === 'fill' && fillPreview) {
      const sel = normRange(selection.range);
      const rect = canvasRef.current!.getBoundingClientRect();
      const hit = cellAt(e.clientX - rect.left, e.clientY - rect.top) ?? drag.anchor;

      // Extend downwards or rightwards, tiling the source block.
      if (hit.r > sel.r2 && sel.c1 === sel.c2) {
        engine.transact(() =>
          engine.copyRange(sheet, { r1: sel.r2 + 1, c1: sel.c1, r2: hit.r, c2: sel.c2 }, sel),
        );
        onSelectionChange({ cursor: hit, range: { r1: sel.r1, c1: sel.c1, r2: hit.r, c2: sel.c2 } });
      } else if (hit.c > sel.c2 && sel.r1 === sel.r2) {
        engine.transact(() =>
          engine.copyRange(sheet, { r1: sel.r1, c1: sel.c2 + 1, r2: sel.r2, c2: hit.c }, sel),
        );
        onSelectionChange({ cursor: hit, range: { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: hit.c } });
      } else if (hit.r < sel.r1 || hit.c < sel.c1) {
        engine.transact(() => engine.copyRange(sheet, sel, fillPreview, 'copy'));
        onSelectionChange({ cursor: drag.anchor, range: sel });
      }
      engine.recalc();
    }

    setFillPreview(null);
    setDropPreview(null);
    scheduleDraw();
  };

  const onDoubleClick = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const rect = canvasRef.current!.getBoundingClientRect();
    const hit = cellAt(e.clientX - rect.left, e.clientY - rect.top);
    if (!hit) return;
    startEdit(hit.r, hit.c, engine.editText(sheet, hit.r, hit.c));
  };

  const onContextMenu = (e: React.MouseEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    const rect = canvasRef.current!.getBoundingClientRect();
    const hit = cellAt(e.clientX - rect.left, e.clientY - rect.top);
    if (hit) {
      onSelectionChange({ cursor: hit, range: { r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c } });
      scheduleDraw();
    }
  };

  // Native drag-and-drop for moving a block.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const onDragOver = (e: DragEvent) => {
      e.preventDefault();
      const rect = canvas.getBoundingClientRect();
      const hit = cellAt(e.clientX - rect.left, e.clientY - rect.top);
      if (hit) setDropPreview(hit);
    };
    const onDrop = (e: DragEvent) => {
      e.preventDefault();
      const rect = canvas.getBoundingClientRect();
      const hit = cellAt(e.clientX - rect.left, e.clientY - rect.top);
      setDropPreview(null);
      if (!hit) return;
      void doPaste(hit);
    };
    canvas.addEventListener('dragover', onDragOver);
    canvas.addEventListener('drop', onDrop);
    return () => {
      canvas.removeEventListener('dragover', onDragOver);
      canvas.removeEventListener('drop', onDrop);
    };
  }, [cellAt, doPaste]);

  /* --------------------------------------------------------------- keyboard */

  const moveCursor = (dr: number, dc: number, extend: boolean) => {
    const r = clamp(selection.cursor.r + dr, 0, MAX_ROWS - 1);
    const c = clamp(selection.cursor.c + dc, 0, MAX_COLS - 1);
    if (extend) {
      const sel = normRange(selection.range);
      onSelectionChange({
        cursor: { r, c },
        range: {
          r1: Math.min(sel.r1, r),
          r2: Math.max(sel.r2, r),
          c1: Math.min(sel.c1, c),
          c2: Math.max(sel.c2, c),
        },
      });
    } else {
      onSelectionChange({ cursor: { r, c }, range: { r1: r, c1: c, r2: r, c2: c } });
    }
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (editing) return; // the textarea handles its own keys
    const ctrl = e.ctrlKey || e.metaKey;
    const { r, c } = selection.cursor;
    const sel = normRange(selection.range);

    switch (e.key) {
      case 'ArrowUp': e.preventDefault(); moveCursor(-1, 0, e.shiftKey); return;
      case 'ArrowDown': e.preventDefault(); moveCursor(1, 0, e.shiftKey); return;
      case 'ArrowLeft': e.preventDefault(); moveCursor(0, -1, e.shiftKey); return;
      case 'ArrowRight': e.preventDefault(); moveCursor(0, 1, e.shiftKey); return;
      case 'Tab':
        e.preventDefault();
        moveCursor(0, e.shiftKey ? -1 : 1, false);
        return;
      case 'Enter':
        e.preventDefault();
        if (e.altKey) {
          startEdit(r, c, engine.editText(sheet, r, c));
        } else {
          moveCursor(e.shiftKey ? -1 : 1, 0, false);
        }
        return;
      case 'F2':
        e.preventDefault();
        startEdit(r, c, engine.editText(sheet, r, c));
        return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        engine.transact(() => engine.clearRange(sheet, sel));
        engine.recalc();
        return;
      case 'Home':
        e.preventDefault();
        onSelectionChange({ cursor: { r, c: 0 }, range: { r1: r, c1: 0, r2: r, c2: 0 } });
        return;
      case 'PageDown':
        e.preventDefault();
        moveCursor(Math.floor(bodyH / DEFAULT_ROW_H), 0, e.shiftKey);
        return;
      case 'PageUp':
        e.preventDefault();
        moveCursor(-Math.floor(bodyH / DEFAULT_ROW_H), 0, e.shiftKey);
        return;
      default:
        break;
    }

    if (ctrl) {
      switch (e.key.toLowerCase()) {
        case 'c': e.preventDefault(); void doCopy(false); return;
        case 'x': e.preventDefault(); void doCopy(true); return;
        case 'v': e.preventDefault(); void doPaste(); return;
        case 'a':
          e.preventDefault();
          onSelectionChange({ cursor: { r, c }, range: { r1: 0, c1: 0, r2: viewRows - 1, c2: viewCols - 1 } });
          return;
        case 'z': e.preventDefault(); if (engine.undo()) { engine.recalc(); } return;
        case 'y': e.preventDefault(); if (engine.redo()) { engine.recalc(); } return;
        case 'b': e.preventDefault(); onFormat('bold'); return;
        case 'i': e.preventDefault(); onFormat('italic'); return;
        case 'u': e.preventDefault(); onFormat('underline'); return;
        default: return;
      }
    }

    // any printable character starts an edit, replacing the cell
    if (e.key.length === 1 && !e.altKey) {
      e.preventDefault();
      startEdit(r, c, e.key);
    }
  };

  /* ------------------------------------------------------------- formatting */

  const onFormat = useCallback(
    (key: keyof CellFormat | 'numFmt', value?: string) => {
      const sel = normRange(selection.range);
      engine.transact(() => {
        for (let r = sel.r1; r <= sel.r2; r++) {
          for (let c = sel.c1; c <= sel.c2; c++) {
            const existing = engine.formatOf(sheet.cells.get(pack(r, c))?.styleId) ?? {};
            const next: Record<string, unknown> = { ...existing };
            if (value === undefined) {
              if (next[key]) delete next[key];
              else next[key] = 1;
            } else if (value === '') {
              delete next[key];
            } else {
              next[key] = value;
            }
            engine.setStyleId(sheet, r, c, engine.styleIdFor(next as CellFormat));
          }
        }
      });
      engine.emit();
    },
    [engine, sheet, selection.range],
  );

  /* exposed for the ribbon, via the parent */
  useEffect(() => {
    const handler = (e: Event) => {
      const detail = (e as CustomEvent).detail as { key: string; value?: string };
      onFormat(detail.key as keyof CellFormat, detail.value);
    };
    window.addEventListener('sheetcraft:format', handler);
    return () => window.removeEventListener('sheetcraft:format', handler);
  }, [onFormat]);

  /* ------------------------------------------------------------------ wheel */

  const onWheel = (e: React.WheelEvent<HTMLCanvasElement>) => {
    const scroll = scrollRef.current;
    if (e.ctrlKey) {
      e.preventDefault();
      return;
    }
    if (e.deltaX !== 0 || Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
      scroll.x = clamp(scroll.x + e.deltaX, 0, Math.max(0, metrics.totalW - bodyW));
      scroll.y = clamp(scroll.y + e.deltaY, 0, Math.max(0, metrics.totalH - bodyH));
      scheduleDraw();
    }
  };

  /* ------------------------------------------------------------ editor box */

  const editorBox = useMemo(() => {
    if (!editing) return null;
    const scroll = scrollRef.current;
    const w = Math.max(metrics.colWidths[editing.c] ?? 120, 120);
    const h = Math.max(metrics.rowHeights[editing.r] ?? 22, 22);
    return {
      left: HEADER_W + metrics.colX[editing.c] - scroll.x,
      top: HEADER_H + metrics.rowY[editing.r] - scroll.y,
      width: w,
      height: h,
    };
  }, [editing, metrics]);

  return (
    <div className="grid-host" ref={hostRef}>
      <canvas
        ref={canvasRef}
        className="grid-canvas"
        tabIndex={0}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseUp={onMouseUp}
        onMouseLeave={onMouseUp}
        onDoubleClick={onDoubleClick}
        onContextMenu={onContextMenu}
        onKeyDown={onKeyDown}
        onWheel={onWheel}
        draggable
        onDragStart={() => {
          void doCopy(false);
        }}
      />
      {editing && editorBox && (
        <textarea
          ref={editorRef}
          className={`cell-editor${editing.text.startsWith('=') ? '' : ' text'}`}
          style={{
            left: editorBox.left,
            top: editorBox.top,
            width: editorBox.width,
            height: editorBox.height,
          }}
          value={editing.text}
          onChange={(e) => setEditing({ ...editing, text: e.target.value })}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.altKey) {
              e.preventDefault();
              commitEdit(e.shiftKey ? 'up' as 'none' : 'down');
            } else if (e.key === 'Escape') {
              e.preventDefault();
              cancelEdit();
            } else if (e.key === 'Tab') {
              e.preventDefault();
              commitEdit('right');
            }
            e.stopPropagation();
          }}
          onBlur={() => commitEdit('none')}
        />
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ helpers */

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

function fontFor(fmt: CellFormat | undefined): string {
  const size = fmt?.size ?? 12.5;
  const style = fmt?.italic ? 'italic ' : '';
  const weight = fmt?.bold ? '600 ' : '';
  const family = fmt?.font ? `"${fmt.font}", ` : '';
  return `${style}${weight}${size}px ${family}'Segoe UI', system-ui, sans-serif`;
}

/** Re-point relative references when a pasted formula moves. */
function translateIfNeeded(text: string, dr: number, dc: number): string {
  if (!text.startsWith('=') || (dr === 0 && dc === 0)) return text;
  return `=${translateFormula(text.slice(1), dr, dc)}`;
}
