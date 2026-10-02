/**
 * Chart data extraction.
 *
 * A chart stores A1-style references rather than copied values, so editing the
 * source range updates the picture. This module resolves those references into
 * plain numbers/labels for rendering.
 */

import { isError, type ChartSeries, type ChartSpec, type Sheet, type Value } from './types';
import { parseRangeExpr, splitRef } from './ref';
import { formatGeneral, toText } from './coerce';

export interface ChartPoint {
  label: string;
  value: number;
}

export interface ChartData {
  title: string;
  type: ChartSpec['type'];
  /** one entry per series */
  series: { name: string; points: ChartPoint[] }[];
  /** union of category labels across series, in order */
  categories: string[];
  max: number;
  min: number;
}

export interface Resolver {
  sheetByName(name: string | null, current: Sheet): Sheet | null;
  value(sheet: Sheet, r: number, c: number): Value;
}

/** Read a range reference into its cell values. */
export function readRef(
  ref: string,
  current: Sheet,
  resolve: Resolver,
): { values: Value[]; width: number } {
  const parts = splitRef(ref);
  if (!parts) return { values: [], width: 0 };
  const parsed = parseRangeExpr(parts.a1);
  if (!parsed) return { values: [], width: 0 };
  const sheet = resolve.sheetByName(parsed.sheet, current);
  if (!sheet) return { values: [], width: 0 };

  const width = parsed.c2 - parsed.c1 + 1;
  const values: Value[] = [];
  for (let r = parsed.r1; r <= parsed.r2; r++) {
    for (let c = parsed.c1; c <= parsed.c2; c++) values.push(resolve.value(sheet, r, c));
  }
  return { values, width };
}

/**
 * Resolve a chart into renderable series.
 *
 * Numeric data is read from the values reference; when a categories reference
 * is supplied it provides the axis labels. Single-column ranges become one
 * series per column so a multi-column table plots as a grouped chart.
 */
export function buildChartData(
  chart: ChartSpec,
  current: Sheet,
  resolve: Resolver,
): ChartData {
  const series: { name: string; points: ChartPoint[] }[] = [];
  const categoryLabels: string[] = [];

  const catsRef = chart.series.find((s) => s.categories)?.categories;
  if (catsRef) {
    const cats = readRef(catsRef, current, resolve);
    for (let i = 0; i < cats.values.length; i++) categoryLabels.push(label(cats.values[i], i));
  }

  for (const seriesDef of chart.series) {
    if (!seriesDef.values) continue;
    const { values } = readRef(seriesDef.values, current, resolve);

    // A single reference covering several columns yields one series per column.
    const columns = columnCount(seriesDef.values);
    if (columns > 1) {
      for (let col = 0; col < columns; col++) {
        const points: ChartPoint[] = [];
        for (let row = 0; row * columns + col < values.length; row++) {
          const v = values[row * columns + col];
          points.push({
            label: categoryLabels[row] ?? label(v, points.length),
            value: numeric(v),
          });
        }
        series.push({ name: seriesDef.name || columnHeader(resolve, current, seriesDef.values, col), points });
      }
      continue;
    }

    const points: ChartPoint[] = values.map((v, i) => ({
      label: categoryLabels[i] ?? label(v, i),
      value: numeric(v),
    }));
    series.push({ name: seriesDef.name || 'Series', points });
  }

  let max = 0;
  let min = 0;
  for (const s of series) {
    for (const p of s.points) {
      if (!Number.isFinite(p.value)) continue;
      if (p.value > max) max = p.value;
      if (p.value < min) min = p.value;
    }
  }
  if (max === min) max = min + 1;

  const categories = series[0]?.points.map((p) => p.label) ?? categoryLabels;
  return { title: chart.title, type: chart.type, series, categories, max, min };
}

function columnCount(ref: string): number {
  const parsed = parseRangeExpr(splitRef(ref)?.a1 ?? '');
  if (!parsed) return 1;
  return parsed.c2 - parsed.c1 + 1;
}

function columnHeader(
  resolve: Resolver,
  current: Sheet,
  ref: string,
  col: number,
): string {
  const parts = splitRef(ref);
  const parsed = parts ? parseRangeExpr(parts.a1) : null;
  if (!parsed) return `Series ${col + 1}`;
  const sheet = resolve.sheetByName(parsed.sheet, current);
  if (!sheet) return `Series ${col + 1}`;
  // Header rows are conventional: use the first row of the block.
  const v = resolve.value(sheet, parsed.r1, parsed.c1 + col);
  return v === null ? `Series ${col + 1}` : toText(v);
}

function numeric(v: Value): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === null || isError(v)) return 0;
  const n = Number(v.replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function label(v: Value, i: number): string {
  if (v === null) return String(i + 1);
  if (isError(v)) return '#ERR';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return formatGeneral(v);
}

/** Palette used for series colours. */
export const SERIES_COLORS = [
  '#2f6fed', '#e8833a', '#22a06b', '#c2436b', '#8b5cf6',
  '#0ea5b7', '#d4a017', '#64748b', '#db2777', '#16a34a',
];

/** Default series list for a selected data block. */
export function inferSeries(
  dataRange: string,
  categoryRange: string | undefined,
  useFirstRowAsLabels: boolean,
  useFirstColumnAsCategories: boolean,
): ChartSeries[] {
  const parsed = parseRangeExpr(splitRef(dataRange)?.a1 ?? '');
  if (!parsed) return [];
  const series: ChartSeries[] = [];

  if (useFirstRowAsLabels && useFirstColumnAsCategories) {
    for (let c = parsed.c1; c <= parsed.c2; c++) {
      series.push({ values: `${splitRef(dataRange)?.sheet ?? ''}!${rangeRef(parsed.r1 + 1, c, parsed.r2, c)}` });
    }
  } else if (useFirstRowAsLabels) {
    for (let c = parsed.c1; c <= parsed.c2; c++) {
      series.push({ values: rangeRef(parsed.r1, c, parsed.r2, c) });
    }
  } else {
    series.push({ values: dataRange });
  }

  if (categoryRange && useFirstColumnAsCategories) {
    for (const s of series) s.categories = categoryRange;
  }
  return series;
}

function rangeRef(r1: number, c1: number, r2: number, c2: number): string {
  const { colName } = colHelpers;
  return r1 === r2 && c1 === c2
    ? `${colName(c1)}${r1 + 1}`
    : `${colName(c1)}${r1 + 1}:${colName(c2)}${r2 + 1}`;
}

import { colIndex, colName as colNameImpl } from './ref';
const colHelpers = { colName: colNameImpl, colIndex };

/** Create a chart with sensible defaults. */
export function newChart(sheet: Sheet, type: ChartSpec['type'], series: ChartSeries[], anchor: { r: number; c: number }): ChartSpec {
  return {
    id: `ch_${Math.random().toString(36).slice(2, 10)}`,
    name: `Chart ${sheet.charts.length + 1}`,
    type,
    title: 'Chart',
    sheetId: sheet.id,
    anchor,
    width: 480,
    height: 300,
    series,
    showLegend: true,
    showTitle: true,
    showGridlines: true,
    showDataLabels: false,
    stacked: false,
    smooth: false,
  };
}

export type { ChartSeries, ChartSpec };
