/**
 * Side panes: charts, data cleaning, dashboard and SQL.
 *
 * Each pane owns its own controls and talks to the engine through the same
 * undoable operations the grid uses.
 */

import { useMemo, useState } from 'react';
import type { ChartSpec, Range, Sheet, Value } from '../core/types';
import { isError } from '../core/types';
import { formatGeneral, toText } from '../core/coerce';
import { normRange, rangeToA1, splitRef, quoteSheet } from '../core/ref';
import { buildChartData, newChart, type Resolver } from '../core/chartdata';
import {
  applyTransform,
  padZeros,
  profileColumns,
  removeDuplicateRows,
  removeEmptyRows,
  replaceErrors,
  splitColumn,
  transforms,
  type ColumnProfile,
} from '../core/dataclean';
import { SqliteFile, loadResult, type TableInfo } from '../core/sqlite';
import type { Engine } from '../core/engine';
import ChartView from './ChartView';

/* ==================================================================== charts */

export function ChartsPane({
  engine,
  sheet,
  selection,
}: {
  engine: Engine;
  sheet: Sheet;
  selection: Range;
}) {
  const [revision, setRevision] = useState(0);
  const [chartType, setChartType] = useState<ChartSpec['type']>('column');
  const [useHeaders, setUseHeaders] = useState(true);
  const [useCategories, setUseCategories] = useState(true);

  const resolver: Resolver = useMemo(
    () => ({
      sheetByName: (name, current) => (name === null ? current : engine.sheetByName(name)),
      value: (s, r, c) => engine.valueAt(s, r, c),
    }),
    [engine],
  );

  const charts = sheet.charts;
  const range = normRange(selection);

  const refresh = () => setRevision((n) => n + 1);

  const addChart = () => {
    const dataRange = `${quoteSheet(sheet.name)}!${rangeToA1(range, true)}`;
    let categories: string | undefined;
    if (useCategories && range.c2 > range.c1) {
      categories = `${quoteSheet(sheet.name)}!${rangeToA1(
        { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c1 },
        true,
      )}`;
    }

    // A single numeric column is the common case; split by column otherwise.
    const series = buildSeries(engine, sheet, range, useHeaders, useCategories);
    const chart = newChart(sheet, chartType, series, { r: 0, c: 0 });
    chart.title = `Chart of ${rangeToA1(range)}`;
    chart.series = series;
    if (categories) for (const s of chart.series) s.categories = categories;
    engine.transact(() => sheet.charts.push(chart));
    engine.emit();
    refresh();
  };

  const update = (id: string, patch: Partial<ChartSpec>) => {
    const chart = sheet.charts.find((c) => c.id === id);
    if (!chart) return;
    engine.transact(() => Object.assign(chart, patch));
    engine.emit();
    refresh();
  };

  const remove = (id: string) => {
    engine.transact(() => {
      sheet.charts = sheet.charts.filter((c) => c.id !== id);
    });
    engine.emit();
    refresh();
  };

  return (
    <div className="pane-body">
      <div className="pane-section">
        <h4>Insert chart</h4>
        <div className="row">
          <label>Range</label>
          <input type="text" value={rangeToA1(range)} readOnly />
        </div>
        <div className="row">
          <label>Type</label>
          <select value={chartType} onChange={(e) => setChartType(e.target.value as ChartSpec['type'])}>
            <option value="column">Column</option>
            <option value="bar">Bar</option>
            <option value="line">Line</option>
            <option value="area">Area</option>
            <option value="pie">Pie</option>
            <option value="doughnut">Doughnut</option>
            <option value="scatter">Scatter</option>
            <option value="radar">Radar</option>
          </select>
        </div>
        <div className="row">
          <label>
            <input type="checkbox" checked={useHeaders} onChange={(e) => setUseHeaders(e.target.checked)} /> First row is a header
          </label>
        </div>
        <div className="row">
          <label>
            <input type="checkbox" checked={useCategories} onChange={(e) => setUseCategories(e.target.checked)} /> First column is a category
          </label>
        </div>
        <button className="btn primary" onClick={addChart} disabled={rangeCells(range) === 0}>
          Add chart
        </button>
      </div>

      <div className="divider" />

      <div className="pane-section">
        <h4>Charts on this sheet</h4>
        {charts.length === 0 ? (
          <div className="empty-note">Select a range of data, then add a chart.</div>
        ) : (
          charts.map((chart) => (
            <ChartView
              key={`${chart.id}-${revision}`}
              chart={chart}
              data={buildChartData(chart, sheet, resolver)}
              onTypeChange={(t) => update(chart.id, { type: t })}
              onToggle={(k) => update(chart.id, { [k]: !chart[k] } as Partial<ChartSpec>)}
              onDelete={() => remove(chart.id)}
            />
          ))
        )}
      </div>
    </div>
  );
}

/** Split a data block into one series per column. */
function buildSeries(
  engine: Engine,
  sheet: Sheet,
  range: Range,
  useHeaders: boolean,
  useCategories: boolean,
) {
  const sheetRef = quoteSheet(sheet.name);
  const firstDataRow = useHeaders ? range.r1 + 1 : range.r1;
  const firstDataCol = useCategories ? range.c1 + 1 : range.c1;

  if (firstDataRow > range.r2 || firstDataCol > range.c2) {
    return [{ values: `${sheetRef}!${rangeToA1(range, true)}` }];
  }

  const series = [];
  for (let c = firstDataCol; c <= range.c2; c++) {
    series.push({
      name: useHeaders ? toText(engine.valueAt(sheet, range.r1, c)) : '',
      values: `${sheetRef}!${rangeToA1({ r1: firstDataRow, c1: c, r2: range.r2, c2: c }, true)}`,
    });
  }
  return series.length ? series : [{ values: `${sheetRef}!${rangeToA1(range, true)}` }];
}

function rangeCells(r: Range): number {
  const n = normRange(r);
  return (n.r2 - n.r1 + 1) * (n.c2 - n.c1 + 1);
}

/* =================================================================== cleaner */

export function DataPane({
  engine,
  sheet,
  selection,
  onMessage,
}: {
  engine: Engine;
  sheet: Sheet;
  selection: Range;
  onMessage: (msg: string) => void;
}) {
  const [revision, setRevision] = useState(0);
  const [hasHeader, setHasHeader] = useState(true);
  const range = normRange(selection);
  const profiles = useMemo(
    () => profileColumns(engine, sheet, range, { hasHeader }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [engine, sheet, range, revision, hasHeader],
  );

  const run = (fn: () => { description?: string; changed?: number; count?: number }) => {
    const result = fn();
    engine.emit();
    setRevision((n) => n + 1);
    const n = result.changed ?? result.count ?? 0;
    onMessage(`${result.description ?? 'Done'} (${n})`);
  };

  const apply = (label: string, fn: (v: Value) => Value) =>
    run(() => applyTransform(engine, sheet, range, label, fn));

  return (
    <div className="pane-body">
      <div className="pane-section">
        <h4>Working range</h4>
        <div className="row">
          <label>Selection</label>
          <input type="text" value={rangeToA1(range)} readOnly />
          <span className="pill">{rangeCells(range)} cells</span>
        </div>
        <div className="row">
          <label>
            <input type="checkbox" checked={hasHeader} onChange={(e) => setHasHeader(e.target.checked)} /> First row is a header
          </label>
        </div>
        <p className="hint">Every action below applies to the selected range and can be undone.</p>
      </div>

      <div className="pane-section">
        <h4>Clean text</h4>
        <div className="row wrap">
          <button className="btn" onClick={() => apply('Trim', transforms.trim)}>Trim</button>
          <button className="btn" onClick={() => apply('Collapse spaces', transforms.collapseSpaces)}>Collapse spaces</button>
          <button className="btn" onClick={() => apply('UPPER', transforms.upper)}>UPPER</button>
          <button className="btn" onClick={() => apply('lower', transforms.lower)}>lower</button>
          <button className="btn" onClick={() => apply('Proper', transforms.proper)}>Proper</button>
          <button className="btn" onClick={() => apply('Digits only', transforms.digitsOnly)}>Digits only</button>
          <button className="btn" onClick={() => apply('Letters only', transforms.lettersOnly)}>Letters only</button>
        </div>
      </div>

      <div className="pane-section">
        <h4>Convert types</h4>
        <div className="row wrap">
          <button className="btn" onClick={() => apply('To number', transforms.toNumber)}>To number</button>
          <button className="btn" onClick={() => apply('To date', transforms.toDateSerial)}>To date</button>
          <button className="btn" onClick={() => apply('Strip symbols', transforms.stripNonNumeric)}>Strip symbols</button>
          <button className="btn" onClick={() => apply('0 to blank', transforms.zeroToBlank)}>0 → blank</button>
          <button className="btn" onClick={() => apply('Blank to 0', transforms.blankToZero)}>blank → 0</button>
        </div>
      </div>

      <div className="pane-section">
        <h4>Rows &amp; errors</h4>
        <div className="row wrap">
          <button className="btn" onClick={() => run(() => ({ description: 'Removed empty rows', count: removeEmptyRows(engine, sheet, range) }))}>
            Remove empty rows
          </button>
          <button className="btn" onClick={() => run(() => ({ description: 'Removed duplicates', count: removeDuplicateRows(engine, sheet, range) }))}>
            Remove duplicates
          </button>
          <button className="btn" onClick={() => run(() => ({ description: 'Replaced errors', count: replaceErrors(engine, sheet, range, '') }))}>
            Clear errors
          </button>
        </div>
      </div>

      <div className="pane-section">
        <h4>Reshape</h4>
        <div className="row">
          <label>Split on</label>
          <input type="text" defaultValue="," style={{ width: 60 }} id="split-sep" />
          <button
            className="btn"
            onClick={() => {
              const el = document.getElementById('split-sep') as HTMLInputElement | null;
              const sep = el?.value || ',';
              run(() => ({ description: `Split on "${sep}"`, count: splitColumn(engine, sheet, range, sep) }));
            }}
          >
            Split column
          </button>
        </div>
        <div className="row">
          <label>Pad to</label>
          <input
            type="number"
            defaultValue={6}
            style={{ width: 70 }}
            id="pad-width"
            onClick={() => {
              const el = document.getElementById('pad-width') as HTMLInputElement | null;
              const width = Number(el?.value || 6);
              run(() => ({ description: `Padded to ${width}`, count: padZeros(engine, sheet, range, width) }));
            }}
          />
          <span className="hint" style={{ margin: 0 }}>click to apply</span>
        </div>
      </div>

      <div className="pane-section">
        <h4>Column profile</h4>
        <table className="data">
          <thead>
            <tr>
              <th>Column</th>
              <th>Type</th>
              <th className="num">Filled</th>
              <th className="num">Blank</th>
              <th className="num">Dupes</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody>
            {profiles.map((p: ColumnProfile) => (
              <tr key={p.index}>
                <td>{p.header || `Column ${p.index + 1}`}</td>
                <td>
                  <span className={`pill ${p.kind === 'mixed' || p.errors > 0 ? 'warn' : 'ok'}`}>{p.kind}</span>
                </td>
                <td className="num">{p.total - p.blanks}</td>
                <td className="num">{p.blanks}</td>
                <td className="num">{p.duplicates}</td>
                <td className="hint">{p.suggestion}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ================================================================= dashboard */

export function DashboardPane({
  engine,
  sheet,
  selection,
}: {
  engine: Engine;
  sheet: Sheet;
  selection: Range;
}) {
  const range = normRange(selection);
  const stats = useMemo(() => {
    const numbers: number[] = [];
    const dates: number[] = [];
    let filled = 0;
    let blanks = 0;
    let errors = 0;
    let text = 0;
    const seen = new Set<string>();

    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const v = engine.valueAt(sheet, r, c);
        if (v === null || v === '') {
          blanks++;
          continue;
        }
        filled++;
        if (isError(v)) {
          errors++;
          continue;
        }
        seen.add(toText(v).trim().toLowerCase());
        if (typeof v === 'number') {
          numbers.push(v);
          if (v > 20000 && v < 80000 && Number.isInteger(v)) dates.push(v);
        } else text++;
      }
    }

    const sorted = [...numbers].sort((a, b) => a - b);
    const sum = numbers.reduce((s, v) => s + v, 0);
    const mean = numbers.length ? sum / numbers.length : 0;
    const variance = numbers.length > 1
      ? numbers.reduce((s, v) => s + (v - mean) ** 2, 0) / (numbers.length - 1)
      : 0;
    const quant = (p: number): number => {
      if (!sorted.length) return 0;
      const idx = (sorted.length - 1) * p;
      const lo = Math.floor(idx);
      const hi = Math.ceil(idx);
      return lo === hi ? sorted[lo] : sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
    };

    return {
      filled,
      blanks,
      errors,
      text,
      uniques: seen.size,
      count: numbers.length,
      sum,
      mean,
      stdev: Math.sqrt(variance),
      min: sorted.length ? sorted[0] : 0,
      max: sorted.length ? sorted[sorted.length - 1] : 0,
      median: quant(0.5),
      p25: quant(0.25),
      p75: quant(0.75),
      dates: dates.length,
      histogram: histogram(numbers),
    };
  }, [engine, sheet, range]);

  return (
    <div className="pane-body">
      <div className="pane-section">
        <h4>Selection</h4>
        <div className="row">
          <input type="text" value={`${sheet.name}!${rangeToA1(range)}`} readOnly />
          <span className="pill">{rangeCells(range)} cells</span>
        </div>
      </div>

      <div className="pane-section">
        <h4>Data quality</h4>
        <div className="kpi-grid">
          <Kpi k="Filled" v={String(stats.filled)} />
          <Kpi k="Blank" v={String(stats.blanks)} />
          <Kpi k="Errors" v={String(stats.errors)} />
          <Kpi k="Text cells" v={String(stats.text)} />
          <Kpi k="Unique values" v={String(stats.uniques)} />
          <Kpi k="Date-like" v={String(stats.dates)} />
        </div>
      </div>

      {stats.count > 0 && (
        <>
          <div className="pane-section">
            <h4>Numeric summary</h4>
            <div className="kpi-grid">
              <Kpi k="Count" v={String(stats.count)} />
              <Kpi k="Sum" v={formatGeneral(stats.sum)} />
              <Kpi k="Mean" v={formatGeneral(stats.mean)} />
              <Kpi k="Std dev" v={formatGeneral(stats.stdev)} />
              <Kpi k="Min" v={formatGeneral(stats.min)} />
              <Kpi k="Median" v={formatGeneral(stats.median)} />
              <Kpi k="Max" v={formatGeneral(stats.max)} />
              <Kpi k="P25" v={formatGeneral(stats.p25)} />
              <Kpi k="P75" v={formatGeneral(stats.p75)} />
            </div>
          </div>

          <div className="pane-section">
            <h4>Distribution</h4>
            <Histogram bins={stats.histogram} />
          </div>
        </>
      )}

      {stats.count === 0 && stats.filled === 0 && (
        <div className="empty-note">Select a range containing data to see its profile.</div>
      )}
    </div>
  );
}

function Kpi({ k, v }: { k: string; v: string }) {
  return (
    <div className="kpi">
      <div className="k">{k}</div>
      <div className="v">{v}</div>
    </div>
  );
}

interface Bin {
  from: number;
  to: number;
  count: number;
}

function histogram(values: number[]): Bin[] {
  if (!values.length) return [];
  const min = Math.min(...values);
  const max = Math.max(...values);
  const bins = 12;
  const span = max - min || 1;
  const out: Bin[] = [];
  for (let i = 0; i < bins; i++) {
    out.push({ from: min + (span * i) / bins, to: min + (span * (i + 1)) / bins, count: 0 });
  }
  for (const v of values) {
    const i = Math.min(bins - 1, Math.floor(((v - min) / span) * bins));
    out[i].count++;
  }
  return out;
}

function Histogram({ bins }: { bins: Bin[] }) {
  const max = Math.max(1, ...bins.map((b) => b.count));
  return (
    <svg viewBox="0 0 480 160" style={{ width: '100%', height: 'auto' }}>
      {bins.map((b, i) => {
        const w = 480 / bins.length;
        const h = (b.count / max) * 120;
        return (
          <g key={i}>
            <rect x={i * w + 1} y={140 - h} width={w - 2} height={h} fill="#2f6fed" opacity={0.85} rx={1}>
              <title>{`${formatGeneral(b.from)} – ${formatGeneral(b.to)}: ${b.count}`}</title>
            </rect>
            {i % 3 === 0 && (
              <text x={i * w + w / 2} y={154} textAnchor="middle" fontSize={8} fill="#6b7684">
                {formatGeneral(b.from)}
              </text>
            )}
          </g>
        );
      })}
      <line x1={0} x2={480} y1={140} y2={140} stroke="#b9c0c8" />
    </svg>
  );
}

/* ======================================================================= sql */

export function SqlPane({
  engine,
  sheet,
  onMessage,
}: {
  engine: Engine;
  sheet: Sheet;
  onMessage: (msg: string) => void;
}) {
  const [db, setDb] = useState<SqliteFile | null>(null);
  const [tables, setTables] = useState<TableInfo[]>([]);
  const [sql, setSql] = useState('SELECT name FROM sqlite_master WHERE type = \'table\';');
  const [output, setOutput] = useState<{ columns: string[]; rows: Value[][]; elapsedMs: number } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const openDb = async () => {
    try {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.db,.sqlite,.sqlite3,.db3';
      input.onchange = async () => {
        const file = input.files?.[0];
        if (!file) return;
        setBusy(true);
        try {
          const bytes = new Uint8Array(await file.arrayBuffer());
          const opened = await SqliteFile.open(file.name, bytes);
          setDb(opened);
          setTables(opened.tables());
          setError(null);
          onMessage(`Opened ${file.name}`);
        } catch (e) {
          setError((e as Error).message);
        } finally {
          setBusy(false);
        }
      };
      input.click();
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const runQuery = () => {
    if (!db) {
      setError('Open a .db file first.');
      return;
    }
    setBusy(true);
    const outcomes = db.runScript(sql);
    const failure = outcomes.find((o) => !o.ok);
    if (failure && !failure.ok) {
      setError(failure.error.message);
      setOutput(null);
      setBusy(false);
      return;
    }
    const last = outcomes[outcomes.length - 1];
    if (last?.ok && last.result.columns.length) {
      setOutput(last.result);
      setError(null);
      const target = { r: sheet.rows, c: 0 };
      const range = loadResult(engine, sheet, target, last.result);
      engine.trimExtent(sheet);
      engine.emit();
      onMessage(`Loaded ${last.result.rowCount} rows into ${rangeToA1(range)}`);
    } else {
      setOutput(null);
      setError(null);
      const affected = outcomes.reduce((n, o) => n + (o.ok ? o.result.affected : 0), 0);
      onMessage(`Statement finished${affected ? ` (${affected} rows affected)` : ''}`);
    }
    setBusy(false);
  };

  const loadTable = (name: string) => {
    setSql(`SELECT * FROM "${name}";`);
    setBusy(true);
    setTimeout(() => {
      runQuery();
      setBusy(false);
    }, 0);
  };

  return (
    <div className="pane-body">
      <div className="pane-section">
        <h4>Database</h4>
        <div className="row">
          <button className="btn" onClick={openDb} disabled={busy}>Open .db file…</button>
          {db && <span className="pill ok">{db.name}</span>}
        </div>
        {tables.length > 0 && (
          <table className="data">
            <thead>
              <tr>
                <th>Table</th>
                <th className="num">Columns</th>
                <th className="num">Rows</th>
              </tr>
            </thead>
            <tbody>
              {tables.map((t) => (
                <tr key={t.name} onClick={() => loadTable(t.name)} style={{ cursor: 'pointer' }}>
                  <td>{t.name}</td>
                  <td className="num">{t.columns}</td>
                  <td className="num">{t.rows}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <p className="hint">Databases are read locally as files — nothing leaves this machine.</p>
      </div>

      <div className="pane-section">
        <h4>SQL</h4>
        <textarea className="code" value={sql} onChange={(e) => setSql(e.target.value)} spellCheck={false} />
        <div className="row" style={{ marginTop: 6 }}>
          <button className="btn primary" onClick={runQuery} disabled={!db || busy}>Run query</button>
          {output && (
            <span className="pill">
              {output.rows.length} rows in {output.elapsedMs.toFixed(1)} ms
            </span>
          )}
        </div>
        {error && <div className="pill err" style={{ marginTop: 6 }}>{error}</div>}
      </div>

      {output && (
        <div className="pane-section">
          <h4>Result preview</h4>
          <div style={{ maxHeight: 260, overflow: 'auto' }}>
            <table className="data">
              <thead>
                <tr>{output.columns.map((c) => <th key={c}>{c}</th>)}</tr>
              </thead>
              <tbody>
                {output.rows.slice(0, 100).map((row, i) => (
                  <tr key={i}>
                    {row.map((v, j) => (
                      <td key={j} className={typeof v === 'number' ? 'num' : undefined}>
                        {typeof v === 'string' ? v : v === null ? '' : formatGeneral(v as number)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {output.rows.length > 100 && <p className="hint">Showing the first 100 of {output.rows.length} rows.</p>}
        </div>
      )}
    </div>
  );
}

export { rangeCells, splitRef };
