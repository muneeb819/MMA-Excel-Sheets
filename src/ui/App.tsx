/**
 * Application shell: title bar, ribbon, formula bar, grid, side panes, sheet
 * tabs and status bar.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import Grid, { type Selection } from './Grid';
import { ChartsPane, DashboardPane, DataPane, SqlPane } from './panes';
import { Modal } from './dialogs';
import { Engine, newWorkbook } from '../core/engine';
import {
  MAX_COLS,
  MAX_ROWS,
  type CellFormat,
  type Range,
  type Value,
} from '../core/types';
import { isError } from '../core/types';
import { a1, colName, normRange, rangeToA1 } from '../core/ref';
import { formatGeneral, toText } from '../core/coerce';
import { functionNames } from '../core/formula/functions/index';
import { importCsv, sniffDelimiter, toCsv } from '../io/csv';
import { exportXlsx, importXlsx, sheetToRows } from '../io/xlsx';
import { readFile, readWorkbookFile, writeCsv, writeXlsx } from '../io/fileIO';
import { AutoSaver, listWorkbooks, loadWorkbook, deleteWorkbook, saveWorkbook } from '../store/db';
import { localPayload, toTsv } from '../core/clipboard';

type Pane = 'none' | 'charts' | 'data' | 'dashboard' | 'sql';

const NUMBER_FORMATS: { label: string; code: string }[] = [
  { label: 'General', code: '' },
  { label: 'Number', code: '#,##0.00' },
  { label: 'Integer', code: '#,##0' },
  { label: 'Currency', code: '"$"#,##0.00' },
  { label: 'Percent', code: '0.00%' },
  { label: 'Accounting', code: '#,##0.00_);(#,##0.00)' },
  { label: 'Scientific', code: '0.00E+00' },
  { label: 'Date', code: 'yyyy-mm-dd' },
  { label: 'Date + time', code: 'yyyy-mm-dd hh:mm' },
  { label: 'Time', code: 'hh:mm:ss' },
  { label: 'Text', code: '@' },
];

export default function App() {
  const [engine] = useState(() => new Engine(newWorkbook()));
  const [revision, setRevision] = useState(0);
  const [selection, setSelection] = useState<Selection>({
    cursor: { r: 0, c: 0 },
    range: { r1: 0, c1: 0, r2: 0, c2: 0 },
  });
  /**
   * True only while the formula bar itself has focus. Using a focus flag rather
   * than a sticky "editing" flag means the bar keeps tracking the selected cell
   * after an edit instead of freezing on the last value typed.
   */
  const [formulaFocused, setFormulaFocused] = useState(false);
  const [formulaText, setFormulaText] = useState('');
  const [pane, setPane] = useState<Pane>('none');
  const [zoom, setZoom] = useState(1);
  const [message, setMessage] = useState('Ready');
  const [editSignal, setEditSignal] = useState<{ r: number; c: number; text: string; mode: 'replace' | 'append' } | null>(null);
  const [dialog, setDialog] = useState<null | 'function' | 'goto' | 'about' | 'open'>(null);
  const [dirty, setDirty] = useState(false);

  const sheet = engine.active;
  const selectionRange = normRange(selection.range);

  const bump = useCallback(() => setRevision((n) => n + 1), []);

  useEffect(() => {
    engine.onRevision(() => setRevision((n) => n + 1));
  }, [engine]);

  /* ---------------------------------------------------------- formula bar */

  useEffect(() => {
    if (formulaFocused) return;
    setFormulaText(engine.editText(sheet, selection.cursor.r, selection.cursor.c));
  }, [engine, sheet, selection.cursor.r, selection.cursor.c, revision, formulaFocused]);

  const startFormulaEdit = useCallback(
    (mode: 'replace' | 'append') => {
      setFormulaFocused(true);
      setEditSignal({ r: selection.cursor.r, c: selection.cursor.c, text: formulaText, mode });
    },
    [selection.cursor.r, selection.cursor.c, formulaText],
  );

  const commitFormula = (move: 'down' | 'right' | 'none' = 'none') => {
    setFormulaFocused(false);
    setEditSignal(null);
    engine.transact(() => engine.setInput(sheet, selection.cursor.r, selection.cursor.c, formulaText));
    engine.recalc();
    setDirty(true);
    if (move === 'down') moveCursor(1, 0);
    else if (move === 'right') moveCursor(0, 1);
  };

  /* -------------------------------------------------------------- cursor */

  const moveCursor = (dr: number, dc: number) => {
    const r = Math.max(0, Math.min(MAX_ROWS - 1, selection.cursor.r + dr));
    const c = Math.max(0, Math.min(MAX_COLS - 1, selection.cursor.c + dc));
    setSelection({ cursor: { r, c }, range: { r1: r, c1: c, r2: r, c2: c } });
  };

  const selectAll = () => {
    setSelection({ cursor: selection.cursor, range: { r1: 0, c1: 0, r2: 999, c2: 60 } });
  };

  /* ------------------------------------------------------------- formats */

  const applyFormat = (key: keyof CellFormat, value?: string) => {
    const sel = normRange(selection.range);
    engine.transact(() => {
      for (let r = sel.r1; r <= sel.r2; r++) {
        for (let c = sel.c1; c <= sel.c2; c++) {
          const existing = engine.formatOf(sheet.cells.get(r * MAX_COLS + c)?.styleId) ?? {};
          const next: Record<string, unknown> = { ...existing };
          if (value === undefined) {
            // toggle
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
    setDirty(true);
  };

  /* ------------------------------------------------------------ structure */

  const addSheet = () => {
    const s = engine.addSheet();
    setDirty(true);
    bump();
    setMessage(`Added ${s.name}`);
  };

  const renameSheet = (name: string) => {
    engine.renameSheet(sheet, name);
    setDirty(true);
    bump();
  };

  const deleteSheet = () => {
    const name = sheet.name;
    engine.removeSheet(sheet);
    setDirty(true);
    bump();
    setMessage(`Deleted ${name}`);
  };

  const insertRows = (count: number) => {
    engine.transact(() => engine.insertRows(sheet, selection.cursor.r, count));
    engine.recalc();
    setDirty(true);
  };

  const insertCols = (count: number) => {
    engine.transact(() => engine.insertCols(sheet, selection.cursor.c, count));
    engine.recalc();
    setDirty(true);
  };

  const deleteRows = (count: number) => {
    engine.transact(() => engine.deleteRows(sheet, selection.cursor.r, count));
    engine.recalc();
    setDirty(true);
  };

  const deleteCols = (count: number) => {
    engine.transact(() => engine.deleteCols(sheet, selection.cursor.c, count));
    engine.recalc();
    setDirty(true);
  };

  /* ----------------------------------------------------------------- files */

  const doImportCsv = async () => {
    const file = await readFile(
      [{ name: 'CSV', extensions: ['csv', 'tsv', 'txt'] }],
      '.csv,.tsv,.txt',
    );
    if (!file) return;
    const text = new TextDecoder().decode(file.bytes);
    const result = importCsv(text, { delimiter: sniffDelimiter(text) });
    engine.transact(() => {
      result.rows.forEach((row, r) => {
        row.forEach((v, c) => engine.setValue(sheet, r, c, v));
      });
      if (result.dateColumns.length) {
        const styleId = engine.styleIdFor({ numFmt: 'yyyy-mm-dd' });
        for (const col of result.dateColumns) {
          for (let r = 0; r < result.rowCount; r++) engine.setStyleId(sheet, r, col, styleId);
        }
      }
    });
    engine.recalc();
    engine.trimAll();
    engine.emit();
    setDirty(true);
    setMessage(`Imported ${result.rowCount} rows × ${result.colCount} columns from ${file.name}`);
  };

  const doImportXlsx = async () => {
    const file = await readWorkbookFile();
    if (!file) return;
    if (/\.(csv|tsv|txt)$/i.test(file.name)) {
      const text = new TextDecoder().decode(file.bytes);
      const result = importCsv(text, { delimiter: sniffDelimiter(text) });
      engine.transact(() => {
        result.rows.forEach((row, r) => row.forEach((v, c) => engine.setValue(sheet, r, c, v)));
      });
      engine.recalc();
      engine.trimAll();
      engine.emit();
      setDirty(true);
      setMessage(`Imported ${result.rowCount} rows from ${file.name}`);
      return;
    }
    // `bytes` is a view over a buffer that can be larger; copy exactly.
    const copy = file.bytes.slice();
    const result = importXlsx(engine, copy.buffer.slice(copy.byteOffset, copy.byteOffset + copy.byteLength) as ArrayBuffer);
    setDirty(true);
    bump();
    setMessage(`Opened ${file.name} — ${result.cellCount} cells across ${result.sheetNames.length} sheet(s)`);
  };

  const doExportXlsx = async () => {
    const buffer = exportXlsx(engine);
    const path = await writeXlsx(engine.wb.name || 'Workbook', new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    if (path) setMessage(`Saved ${path}`);
  };

  const doExportCsv = async () => {
    const rows = sheetToRows(engine, sheet);
    const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
    const padded = rows.map((r) => r.concat(Array(Math.max(0, width - r.length)).fill('')));
    const path = await writeCsv(sheet.name, toCsv(padded));
    if (path) setMessage(`Saved ${path}`);
  };

  const saveLocal = async () => {
    try {
      await saveWorkbook(engine.wb);
      setDirty(false);
      setMessage('Saved to this computer');
    } catch (e) {
      setMessage(`Save failed: ${(e as Error).message}`);
    }
  };

  /* --------------------------------------------------------------- status */

  const stats = useMemo(() => {
    const sel = normRange(selection.range);
    let count = 0;
    let numeric = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    let errors = 0;
    let filled = 0;
    const cells = (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1);
    if (cells <= 500_000) {
      for (let r = sel.r1; r <= sel.r2; r++) {
        for (let c = sel.c1; c <= sel.c2; c++) {
          count++;
          const v = engine.valueAt(sheet, r, c);
          if (v === null || v === '') continue;
          filled++;
          if (isError(v)) {
            errors++;
            continue;
          }
          if (typeof v === 'number') {
            numeric++;
            sum += v;
            if (v < min) min = v;
            if (v > max) max = v;
          }
        }
      }
    }
    return { count, filled, numeric, sum, min: numeric ? min : 0, max: numeric ? max : 0, errors, cells };
  }, [engine, sheet, selection.range, revision]);

  const activeValue = engine.valueAt(sheet, selection.cursor.r, selection.cursor.c);

  /* ------------------------------------------------------------- keyboard */

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT')) {
        return;
      }
      const ctrl = e.ctrlKey || e.metaKey;
      if (ctrl && e.key.toLowerCase() === 's') {
        e.preventDefault();
        void saveLocal();
      } else if (ctrl && e.key.toLowerCase() === 'o') {
        e.preventDefault();
        void doImportXlsx();
      } else if (ctrl && e.key.toLowerCase() === 'g') {
        e.preventDefault();
        setDialog('goto');
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  /* Desktop menu commands, when running inside the Electron shell. */
  useEffect(() => {
    const desktop = typeof window !== 'undefined' ? window.sheetcraftDesktop : undefined;
    if (!desktop) return;
    return desktop.onMenu((action) => {
      if (action === 'save') void saveLocal();
      else if (action === 'open') void doImportXlsx();
      else if (action === 'export-csv') void doExportCsv();
      else if (action === 'export-xlsx') void doExportXlsx();
      else if (action.startsWith('pane:')) setPane(action.slice(5) as Pane);
    });
  });

  /* ------------------------------------------------------------ rendering */

  return (
    <div className="app">
      <div className="titlebar">
        <span className="brand">SheetCraft</span>
        <span className="doc-name">
          {engine.wb.name}
          {dirty ? ' •' : ''}
        </span>
        <span className="spacer" />
        <button className="tool" onClick={() => void saveLocal()} title="Save to this computer (Ctrl+S)">
          Save
        </button>
        <button className="tool" onClick={() => setDialog('open')}>Open…</button>
        <button className="tool" onClick={() => setDialog('about')} title="About">?</button>
      </div>

      {/* ------------------------------------------------------------ ribbon */}
      <div className="ribbon">
        <div className="ribbon-group">
          <div className="ribbon-row">
            <button className="tool" onClick={() => void doImportXlsx()}>Open</button>
            <button className="tool" onClick={() => void doImportCsv()}>Import CSV</button>
          </div>
          <div className="ribbon-row">
            <button className="tool" onClick={() => void doExportXlsx()}>Save .xlsx</button>
            <button className="tool" onClick={() => void doExportCsv()}>Save .csv</button>
          </div>
          <span className="label">File</span>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-row">
            <button className="tool icon" title="Undo (Ctrl+Z)" disabled={!engine.canUndo()} onClick={() => { if (engine.undo()) { engine.recalc(); setDirty(true); } }}>↶</button>
            <button className="tool icon" title="Redo (Ctrl+Y)" disabled={!engine.canRedo()} onClick={() => { if (engine.redo()) { engine.recalc(); setDirty(true); } }}>↷</button>
          </div>
          <span className="label">History</span>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-row">
            <button className="tool" onClick={() => setDialog('function')}>fx Insert function</button>
          </div>
          <span className="label">Formulas</span>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-row">
            <button className={`tool icon${isActiveFmt('bold') ? ' active' : ''}`} title="Bold (Ctrl+B)" onClick={() => applyFormat('bold')}>B</button>
            <button className={`tool icon${isActiveFmt('italic') ? ' active' : ''}`} title="Italic (Ctrl+I)" onClick={() => applyFormat('italic')}><i>I</i></button>
            <button className={`tool icon${isActiveFmt('underline') ? ' active' : ''}`} title="Underline (Ctrl+U)" onClick={() => applyFormat('underline')}>U</button>
            <span className="tool-sep" />
            <button className={`tool icon${currentAlign() === 'left' ? ' active' : ''}`} title="Align left" onClick={() => applyFormat('align', 'left')}>⇤</button>
            <button className={`tool icon${currentAlign() === 'center' ? ' active' : ''}`} title="Centre" onClick={() => applyFormat('align', 'center')}>⇹</button>
            <button className={`tool icon${currentAlign() === 'right' ? ' active' : ''}`} title="Align right" onClick={() => applyFormat('align', 'right')}>⇥</button>
          </div>
          <div className="ribbon-row">
            <select
              value={currentNumFmt()}
              onChange={(e) => applyFormat('numFmt', e.target.value)}
              style={{ width: 140, fontSize: 12 }}
            >
              {NUMBER_FORMATS.map((f) => (
                <option key={f.label} value={f.code}>{f.label}</option>
              ))}
            </select>
            <ColorButton onPick={(c) => applyFormat('bg', c)} title="Fill colour" />
            <ColorButton onPick={(c) => applyFormat('color', c)} title="Text colour" />
          </div>
          <span className="label">Format</span>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-row">
            <button className="tool" onClick={() => insertRows(1)}>Row +</button>
            <button className="tool" onClick={() => insertCols(1)}>Col +</button>
          </div>
          <div className="ribbon-row">
            <button className="tool" onClick={() => deleteRows(1)}>Row −</button>
            <button className="tool" onClick={() => deleteCols(1)}>Col −</button>
          </div>
          <span className="label">Cells</span>
        </div>

        <div className="ribbon-group">
          <div className="ribbon-row">
            <button className={`tool${pane === 'dashboard' ? ' active' : ''}`} onClick={() => setPane(pane === 'dashboard' ? 'none' : 'dashboard')}>Dashboard</button>
            <button className={`tool${pane === 'charts' ? ' active' : ''}`} onClick={() => setPane(pane === 'charts' ? 'none' : 'charts')}>Charts</button>
          </div>
          <div className="ribbon-row">
            <button className={`tool${pane === 'data' ? ' active' : ''}`} onClick={() => setPane(pane === 'data' ? 'none' : 'data')}>Clean data</button>
            <button className={`tool${pane === 'sql' ? ' active' : ''}`} onClick={() => setPane(pane === 'sql' ? 'none' : 'sql')}>SQL</button>
          </div>
          <span className="label">Analyse</span>
        </div>

        <div className="ribbon-group zoom">
          <span className="label" style={{ margin: 0 }}>Zoom</span>
          <button className="tool icon" onClick={() => setZoom((z) => Math.max(0.5, Math.round((z - 0.1) * 10) / 10))}>−</button>
          <span style={{ minWidth: 38, textAlign: 'center' }}>{Math.round(zoom * 100)}%</span>
          <button className="tool icon" onClick={() => setZoom((z) => Math.min(2, Math.round((z + 0.1) * 10) / 10))}>+</button>
        </div>
      </div>

      {/* ------------------------------------------------------- formula bar */}
      <div className="formula-bar">
        <div className="name-box">{a1(selection.cursor.r, selection.cursor.c)}</div>
        <input
          className="formula-input"
          value={formulaText}
          spellCheck={false}
          onChange={(e) => {
            setFormulaText(e.target.value);
            setFormulaFocused(true);
          }}
          onFocus={() => setFormulaFocused(true)}
          onBlur={() => {
            setFormulaFocused(false);
            commitFormula('none');
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              commitFormula(e.shiftKey ? 'none' : 'down');
            } else if (e.key === 'Escape') {
              e.preventDefault();
              setFormulaFocused(false);
              setEditSignal(null);
              setFormulaText(engine.editText(sheet, selection.cursor.r, selection.cursor.c));
            }
            e.stopPropagation();
          }}
        />
        <div className="fx">fx</div>
      </div>

      {/* ---------------------------------------------------------- workspace */}
      <div className="workspace">
        <Grid
          engine={engine}
          sheet={sheet}
          selection={selection}
          onSelectionChange={setSelection}
          revision={revision}
          zoom={zoom}
          editSignal={editSignal}
          onEditSignalHandled={() => setEditSignal(null)}
        />

        {pane !== 'none' && (
          <div className={`side-pane${pane === 'sql' ? ' wide' : ''}`}>
            <div className="pane-header">
              <span>
                {pane === 'charts' && 'Charts'}
                {pane === 'data' && 'Clean data'}
                {pane === 'dashboard' && 'Dashboard'}
                {pane === 'sql' && 'SQL'}
              </span>
              <span className="spacer" />
              <button className="tool icon" onClick={() => setPane('none')} title="Close">✕</button>
            </div>
            {pane === 'charts' && <ChartsPane engine={engine} sheet={sheet} selection={selectionRange} />}
            {pane === 'data' && (
              <DataPane engine={engine} sheet={sheet} selection={selectionRange} onMessage={setMessage} />
            )}
            {pane === 'dashboard' && (
              <DashboardPane engine={engine} sheet={sheet} selection={selectionRange} />
            )}
            {pane === 'sql' && <SqlPane engine={engine} sheet={sheet} onMessage={setMessage} />}
          </div>
        )}
      </div>

      {/* ------------------------------------------------------------- tabs */}
      <div className="tabbar">
        <div className="sheet-tabs">
          {engine.wb.sheets.map((s, i) => (
            <div
              key={s.id}
              className={`sheet-tab${i === engine.wb.activeSheet ? ' active' : ''}`}
              onClick={() => {
                engine.wb.activeSheet = i;
                setSelection({ cursor: { r: 0, c: 0 }, range: { r1: 0, c1: 0, r2: 0, c2: 0 } });
                bump();
              }}
              onDoubleClick={() => {
                const name = prompt('Sheet name', s.name);
                if (name) renameSheet(name);
              }}
              title={`${s.name} — double-click to rename`}
            >
              {s.name}
              {engine.wb.sheets.length > 1 && (
                <button
                  className="close"
                  onClick={(e) => {
                    e.stopPropagation();
                    engine.wb.activeSheet = i;
                    deleteSheet();
                  }}
                >
                  ✕
                </button>
              )}
            </div>
          ))}
          <button className="tool" onClick={addSheet} title="Add sheet">+</button>
        </div>
      </div>

      {/* ------------------------------------------------------- status bar */}
      <div className="statusbar">
        <span>{rangeToA1(selectionRange)}</span>
        <span>
          {stats.count} cell{stats.count === 1 ? '' : 's'} selected
        </span>
        {stats.numeric > 0 && (
          <>
            <span>Sum: {formatGeneral(stats.sum)}</span>
            <span>Avg: {formatGeneral(stats.sum / stats.numeric)}</span>
            <span>Min: {formatGeneral(stats.min)}</span>
            <span>Max: {formatGeneral(stats.max)}</span>
          </>
        )}
        {stats.errors > 0 && <span className="badge error">{stats.errors} error(s)</span>}
        <span className="spacer" />
        <span>{message}</span>
      </div>

      {dialog === 'function' && (
        <Modal title="Insert function" onClose={() => setDialog(null)}>
          <FunctionDialog
            onPick={(text) => {
              setFormulaText((prev) => (prev && !prev.endsWith('(') ? `${prev}${text}` : text));
              setFormulaFocused(true);
              setDialog(null);
            }}
          />
        </Modal>
      )}
      {dialog === 'goto' && (
        <Modal title="Go to" onClose={() => setDialog(null)}>
          <GotoDialog onGo={(r, c) => setSelection({ cursor: { r, c }, range: { r1: r, c1: c, r2: r, c2: c } })} />
        </Modal>
      )}
      {dialog === 'open' && (
        <Modal title="Open a saved workbook" onClose={() => setDialog(null)}>
          <OpenDialog
            engine={engine}
            onOpened={() => {
              setDirty(false);
              bump();
              setDialog(null);
            }}
          />
        </Modal>
      )}
      {dialog === 'about' && (
        <Modal title="About SheetCraft" onClose={() => setDialog(null)}>
          <div style={{ lineHeight: 1.7 }}>
            <p>
              <strong>SheetCraft</strong> is an offline spreadsheet suite: a formula engine with
              Excel-compatible functions, charts, data cleaning, dashboards and local SQLite.
            </p>
            <p className="hint">
              Everything runs on this machine. Workbooks are stored locally and can be exported to
              .xlsx or .csv at any time.
            </p>
            <p className="hint">Registered functions: {functionNames().length}</p>
          </div>
        </Modal>
      )}
    </div>
  );

  /* ------------------------------------------------------------- helpers */

  function isActiveFmt(key: keyof CellFormat): boolean {
    const fmt = engine.formatOf(sheet.cells.get(selection.cursor.r * MAX_COLS + selection.cursor.c)?.styleId);
    return Boolean(fmt?.[key]);
  }

  function currentAlign(): string {
    const fmt = engine.formatOf(sheet.cells.get(selection.cursor.r * MAX_COLS + selection.cursor.c)?.styleId);
    return fmt?.align ?? (typeof activeValue === 'number' ? 'right' : 'left');
  }

  function currentNumFmt(): string {
    const fmt = engine.formatOf(sheet.cells.get(selection.cursor.r * MAX_COLS + selection.cursor.c)?.styleId);
    return fmt?.numFmt ?? '';
  }
}

/* ------------------------------------------------------------ small parts */

function ColorButton({ onPick, title }: { onPick: (hex: string) => void; title: string }) {
  const ref = useRef<HTMLInputElement | null>(null);
  return (
    <>
      <button className="tool icon" title={title} onClick={() => ref.current?.click()}>🎨</button>
      <input
        ref={ref}
        type="color"
        style={{ width: 0, height: 0, border: 'none', padding: 0, opacity: 0, position: 'absolute' }}
        onChange={(e) => onPick(e.target.value)}
      />
    </>
  );
}

function FunctionDialog({ onPick }: { onPick: (text: string) => void }) {
  const [filter, setFilter] = useState('');
  const names = useMemo(() => {
    const q = filter.trim().toUpperCase();
    return functionNames().filter((n) => !q || n.includes(q)).slice(0, 400);
  }, [filter]);

  return (
    <div>
      <div className="row">
        <input
          type="text"
          autoFocus
          placeholder="Search functions"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
          style={{ width: '100%' }}
        />
      </div>
      <div className="fn-list">
        {names.map((n) => (
          <div key={n} onClick={() => onPick(`${n}(`)}>
            {n}
          </div>
        ))}
      </div>
    </div>
  );
}

function GotoDialog({ onGo }: { onGo: (r: number, c: number) => void }) {
  const [ref, setRef] = useState('A1');
  return (
    <form
      onSubmit={(e) => {
        e.preventDefault();
        const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(ref.trim());
        if (!m) return;
        let c = 0;
        for (let i = 0; i < m[1].length; i++) c = c * 26 + (m[1].toUpperCase().charCodeAt(i) - 64);
        onGo(Math.max(0, parseInt(m[2], 10) - 1), Math.max(0, c - 1));
      }}
    >
      <div className="row">
        <label>Reference</label>
        <input type="text" value={ref} onChange={(e) => setRef(e.target.value)} autoFocus />
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 10 }}>
        <button className="btn" type="button" onClick={() => history.back()}>Cancel</button>
        <button className="btn primary" type="submit">Go</button>
      </div>
    </form>
  );
}

function OpenDialog({ engine, onOpened }: { engine: Engine; onOpened: () => void }) {
  const [books, setBooks] = useState<{ id: string; name: string; savedAt: number; cellCount: number }[]>([]);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(() => {
    listWorkbooks()
      .then(setBooks)
      .catch((e) => setError((e as Error).message));
  }, []);

  useEffect(refresh, [refresh]);

  return (
    <div>
      {error && <div className="pill err">{error}</div>}
      {books.length === 0 ? (
        <div className="empty-note">No workbooks saved on this computer yet.</div>
      ) : (
        <table className="data">
          <thead>
            <tr>
              <th>Name</th>
              <th className="num">Cells</th>
              <th>Saved</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {books.map((b) => (
              <tr key={b.id}>
                <td>{b.name}</td>
                <td className="num">{b.cellCount}</td>
                <td>{new Date(b.savedAt).toLocaleString()}</td>
                <td>
                  <button
                    className="btn"
                    onClick={async () => {
                      const wb = await loadWorkbook(b.id);
                      if (!wb) return;
                      engine.wb = wb;
                      engine.invalidateAll();
                      engine.recalc();
                      onOpened();
                    }}
                  >
                    Open
                  </button>{' '}
                  <button
                    className="btn danger"
                    onClick={async () => {
                      await deleteWorkbook(b.id);
                      refresh();
                    }}
                  >
                    Delete
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 10 }}>
        <button className="btn" onClick={refresh}>Refresh</button>
      </div>
    </div>
  );
}

/* --------------------------------------------------------------- helpers */

export { AutoSaver, localPayload, toTsv, colName, toText };
