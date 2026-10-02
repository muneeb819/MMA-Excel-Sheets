/**
 * Local SQLite connectivity.
 *
 * Databases are ordinary `.db` files on disk, opened in the browser through
 * sql.js (WebAssembly), so queries run entirely offline — no server, no driver
 * install. Results are loaded straight into a sheet.
 */

import initSqlJs, { type Database, type SqlJsStatic, type SqlValue } from 'sql.js';
import type { Range, Sheet, Value } from './types';
import { parseInput } from './coerce';
import type { Engine } from './engine';

export interface TableInfo {
  name: string;
  columns: number;
  rows: number;
}

export interface QueryResult {
  columns: string[];
  rows: Value[][];
  rowCount: number;
  affected: number;
  elapsedMs: number;
}

export interface QueryError {
  message: string;
  query: string;
}

export type SqlOutcome =
  | { ok: true; result: QueryResult }
  | { ok: false; error: QueryError };

let sqlPromise: Promise<SqlJsStatic> | null = null;

/** Load the WebAssembly module once and reuse it. */
function engine(): Promise<SqlJsStatic> {
  if (!sqlPromise) {
    sqlPromise = initSqlJs({
      // In the desktop build the wasm sits next to the app; in dev it is served
      // from node_modules. Try both before giving up.
      locateFile: (file: string) => {
        const packaged = `sql-wasm.wasm`;
        if (typeof window !== 'undefined' && (window as { __SHEETCRAFT_PACKAGED__?: boolean }).__SHEETCRAFT_PACKAGED__) {
          return packaged;
        }
        return `node_modules/sql.js/dist/${file}`;
      },
    });
  }
  return sqlPromise;
}

export class SqliteFile {
  private db: Database | null = null;
  readonly name: string;

  private constructor(name: string, db: Database) {
    this.name = name;
    this.db = db;
  }

  /** Open a `.db` file from raw bytes. */
  static async open(name: string, bytes: Uint8Array): Promise<SqliteFile> {
    const SQL = await engine();
    return new SqliteFile(name, new SQL.Database(bytes));
  }

  /** Create an empty in-memory database (used for scratch queries). */
  static async createScratch(): Promise<SqliteFile> {
    const SQL = await engine();
    return new SqliteFile(':memory:', new SQL.Database());
  }

  get isOpen(): boolean {
    return this.db !== null;
  }

  /** Tables in the database with their row counts. */
  tables(): TableInfo[] {
    if (!this.db) return [];
    const stmt = this.db.prepare(
      "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name",
    );
    const out: TableInfo[] = [];
    while (stmt.step()) {
      const row = stmt.getAsObject() as Record<string, SqlValue>;
      const tableName = String(row.name);
      let count = 0;
      try {
        const countStmt = this.db.prepare(`SELECT COUNT(*) AS n FROM "${tableName.replace(/"/g, '""')}"`);
        if (countStmt.step()) count = Number((countStmt.getAsObject() as { n: SqlValue }).n);
        countStmt.free();
      } catch {
        count = 0;
      }
      const cols = this.columns(tableName);
      out.push({ name: tableName, columns: cols.length, rows: count });
    }
    stmt.free();
    return out;
  }

  /** Column names of a table, empty when it does not exist. */
  columns(table: string): string[] {
    if (!this.db) return [];
    try {
      const stmt = this.db.prepare(`PRAGMA table_info("${table.replace(/"/g, '""')}")`);
      const out: string[] = [];
      while (stmt.step()) {
        out.push(String((stmt.getAsObject() as { name: SqlValue }).name));
      }
      stmt.free();
      return out;
    } catch {
      return [];
    }
  }

  /** Run a statement or query. Never throws; failures come back as `ok: false`. */
  run(sql: string, params: SqlValue[] = []): SqlOutcome {
    if (!this.db) {
      return { ok: false, error: { message: 'Database is not open', query: sql } };
    }
    const started = performance.now();
    let stmt: ReturnType<Database['prepare']>;
    try {
      stmt = this.db.prepare(sql);
    } catch (e) {
      return { ok: false, error: { message: (e as Error).message, query: sql } };
    }

    try {
      stmt.bind(params);
      const columns: string[] = stmt.getColumnNames();
      const rows: Value[][] = [];
      while (stmt.step()) {
        rows.push(stmt.get() as Value[]);
      }
      const elapsedMs = performance.now() - started;

      // DDL/DML statements have no result set; report how many rows changed.
      if (!columns.length) {
        let affected = 0;
        try {
          const changes = this.db.exec('SELECT changes()');
          affected = Number(changes[0]?.values[0]?.[0] ?? 0);
        } catch {
          affected = 0;
        }
        return { ok: true, result: { columns: [], rows: [], rowCount: 0, affected, elapsedMs } };
      }
      return { ok: true, result: { columns, rows, rowCount: rows.length, affected: 0, elapsedMs } };
    } catch (e) {
      return { ok: false, error: { message: (e as Error).message, query: sql } };
    } finally {
      stmt.free();
    }
  }

  /** Run several statements separated by `;`. */
  runScript(sql: string): SqlOutcome[] {
    return splitStatements(sql).map((s) => this.run(s));
  }

  /** Serialise back to bytes so the database can be saved. */
  export(): Uint8Array | null {
    return this.db ? this.db.export() : null;
  }

  close(): void {
    this.db?.close();
    this.db = null;
  }
}

/** Split a script into statements, ignoring semicolons inside strings. */
export function splitStatements(sql: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inString = false;
  let quote = '';

  for (let i = 0; i < sql.length; i++) {
    const ch = sql[i];
    if (inString) {
      cur += ch;
      if (ch === quote) {
        if (sql[i + 1] === quote) {
          cur += quote;
          i++;
        } else inString = false;
      }
      continue;
    }
    if (ch === "'" || ch === '"' || ch === '`') {
      inString = true;
      quote = ch;
      cur += ch;
      continue;
    }
    if (ch === ';') {
      if (cur.trim()) out.push(cur.trim());
      cur = '';
      continue;
    }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

/**
 * Write query results into a sheet starting at `target`.
 * Values are coerced with the same rules as typed input, so `007` stays text.
 */
export function loadResult(
  engine: Engine,
  sheet: Sheet,
  target: { r: number; c: number },
  result: QueryResult,
  includeHeader = true,
): Range {
  let maxRow = 0;
  let maxCol = 0;

  if (includeHeader && result.columns.length) {
    for (let c = 0; c < result.columns.length; c++) {
      engine.setValue(sheet, target.r, target.c + c, result.columns[c]);
    }
    maxRow = 1;
    maxCol = result.columns.length;
  }

  for (let r = 0; r < result.rows.length; r++) {
    const row = result.rows[r];
    for (let c = 0; c < row.length; c++) {
      engine.setValue(sheet, target.r + (includeHeader ? 1 : 0) + r, target.c + c, coerce(row[c]));
    }
    maxRow = Math.max(maxRow, result.rows.length ? (includeHeader ? 1 : 0) + r + 1 : maxRow);
    maxCol = Math.max(maxCol, row.length);
  }

  engine.recalc();
  return {
    r1: target.r,
    c1: target.c,
    r2: target.r + Math.max(0, maxRow - 1),
    c2: target.c + Math.max(0, maxCol - 1),
  };
}

function coerce(v: unknown): Value {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Uint8Array) return `<${v.length} bytes>`;
  return parseInput(String(v));
}
