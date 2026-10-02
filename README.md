# SheetCraft

An offline spreadsheet suite that behaves like Microsoft Excel. It runs entirely on
your own machine — no account, no server, no internet required once installed.

```
sheetcraft/
  src/core/        formula engine, recalculation, coercion, data tools, SQLite
  src/ui/          grid, ribbon, panes, charts
  src/io/          CSV, XLSX, file access
  electron/        desktop shell (main + preload)
  scripts/         self-test and build helpers
```

## Install

### Easiest: run the installer

Double-click **`release/SheetCraft-Setup-1.0.0.exe`** (or the portable
`SheetCraft-Portable-1.0.0.exe`, which needs no installation at all).

### Build from source

```powershell
cd sheetcraft
.\setup.ps1          # installs dependencies and builds
.\setup.ps1 -Package # also produces the Windows installer
```

Then launch it:

```powershell
npm run desktop      # desktop app
npm run dev          # browser dev server (http://localhost:5173)
```

> On PowerShell you may need `npm.cmd` instead of `npm` if script execution is
> disabled — `setup.ps1` handles that automatically.

## What it does

**Spreadsheet.** Virtualised canvas grid, multi-sheet tabs, copy/cut/paste (with a
TSV fallback that talks to Excel), fill handle, frozen headers, insert/delete rows
and columns, column resizing, cell comments, and full undo/redo.

**Formulas.** ~330 Excel functions across maths, statistics, text, dates, lookup,
financial, logical and information families, plus dynamic arrays.

| Family | Examples |
| --- | --- |
| Maths | `SUM`, `SUMIFS`, `COUNTIF`, `ROUND`, `MOD`, `MMULT`, `MINVERSE`, `SUBTOTAL` |
| Statistics | `AVERAGE`, `STDEV.S`, `MEDIAN`, `CORREL`, `FORECAST`, `NORM.DIST`, `T.TEST` |
| Text | `TEXT`, `LEFT`, `MID`, `SUBSTITUTE`, `TEXTJOIN`, `PROPER`, `REGEXEXTRACT` |
| Dates | `DATE`, `EDATE`, `EOMONTH`, `DATEDIF`, `WORKDAY`, `YEARFRAC`, `WEEKNUM` |
| Lookup | `VLOOKUP`, `XLOOKUP`, `MATCH`, `INDEX`, `OFFSET`, `INDIRECT` |
| Financial | `PMT`, `FV`, `PV`, `RATE`, `NPER`, `IRR`, `NPV`, `XIRR`, `CUMIPMT` |
| Logical | `IF`, `IFS`, `SWITCH`, `IFERROR`, `XOR` (all lazy, like Excel) |
| Dynamic arrays | `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `SORTBY`, `TAKE` |

Formulas support cross-sheet references, defined names, `A1`/`$A$1` anchoring, and
spill ranges. Circular references report `#CIRCULAR!` instead of hanging.

**Charts.** Column, bar, line, area, pie, doughnut, scatter and radar, drawn as SVG.
Charts hold range references, so editing the source data redraws them.

**Clean data.** Trim, case, type conversion, duplicate/empty-row removal, column
splitting, zero padding and a column profile that flags mixed types, blanks and
duplicates.

**Dashboard.** Selection KPIs, numeric summary, and a distribution histogram.

**SQL.** Open any local `.db` file and run SQL against it via WebAssembly. Results
load straight into a sheet. Nothing is uploaded anywhere.

**Files.** Read and write real `.xlsx` (formulas, fonts, fills, alignment and
number formats survive the round trip) and `.csv`.

## Keyboard

| Key | Action |
| --- | --- |
| Arrows / Shift+Arrows | Move / extend selection |
| Enter, Tab | Commit and move |
| F2 or double-click | Edit in place |
| Ctrl+C / X / V | Copy, cut, paste |
| Ctrl+Z / Ctrl+Y | Undo, redo |
| Ctrl+B / I / U | Bold, italic, underline |
| Ctrl+A | Select all |
| Delete | Clear cells |
| Ctrl+S | Save to this computer |
| Ctrl+O | Open a workbook |

## Testing

```powershell
npm run typecheck   # strict TypeScript
npm run selftest    # 141 engine assertions against known Excel results
```

The self-test checks the formula engine against documented Excel values —
`PMT`, `FV`, `IRR`, `SLOPE`, `STDEV.S`, `EDATE` rollover, number formatting,
reference translation on copy, spill behaviour and circular-reference detection.

## Notes and limits

- **Autosave is not enabled.** Use **Save** in the title bar (Ctrl+S) to keep a
  workbook in the app's local store, and **Save .xlsx** to write a file.
- **ROW()/COLUMN() without an argument** report `#VALUE!`; the evaluator has no
  notion of "current cell" outside the formula bar.
- **INDIRECT** resolves references on the current sheet only.
- **INDIRECT with R1C1-style text** (`INDIRECT(x, FALSE)`) is not implemented.
- `LET` substitutes bound names only for a top-level bare-name body; nested use
  falls back to normal name resolution.
- **Radar charts** render with the line renderer rather than a true radar grid.
- Undo uses whole-sheet snapshots (capped at 120 steps), which is simple and
  reliable but heavier than a command log on very large sheets.

## Privacy

No telemetry, no network calls, no cloud. Workbooks live in the browser's
IndexedDB, databases are read from disk, and exports are written where you ask.
