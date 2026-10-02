/**
 * File access that prefers the desktop shell.
 *
 * In the packaged app this goes through native dialogs; in a browser it falls
 * back to a hidden file input and a download link, so the same code paths serve
 * both.
 */

const XLSX_FILTERS = [
  { name: 'Excel workbook', extensions: ['xlsx', 'xlsm'] },
  { name: 'All files', extensions: ['*'] },
];
const CSV_FILTERS = [
  { name: 'Comma separated values', extensions: ['csv', 'tsv', 'txt'] },
  { name: 'All files', extensions: ['*'] },
];

function bridge() {
  return typeof window !== 'undefined' ? window.sheetcraftDesktop : undefined;
}

export const isDesktop = (): boolean => Boolean(bridge()?.isDesktop);

/** Prompt for a file and return its bytes, or null if cancelled. */
export async function readFile(
  filters: { name: string; extensions: string[] }[],
  accept: string,
): Promise<{ name: string; bytes: Uint8Array } | null> {
  const desktop = bridge();
  if (desktop) {
    const result = await desktop.openFile(filters);
    return result ? { name: result.name, bytes: result.bytes } : null;
  }
  return readViaInput(accept);
}

/** Read the workbook file (xlsx or csv) the user chose. */
export async function readWorkbookFile(): Promise<{ name: string; bytes: Uint8Array } | null> {
  return readFile([...XLSX_FILTERS.slice(0, 1), ...CSV_FILTERS.slice(0, 1)], '.xlsx,.xlsm,.csv,.tsv,.txt');
}

/** Read a SQLite database file. */
export async function readDatabaseFile(): Promise<{ name: string; bytes: Uint8Array } | null> {
  return readFile(
    [{ name: 'SQLite database', extensions: ['db', 'sqlite', 'sqlite3', 'db3'] }],
    '.db,.sqlite,.sqlite3,.db3',
  );
}

/** Write bytes out, asking the user where. Returns the path/name when saved. */
export async function writeFile(
  suggestedName: string,
  bytes: Blob | Uint8Array,
  filters: { name: string; extensions: string[] }[],
): Promise<string | null> {
  const desktop = bridge();
  if (desktop) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(await bytes.arrayBuffer());
    return desktop.saveFile(suggestedName, data, filters);
  }
  downloadBlob(bytes instanceof Blob ? bytes : new Blob([bytes.slice().buffer as ArrayBuffer]), suggestedName);
  return suggestedName;
}

export async function writeXlsx(suggestedName: string, blob: Blob): Promise<string | null> {
  return writeFile(suggestedName.endsWith('.xlsx') ? suggestedName : `${suggestedName}.xlsx`, blob, XLSX_FILTERS);
}

export async function writeCsv(suggestedName: string, text: string): Promise<string | null> {
  const name = suggestedName.endsWith('.csv') ? suggestedName : `${suggestedName}.csv`;
  return writeFile(name, new Blob([text], { type: 'text/csv;charset=utf-8' }), CSV_FILTERS);
}

/* ------------------------------------------------------------- fallbacks */

function readViaInput(accept: string): Promise<{ name: string; bytes: Uint8Array } | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) {
        resolve(null);
        return;
      }
      resolve({ name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) });
    };
    // A cancelled picker fires no event in some browsers; resolve on blur.
    window.addEventListener('focus', () => setTimeout(() => resolve(null), 800), { once: true });
    input.click();
  });
}

function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
