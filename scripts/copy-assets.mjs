/**
 * Copy runtime assets into dist/ so the packaged app works from file://.
 *
 * sql.js needs its wasm next to index.html; without this the SQLite pane fails
 * in the desktop build.
 *
 * IMPORTANT: run this AFTER `vite build`. Vite empties outDir by default, so a
 * copy performed first is silently deleted again.
 */

import { copyFileSync, mkdirSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');

const copies = [
  {
    from: join(root, 'node_modules', 'sql.js', 'dist', 'sql-wasm.wasm'),
    to: join(root, 'dist', 'sql-wasm.wasm'),
  },
];

let copied = 0;
for (const { from, to } of copies) {
  if (!existsSync(from)) {
    console.warn(`[assets] missing, skipped: ${from}`);
    continue;
  }
  mkdirSync(dirname(to), { recursive: true });
  copyFileSync(from, to);
  copied++;
  console.log(`[assets] ${from.split(/[\\/]/).pop()} -> dist/`);
}

console.log(`[assets] ${copied} file(s) copied`);
