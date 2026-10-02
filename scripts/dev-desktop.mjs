/**
 * Development launcher: Vite dev server + Electron with DevTools.
 *
 * Keeps the renderer in hot-reload mode while still running inside the real
 * desktop shell, so what you test is the app you ship (native file dialogs,
 * application menu, SQLite) rather than a browser-only approximation.
 *
 *   npm run desktop:dev
 */

import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';

const require = createRequire(import.meta.url);
const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const PORT = Number(process.env.SHEETCRAFT_DEV_PORT ?? 5173);
const URL = `http://localhost:${PORT}`;

// Vite's package "exports" map does not expose its bin entry, so resolve the
// file directly rather than through require.resolve.
const VITE_BIN = join(ROOT, 'node_modules', 'vite', 'bin', 'vite.js');

const children = [];

/** Kill every child on exit so Ctrl+C leaves nothing behind. */
function shutdown(code = 0) {
  for (const child of children) {
    if (child.exitCode === null && !child.killed) {
      try {
        child.kill();
      } catch {
        // already gone
      }
    }
  }
  process.exit(code);
}

process.on('SIGINT', () => shutdown(0));
process.on('SIGTERM', () => shutdown(0));

function run(label, command, args, options = {}) {
  // No shell: spawning through cmd.exe breaks on paths containing spaces
  // (e.g. "C:\Program Files\nodejs\node.exe").
  const child = spawn(command, args, {
    stdio: 'inherit',
    shell: false,
    ...options,
  });
  child.on('exit', (code) => {
    if (label === 'electron' && code !== 0) {
      console.error(`[dev] Electron exited with ${code}`);
      shutdown(code ?? 0);
    }
  });
  children.push(child);
  return child;
}

/** Poll until the dev server answers. */
async function waitForServer(timeoutMs = 60_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(URL);
      if (res.status < 500) return true;
    } catch {
      // not listening yet
    }
    await delay(400);
  }
  return false;
}

console.log(`[dev] Vite starting on ${URL}`);
// Invoke Vite's JS entry directly rather than via npm.cmd, which avoids a
// shell round-trip and its quoting warnings.
if (!existsSync(VITE_BIN)) {
  console.error(`[dev] Vite not found at ${VITE_BIN} — run npm install first`);
  shutdown(1);
}
run('vite', process.execPath, [VITE_BIN, '--port', String(PORT), '--strictPort']);

if (!(await waitForServer())) {
  console.error('[dev] Vite did not come up in time');
  shutdown(1);
}

console.log(`[dev] Vite is up — launching Electron against ${URL}`);
console.log('[dev] DevTools: Ctrl+Shift+I (or View > Toggle DevTools)');

// `require('electron')` resolves to the electron.exe path, so spawn it directly
// with the project folder as the app argument.
const electronBin = require('electron');
run('electron', electronBin, [ROOT], {
  env: { ...process.env, SHEETCRAFT_DEV: '1', SHEETCRAFT_DEV_URL: URL },
});
