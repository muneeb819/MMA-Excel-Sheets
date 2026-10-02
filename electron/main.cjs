/**
 * Electron main process.
 *
 * The renderer is a static bundle, so the window simply loads dist/index.html
 * over file://. Everything the app needs (SheetJS, the sql.js wasm) is packaged
 * next to it, which keeps the whole thing offline.
 */

const { app, BrowserWindow, Menu, dialog, shell, ipcMain } = require('electron');
const path = require('node:path');
const fs = require('node:fs');

const isDev = !app.isPackaged && process.env.SHEETCRAFT_DEV === '1';

let mainWindow = null;

/** Create the application window. */
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 900,
    minHeight: 560,
    show: false,
    backgroundColor: '#f6f7f9',
    title: 'SheetCraft',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: false,
    },
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    if (isDev) mainWindow.webContents.openDevTools({ mode: 'detach' });
  });

  const indexPath = path.join(__dirname, '..', 'dist', 'index.html');
  if (isDev && process.env.SHEETCRAFT_DEV_URL) {
    void mainWindow.loadURL(process.env.SHEETCRAFT_DEV_URL);
  } else {
    void mainWindow.loadFile(indexPath);
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // External links open in the user's browser, never inside the app.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url);
    return { action: 'deny' };
  });
}

/** Menu: the same file actions the ribbon offers, with accelerators. */
function buildMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'Save workbook',
          accelerator: 'CmdOrCtrl+S',
          click: () => mainWindow?.webContents.send('menu', 'save'),
        },
        {
          label: 'Open workbook…',
          accelerator: 'CmdOrCtrl+O',
          click: () => mainWindow?.webContents.send('menu', 'open'),
        },
        { type: 'separator' },
        {
          label: 'Export selection as CSV…',
          click: () => mainWindow?.webContents.send('menu', 'export-csv'),
        },
        {
          label: 'Export workbook as XLSX…',
          click: () => mainWindow?.webContents.send('menu', 'export-xlsx'),
        },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
      ],
    },
    {
      label: 'View',
      submenu: [
        { label: 'Dashboard', click: () => mainWindow?.webContents.send('menu', 'pane:dashboard') },
        { label: 'Charts', click: () => mainWindow?.webContents.send('menu', 'pane:charts') },
        { label: 'Clean data', click: () => mainWindow?.webContents.send('menu', 'pane:data') },
        { label: 'SQL', click: () => mainWindow?.webContents.send('menu', 'pane:sql') },
        { type: 'separator' },
        { role: 'reload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
      ],
    },
    {
      label: 'Help',
      submenu: [
        {
          label: 'About SheetCraft',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: 'About SheetCraft',
              message: `SheetCraft ${app.getVersion()}`,
              detail:
                'An offline spreadsheet suite.\n\n' +
                'Formulas, charts, data cleaning, dashboards and local SQLite all run on this machine.',
              buttons: ['Close'],
            });
          },
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/* ------------------------------------------------------------ IPC handlers */

/** Native open dialog returning file bytes plus the chosen path. */
ipcMain.handle('file:open', async (_event, filters) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: filters ?? [{ name: 'Workbooks', extensions: ['xlsx', 'xlsm', 'csv', 'tsv', 'txt'] }],
  });
  if (result.canceled || !result.filePaths.length) return null;
  const filePath = result.filePaths[0];
  const bytes = fs.readFileSync(filePath);
  return { name: path.basename(filePath), path: filePath, bytes: Array.from(bytes) };
});

/** Native save dialog; `bytes` arrives as a plain array. */
ipcMain.handle('file:save', async (_event, defaultName, bytes, filters) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultName,
    filters: filters ?? [{ name: 'Workbook', extensions: ['xlsx'] }],
  });
  if (result.canceled || !result.filePath) return null;
  fs.writeFileSync(result.filePath, Buffer.from(bytes));
  return result.filePath;
});

/* --------------------------------------------------------------- lifecycle */

void app.whenReady().then(() => {
  buildMenu();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
