/**
 * Preload bridge.
 *
 * Exposes a tiny, explicit surface to the renderer: native open/save dialogs
 * and a menu-event channel. No Node APIs leak into page scripts.
 */

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('sheetcraftDesktop', {
  /** true when running inside the desktop shell */
  isDesktop: true,

  /**
   * Show an open dialog and read the file.
   * @param {Array<{name: string, extensions: string[]}>} [filters]
   * @returns {Promise<{name: string, path: string, bytes: number[]}|null>}
   */
  async openFile(filters) {
    const result = await ipcRenderer.invoke('file:open', filters);
    return result ? { name: result.name, path: result.path, bytes: Uint8Array.from(result.bytes) } : null;
  },

  /**
   * Show a save dialog and write the bytes.
   * @returns {Promise<string|null>} the chosen path
   */
  async saveFile(defaultName, bytes, filters) {
    return ipcRenderer.invoke('file:save', defaultName, Array.from(bytes), filters);
  },

  /** Subscribe to menu commands; returns an unsubscribe function. */
  onMenu(handler) {
    const listener = (_event, action) => handler(action);
    ipcRenderer.on('menu', listener);
    return () => ipcRenderer.removeListener('menu', listener);
  },
});
