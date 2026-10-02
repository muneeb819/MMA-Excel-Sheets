/** Type surface for the desktop bridge exposed by the preload script. */

export interface DesktopBridge {
  isDesktop: true;
  openFile(filters?: { name: string; extensions: string[] }[]): Promise<{
    name: string;
    path: string;
    bytes: Uint8Array;
  } | null>;
  saveFile(
    defaultName: string,
    bytes: Uint8Array,
    filters?: { name: string; extensions: string[] }[],
  ): Promise<string | null>;
  onMenu(handler: (action: string) => void): () => void;
}

declare global {
  interface Window {
    sheetcraftDesktop?: DesktopBridge;
  }
}

export {};
