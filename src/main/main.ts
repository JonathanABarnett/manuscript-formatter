import { BrowserWindow, app, dialog, ipcMain, shell } from 'electron';
import { join } from 'node:path';
import { CHANNEL, type AnalyzePayload, type FormatPayload, type Outcome } from '../shared/ipc.js';
import { analyzeManuscript, analyzeReference, suggestOptions } from '../core/format.js';
import {
  formatManuscript,
  readDocxFile,
  suggestOutputPath,
  type AnalysisResult,
  type LoadedManuscript,
  type LoadedReference,
} from '../core/platform/node.js';
import type { FormatResult } from '../core/types.js';

const DOCX_FILTER = [{ name: 'Word documents', extensions: ['docx'] }];

/**
 * Parsed documents are cached by path so pressing Format does not re-read and
 * re-classify what the review screen already analyzed.
 */
const referenceCache = new Map<string, LoadedReference>();
const manuscriptCache = new Map<string, LoadedManuscript>();

let mainWindow: BrowserWindow | null = null;

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1220,
    height: 840,
    minWidth: 940,
    minHeight: 620,
    title: 'Manuscript Formatter',
    backgroundColor: '#f6f5f2',
    show: false,
    webPreferences: {
      preload: join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.once('ready-to-show', () => mainWindow?.show());
  void mainWindow.loadFile(join(__dirname, '../renderer/index.html'));
  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // Anything that wants to leave the app opens in the user's browser instead.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
}

/** Turn a thrown error into a message the UI can show verbatim. */
async function attempt<T>(work: () => Promise<T>): Promise<Outcome<T>> {
  try {
    return { ok: true, value: await work() };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}

async function loadReference(path: string): Promise<LoadedReference> {
  const cached = referenceCache.get(path);
  if (cached) return cached;
  const loaded = await analyzeReference(await readDocxFile(path));
  referenceCache.clear();
  referenceCache.set(path, loaded);
  return loaded;
}

async function loadManuscript(path: string): Promise<LoadedManuscript> {
  const cached = manuscriptCache.get(path);
  if (cached) return cached;
  const loaded = await analyzeManuscript(await readDocxFile(path));
  manuscriptCache.clear();
  manuscriptCache.set(path, loaded);
  return loaded;
}

function registerHandlers(): void {
  ipcMain.handle(CHANNEL.pickDocx, async (_event, kind: 'reference' | 'manuscript') => {
    const result = await dialog.showOpenDialog({
      title:
        kind === 'reference'
          ? 'Choose the formatted document to copy the layout from'
          : 'Choose the manuscript to format',
      buttonLabel: 'Choose',
      properties: ['openFile'],
      filters: DOCX_FILTER,
    });
    return result.canceled ? null : (result.filePaths[0] ?? null);
  });

  ipcMain.handle(CHANNEL.pickOutput, async (_event, defaultPath: string) => {
    const result = await dialog.showSaveDialog({
      title: 'Save the formatted document as',
      buttonLabel: 'Save',
      defaultPath,
      filters: DOCX_FILTER,
    });
    return result.canceled ? null : (result.filePath ?? null);
  });

  ipcMain.handle(CHANNEL.suggestOutput, async (_event, manuscriptPath: string) =>
    suggestOutputPath(manuscriptPath),
  );

  ipcMain.handle(CHANNEL.analyze, async (_event, payload: AnalyzePayload) =>
    attempt<AnalysisResult>(async () => {
      const [reference, manuscript] = await Promise.all([
        loadReference(payload.referencePath),
        loadManuscript(payload.manuscriptPath),
      ]);
      return {
        profile: reference.profile,
        analysis: manuscript.analysis,
        suggestedOptions: suggestOptions(reference.profile),
      };
    }),
  );

  ipcMain.handle(CHANNEL.format, async (_event, payload: FormatPayload) =>
    attempt<FormatResult>(async () => {
      const reference = await loadReference(payload.referencePath);
      const manuscript = await loadManuscript(payload.manuscriptPath);
      return formatManuscript({
        referencePath: payload.referencePath,
        manuscriptPath: payload.manuscriptPath,
        outputPath: payload.outputPath,
        options: payload.options,
        loaded: { reference, manuscript },
      });
    }),
  );

  ipcMain.handle(CHANNEL.reveal, async (_event, path: string) => {
    shell.showItemInFolder(path);
  });

  ipcMain.handle(CHANNEL.open, async (_event, path: string) => {
    await shell.openPath(path);
  });
}

app.whenReady().then(() => {
  registerHandlers();
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
