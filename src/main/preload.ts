import { contextBridge, ipcRenderer, webUtils } from 'electron';
import { CHANNEL, type AnalyzePayload, type FormatPayload, type FormatterApi } from '../shared/ipc.js';

/**
 * The only surface the renderer gets. Every entry is an explicit, typed call —
 * the page never touches Node, the filesystem or Electron directly.
 */
const api: FormatterApi = {
  platform: 'desktop',
  pickDocx: (kind) => ipcRenderer.invoke(CHANNEL.pickDocx, kind),
  useBuiltIn: (request) => ipcRenderer.invoke(CHANNEL.useBuiltIn, request),
  pickOutput: (defaultPath) => ipcRenderer.invoke(CHANNEL.pickOutput, defaultPath),
  suggestOutput: (manuscriptPath) => ipcRenderer.invoke(CHANNEL.suggestOutput, manuscriptPath),
  analyze: (payload: AnalyzePayload) => ipcRenderer.invoke(CHANNEL.analyze, payload),
  format: (payload: FormatPayload) => ipcRenderer.invoke(CHANNEL.format, payload),
  reveal: (path) => ipcRenderer.invoke(CHANNEL.reveal, path),
  open: (path) => ipcRenderer.invoke(CHANNEL.open, path),
  saveChoices: (suggestedName, json) => ipcRenderer.invoke(CHANNEL.saveChoices, suggestedName, json),
  loadChoices: () => ipcRenderer.invoke(CHANNEL.loadChoices),
  pathForFile: (file) => webUtils.getPathForFile(file),
};

contextBridge.exposeInMainWorld('formatter', api);
