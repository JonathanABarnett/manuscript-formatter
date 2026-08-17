import type { AnalysisResult } from '../core/format.js';
import type { FormatOptions, FormatResult } from '../core/types.js';

/** Channel names shared by the main process and the preload bridge. */
export const CHANNEL = {
  pickDocx: 'mf:pick-docx',
  pickOutput: 'mf:pick-output',
  suggestOutput: 'mf:suggest-output',
  analyze: 'mf:analyze',
  format: 'mf:format',
  reveal: 'mf:reveal',
  open: 'mf:open',
} as const;

export interface AnalyzePayload {
  referencePath: string;
  manuscriptPath: string;
}

export interface FormatPayload {
  referencePath: string;
  manuscriptPath: string;
  outputPath: string;
  options: FormatOptions;
}

/** Errors are values here: the UI shows them instead of the app crashing. */
export type Outcome<T> = { ok: true; value: T } | { ok: false; error: string };

/**
 * The surface the UI talks to. Electron implements it over IPC; the browser
 * build implements it against the in-page engine, so the renderer is identical
 * in both shells.
 */
export interface FormatterApi {
  /** Which shell is hosting the UI. Changes only how the result is delivered. */
  platform: 'desktop' | 'web';
  pickDocx(kind: 'reference' | 'manuscript'): Promise<string | null>;
  pickOutput(defaultPath: string): Promise<string | null>;
  suggestOutput(manuscriptPath: string): Promise<string>;
  analyze(payload: AnalyzePayload): Promise<Outcome<AnalysisResult>>;
  format(payload: FormatPayload): Promise<Outcome<FormatResult>>;
  /** Desktop: show the file in its folder. Web: not offered. */
  reveal(path: string): Promise<void>;
  /** Desktop: open in Word. Web: download the generated file. */
  open(path: string): Promise<void>;
  /** Identifier for a dropped file — an absolute path, or its name on the web. */
  pathForFile(file: File): string;
}

declare global {
  interface Window {
    formatter: FormatterApi;
  }
}

export type { AnalysisResult, FormatOptions, FormatResult };
