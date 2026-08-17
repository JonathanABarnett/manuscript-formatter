import { access, constants, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';
import {
  analyzeDocuments,
  formatToBuffer,
  type AnalysisResult,
  type LoadedManuscript,
  type LoadedReference,
} from '../format.js';
import type { DocxInput, FormatOptions, FormatResult } from '../types.js';

/**
 * Filesystem wrappers around the platform-neutral engine. Only the desktop app
 * and the CLI import this; the browser build never pulls in `node:` modules.
 */

/** Read a .docx from disk into the engine's input shape. */
export async function readDocxFile(filePath: string): Promise<DocxInput> {
  try {
    return { data: await readFile(filePath), name: basename(filePath) };
  } catch (err) {
    throw new Error(`Cannot read "${filePath}": ${(err as Error).message}`);
  }
}

/** Inspect both documents on disk and propose a formatting plan. */
export async function analyzeDocumentPaths(
  referencePath: string,
  manuscriptPath: string,
): Promise<AnalysisResult> {
  const [reference, manuscript] = await Promise.all([
    readDocxFile(referencePath),
    readDocxFile(manuscriptPath),
  ]);
  return analyzeDocuments(reference, manuscript);
}

export interface FormatRequest {
  referencePath: string;
  manuscriptPath: string;
  outputPath: string;
  options?: Partial<FormatOptions>;
  /** Reuse an already-parsed pair, to avoid reading both documents twice. */
  loaded?: { reference: LoadedReference; manuscript: LoadedManuscript };
}

/**
 * Produce the formatted document on disk. Neither input file is modified; the
 * output path must differ from both.
 */
export async function formatManuscript(request: FormatRequest): Promise<FormatResult> {
  const referencePath = resolve(request.referencePath);
  const manuscriptPath = resolve(request.manuscriptPath);
  const outputPath = resolve(request.outputPath);

  if (outputPath === referencePath || outputPath === manuscriptPath) {
    throw new Error(
      'The output file must be different from the reference and the manuscript. ' +
        'Choose another name so your source documents stay untouched.',
    );
  }
  if (extname(outputPath).toLowerCase() !== '.docx') {
    throw new Error('The output file must end in .docx.');
  }
  await assertWritableDirectory(dirname(outputPath));

  const composed = await formatToBuffer({
    reference: request.loaded?.reference ?? (await readDocxFile(referencePath)),
    manuscript: request.loaded?.manuscript ?? (await readDocxFile(manuscriptPath)),
    options: request.options,
  });

  await writeFile(outputPath, composed.data);
  return { outputPath, stats: composed.stats, warnings: composed.warnings };
}

/** `Novel.docx` -> `Novel (formatted).docx`, then ` 2`, ` 3`, ... if taken. */
export async function suggestOutputPath(manuscriptPath: string): Promise<string> {
  const dir = dirname(manuscriptPath);
  const ext = extname(manuscriptPath);
  const base = basename(manuscriptPath, ext);
  let candidate = join(dir, `${base} (formatted).docx`);
  let counter = 2;
  while (await exists(candidate)) {
    candidate = join(dir, `${base} (formatted ${counter}).docx`);
    counter++;
  }
  return candidate;
}

async function exists(path: string): Promise<boolean> {
  try {
    await access(path, constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

async function assertWritableDirectory(dir: string): Promise<void> {
  try {
    await access(dir, constants.W_OK);
  } catch {
    throw new Error(`Cannot write to "${dir}". Pick a folder you have permission to write to.`);
  }
}

export { analyzeDocuments };
export type { AnalysisResult, LoadedManuscript, LoadedReference };
