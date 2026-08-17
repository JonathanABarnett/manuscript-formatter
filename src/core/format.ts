import { analyzeReference, type LoadedReference } from './analyze/reference.js';
import { analyzeManuscript, type LoadedManuscript } from './analyze/manuscript.js';
import { composeDocument } from './build/compose.js';
import {
  DEFAULT_FORMAT_OPTIONS,
  type DocxInput,
  type FormatOptions,
  type FormatStats,
  type ManuscriptAnalysis,
  type ReferenceProfile,
} from './types.js';

/**
 * Platform-neutral engine API. Nothing here touches the filesystem, so the
 * same code runs in the desktop app, the CLI and the browser. The Node-only
 * path wrappers live in `platform/node.ts`.
 */

export interface AnalysisResult {
  profile: ReferenceProfile;
  analysis: ManuscriptAnalysis;
  /** Options pre-set from what the reference document does. */
  suggestedOptions: FormatOptions;
}

/** Inspect both documents and propose a formatting plan. Nothing is written. */
export async function analyzeDocuments(
  reference: DocxInput,
  manuscript: DocxInput,
): Promise<AnalysisResult> {
  const [loadedReference, loadedManuscript] = await Promise.all([
    analyzeReference(reference),
    analyzeManuscript(manuscript),
  ]);
  return {
    profile: loadedReference.profile,
    analysis: loadedManuscript.analysis,
    suggestedOptions: suggestOptions(loadedReference.profile, loadedManuscript.analysis),
  };
}

/**
 * Formatting defaults that match what the reference document already does.
 * When the manuscript is given too, a contents page is offered by default
 * unless the book already has one.
 */
export function suggestOptions(
  profile: ReferenceProfile,
  analysis?: ManuscriptAnalysis,
): FormatOptions {
  const wantsContents = analysis ? !analysis.hasContentsPage && analysis.chapterCount > 1 : false;
  return {
    ...DEFAULT_FORMAT_OPTIONS,
    chapterStart: profile.chapterStartsOnOddPage
      ? 'oddPage'
      : profile.chapterStartsOnNewPage
        ? 'newPage'
        : 'continuous',
    firstParagraphNoIndent: profile.usesFirstParagraphNoIndent,
    extraSections: { ...DEFAULT_FORMAT_OPTIONS.extraSections, contents: wantsContents },
  };
}

export interface ComposeRequest {
  /** Raw bytes, or a document already parsed by `analyzeDocuments`. */
  reference: DocxInput | LoadedReference;
  manuscript: DocxInput | LoadedManuscript;
  options?: Partial<FormatOptions>;
}

export interface ComposedDocument {
  data: Uint8Array;
  fileName: string;
  stats: FormatStats;
  warnings: string[];
}

const isRaw = (value: DocxInput | { profile: unknown } | { analysis: unknown }): value is DocxInput =>
  'data' in value;

/**
 * Produce the formatted document as bytes. Neither input is modified — the
 * output is built from a copy of the reference package. Passing the parsed
 * documents from `analyzeDocuments` avoids reading them a second time.
 */
export async function formatToBuffer(request: ComposeRequest): Promise<ComposedDocument> {
  const reference = isRaw(request.reference)
    ? await analyzeReference(request.reference)
    : request.reference;
  const manuscript = isRaw(request.manuscript)
    ? await analyzeManuscript(request.manuscript)
    : request.manuscript;

  const options: FormatOptions = {
    ...suggestOptions(reference.profile, manuscript.analysis),
    ...request.options,
    roleStyles: { ...(request.options?.roleStyles ?? {}) },
    roleOverrides: { ...(request.options?.roleOverrides ?? {}) },
  };

  const { pkg, stats, warnings } = await composeDocument(reference, manuscript, options);
  return {
    data: await pkg.toBuffer(),
    fileName: suggestOutputName(manuscript.analysis.fileName),
    stats,
    warnings: [...reference.profile.warnings, ...manuscript.analysis.warnings, ...warnings],
  };
}

/** `Novel.docx` -> `Novel (formatted).docx`. */
export function suggestOutputName(manuscriptName: string): string {
  const base = manuscriptName.replace(/\.docx$/i, '');
  return `${base} (formatted).docx`;
}

export { analyzeReference, analyzeManuscript };
export type { LoadedReference, LoadedManuscript };
