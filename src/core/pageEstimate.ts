/**
 * How long a book will run. One piece of arithmetic, shared by the size
 * picker (which works from a design's constants) and the preflight report
 * (which works from a document that has already been read), so the two can
 * never quote different figures for the same book.
 *
 * It is an estimate and is always presented as one. Word decides the real
 * count when it paginates, and KDP's Print Previewer confirms it.
 */

export interface PageEstimateInput {
  /** Width of the text block, inside the margins and gutter. */
  textWidthIn: number;
  /** Height of the text block, inside the top and bottom margins. */
  textHeightIn: number;
  fontSizePt: number;
  /** Multiple of single line spacing, e.g. 1.15. */
  lineSpacing: number;
  wordCount: number;
}

/**
 * Turn Word's line spacing into a plain multiple of the type size. Under the
 * `auto` rule the value counts 240ths of a line; under `exact` and `atLeast`
 * it is an absolute height in twips, which has to be measured against the
 * type size instead.
 */
export function lineSpacingMultiple(
  value: number | null,
  rule: string | null,
  fontSizePt: number,
): number {
  const DEFAULT = 1.15;
  if (!value || value <= 0 || fontSizePt <= 0) return DEFAULT;
  if (rule === 'exact' || rule === 'atLeast') {
    // twips -> points -> multiple of the type size
    return value / 20 / fontSizePt;
  }
  return value / 240;
}

/**
 * An average word runs about 6.5 characters including its trailing space, and
 * a character is roughly half the point size wide. That gives words per line;
 * the type size and line spacing give lines per page.
 */
const CHARS_PER_WORD = 6.5;
const CHAR_WIDTH_RATIO = 0.5;

export function estimatePages(input: PageEstimateInput): number | null {
  const { textWidthIn, textHeightIn, fontSizePt, lineSpacing, wordCount } = input;
  if (textWidthIn <= 0 || textHeightIn <= 0 || fontSizePt <= 0 || wordCount <= 0) return null;

  const wordWidthIn = CHARS_PER_WORD * CHAR_WIDTH_RATIO * (fontSizePt / 72);
  const wordsPerLine = textWidthIn / wordWidthIn;

  const lineHeightIn = (fontSizePt * lineSpacing) / 72;
  const linesPerPage = textHeightIn / lineHeightIn;

  const wordsPerPage = wordsPerLine * linesPerPage;
  if (!Number.isFinite(wordsPerPage) || wordsPerPage <= 0) return null;
  return Math.max(1, Math.round(wordCount / wordsPerPage));
}

/** The length of novel the size picker quotes its page counts against. */
export const TYPICAL_NOVEL_WORDS = 80_000;
