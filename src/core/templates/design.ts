import { estimatePages } from '../pageEstimate.js';

/**
 * The book designs the app can build for itself, so an author who has no
 * template of their own never has to go and find one.
 *
 * Everything here is written for this app. No Amazon file is redistributed and
 * no placeholder text ships inside a generated template.
 */

export interface TrimSize {
  id: string;
  label: string;
  /** Shown under the label to explain when to pick it. */
  note: string;
  widthIn: number;
  heightIn: number;
  recommended?: boolean;
  /** Page margins in inches. `inside` gains the gutter on top. */
  margins: { top: number; bottom: number; outside: number; inside: number };
}

/**
 * Extra inside margin for the binding. KDP scales this with the finished page
 * count; 0.375" covers a book up to 150 pages and 0.5" up to 300, so the
 * larger figure is used and the page-count caveat is surfaced in the app.
 */
export const GUTTER_IN = 0.5;

/** The paperback sizes that cover most self-published fiction. */
export const TRIM_SIZES: TrimSize[] = [
  {
    id: '6x9',
    label: '6 × 9 inches',
    note: 'The usual size for a paperback, and the safest choice if you are unsure.',
    widthIn: 6,
    heightIn: 9,
    recommended: true,
    margins: { top: 0.8, bottom: 0.8, outside: 0.65, inside: 0.65 },
  },
  {
    id: '5.5x8.5',
    label: '5.5 × 8.5 inches',
    note: 'A little smaller in the hand. Common for novels and memoirs.',
    widthIn: 5.5,
    heightIn: 8.5,
    margins: { top: 0.75, bottom: 0.75, outside: 0.6, inside: 0.6 },
  },
  {
    id: '5.25x8',
    label: '5.25 × 8 inches',
    note: 'A shade wider than 5 × 8. Popular for paperback fiction.',
    widthIn: 5.25,
    heightIn: 8,
    margins: { top: 0.75, bottom: 0.75, outside: 0.6, inside: 0.6 },
  },
  {
    id: '5x8',
    // Not "pocket sized": a true pocket paperback is 4.25 x 6.87 inches.
    label: '5 × 8 inches',
    note: 'The smallest of the American sizes, so the same book runs to the most pages.',
    widthIn: 5,
    heightIn: 8,
    margins: { top: 0.7, bottom: 0.7, outside: 0.55, inside: 0.55 },
  },
  {
    id: 'a5',
    label: 'A5 (5.83 × 8.27 inches)',
    note: 'The usual European paperback size. Choose it if most of your readers are in the UK or Europe.',
    // 148 x 210 mm, as Word measures it: 8391 x 11906 twips.
    widthIn: 8391 / 1440,
    heightIn: 11906 / 1440,
    margins: { top: 0.75, bottom: 0.75, outside: 0.6, inside: 0.6 },
  },
];

export interface BookLook {
  id: string;
  label: string;
  note: string;
  /** Fonts shipped with Word on both Windows and macOS, so nothing goes missing. */
  bodyFont: string;
  headingFont: string;
  bodySizePt: number;
  /** First-line indent for running paragraphs, in inches. */
  indentIn: number;
  /** Multiple of single line spacing for body text. */
  lineSpacing: number;
  justified: boolean;
  chapterSizePt: number;
  chapterAlign: 'center' | 'left';
  chapterCaps: boolean;
  chapterBold: boolean;
  /** Blank lines above and below a chapter title, which sink it down the page. */
  chapterBlanksBefore: number;
  chapterBlanksAfter: number;
  /** The mark set between scenes. */
  sceneMark: string;
}

export const BOOK_LOOKS: BookLook[] = [
  {
    id: 'classic',
    label: 'Classic',
    note: 'Traditional and unfussy. Centred chapter titles, indented paragraphs.',
    bodyFont: 'Garamond',
    headingFont: 'Garamond',
    bodySizePt: 11.5,
    indentIn: 0.25,
    lineSpacing: 1.15,
    justified: true,
    chapterSizePt: 18,
    chapterAlign: 'center',
    chapterCaps: true,
    chapterBold: false,
    chapterBlanksBefore: 6,
    chapterBlanksAfter: 2,
    sceneMark: '* * *',
  },
  {
    id: 'modern',
    label: 'Modern',
    note: 'Cleaner and a little more open. Chapter titles sit to the left.',
    bodyFont: 'Georgia',
    headingFont: 'Georgia',
    bodySizePt: 10.5,
    indentIn: 0.22,
    lineSpacing: 1.25,
    justified: true,
    chapterSizePt: 20,
    chapterAlign: 'left',
    chapterCaps: false,
    chapterBold: true,
    chapterBlanksBefore: 4,
    chapterBlanksAfter: 2,
    sceneMark: '§',
  },
  {
    id: 'minimal',
    label: 'Minimal',
    note: 'Quiet and plain. Small chapter titles, nothing decorative.',
    bodyFont: 'Cambria',
    headingFont: 'Cambria',
    bodySizePt: 11,
    indentIn: 0.2,
    lineSpacing: 1.2,
    justified: false,
    chapterSizePt: 14,
    chapterAlign: 'center',
    chapterCaps: false,
    chapterBold: false,
    chapterBlanksBefore: 3,
    chapterBlanksAfter: 1,
    sceneMark: '·',
  },
  {
    // KDP lists large print as its own edition and asks for 16-point type or
    // larger. Ragged right and open spacing are what large-print readers ask for.
    id: 'largePrint',
    label: 'Large Print',
    note: 'Sixteen-point type, open spacing, ragged right. For a large-print edition; expect roughly twice the pages.',
    bodyFont: 'Georgia',
    headingFont: 'Georgia',
    bodySizePt: 16,
    indentIn: 0.3,
    lineSpacing: 1.3,
    justified: false,
    chapterSizePt: 24,
    chapterAlign: 'left',
    chapterCaps: false,
    chapterBold: true,
    chapterBlanksBefore: 3,
    chapterBlanksAfter: 2,
    sceneMark: '* * *',
  },
];

/**
 * Roughly how many pages a book of `wordCount` runs to in this size and look.
 * Uses the same arithmetic as the preflight report, so the figure quoted when
 * choosing a size matches the one quoted after the book is read.
 */
export function estimatePagesForDesign(
  trim: TrimSize,
  look: BookLook,
  wordCount: number,
): number | null {
  return estimatePages({
    textWidthIn: trim.widthIn - trim.margins.inside - trim.margins.outside - GUTTER_IN,
    textHeightIn: trim.heightIn - trim.margins.top - trim.margins.bottom,
    fontSizePt: look.bodySizePt,
    lineSpacing: look.lineSpacing,
    wordCount,
  });
}

export function findTrim(id: string): TrimSize {
  const trim = TRIM_SIZES.find((t) => t.id === id);
  if (!trim) throw new Error(`Unknown book size "${id}".`);
  return trim;
}

export function findLook(id: string): BookLook {
  const look = BOOK_LOOKS.find((l) => l.id === id);
  if (!look) throw new Error(`Unknown book design "${id}".`);
  return look;
}

/** e.g. `6x9-classic`, used as the generated design's cache key. */
export function designId(trimId: string, lookId: string): string {
  return `${trimId}-${lookId}`;
}

/**
 * File name for a generated design, e.g. `6x9 Classic design.docx`. Kept free
 * of characters Windows rejects in a path, since the desktop build writes it
 * to disk before handing it on.
 */
export function designFileName(trimId: string, lookId: string): string {
  const trim = findTrim(trimId);
  const look = findLook(lookId);
  return `${trim.id} ${look.label} design.docx`;
}
