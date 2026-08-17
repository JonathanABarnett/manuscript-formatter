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

/** The three paperback sizes that cover most self-published fiction. */
export const TRIM_SIZES: TrimSize[] = [
  {
    id: '6x9',
    label: '6 × 9 inches',
    note: 'The most common size for paperbacks. Pick this if you are unsure.',
    widthIn: 6,
    heightIn: 9,
    recommended: true,
    margins: { top: 0.8, bottom: 0.8, outside: 0.65, inside: 0.65 },
  },
  {
    id: '5.5x8.5',
    label: '5.5 × 8.5 inches',
    note: 'A slightly smaller, chunkier book. Common for novels and memoirs.',
    widthIn: 5.5,
    heightIn: 8.5,
    margins: { top: 0.75, bottom: 0.75, outside: 0.6, inside: 0.6 },
  },
  {
    id: '5x8',
    label: '5 × 8 inches',
    note: 'Pocket sized. Fewer words per page, so the book runs longer.',
    widthIn: 5,
    heightIn: 8,
    margins: { top: 0.7, bottom: 0.7, outside: 0.55, inside: 0.55 },
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
];

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
