import { twipsToInches } from './ooxml/ns.js';
import { estimatePages } from './pageEstimate.js';
import { sectionHasContent } from './build/matter.js';
import type {
  FormatOptions,
  ManuscriptAnalysis,
  ReferenceProfile,
} from './types.js';

/**
 * A plain-English check over the book before it is made. Everything here is
 * something an author can act on; nothing is phrased in typesetting terms.
 *
 * It is deliberately not a guarantee. KDP's own Print Previewer is the check
 * that decides what gets printed, and the report says so.
 */

export type CheckLevel = 'ready' | 'check' | 'attention';

export interface PreflightCheck {
  id: string;
  level: CheckLevel;
  /** One line, readable on its own. */
  title: string;
  /** What it means and what to do about it. */
  detail: string;
  /**
   * The paragraphs this is about, so the reviewer can be shown the actual
   * place rather than having to hunt for it.
   */
  examples?: Array<{ index: number; preview: string }>;
}

export interface PreflightReport {
  /** The most serious level present, so the UI can lead with it. */
  level: CheckLevel;
  checks: PreflightCheck[];
  /** Rough page count, for the margin advice. Null when it cannot be judged. */
  estimatedPages: number | null;
}

/** Wording that means a template's own filler was never replaced. */
const PLACEHOLDER_PATTERNS: RegExp[] = [
  /\binsert (your |chapter |dedication|acknowledg)/i,
  /\blorem ipsum\b/i,
  /\[(your|book|author|title|insert)[^\]]*\]/i,
  /\b(chapter|book|author) (name|title) (goes )?here\b/i,
  /\btype (your|here)\b/i,
  /\bxxx-xxx\b/i,
  /\btemplate id\b/i,
  /\bplaceholder\b/i,
  /\bTK\b/,
];

/**
 * KDP's inside-margin requirement grows with the finished page count. The app
 * cannot know that count, so this drives advice rather than a hard rule.
 */
const GUTTER_BY_PAGES: Array<{ upTo: number; inches: number }> = [
  { upTo: 150, inches: 0.375 },
  { upTo: 300, inches: 0.5 },
  { upTo: 500, inches: 0.625 },
  { upTo: 700, inches: 0.75 },
  { upTo: Number.POSITIVE_INFINITY, inches: 0.875 },
];

export interface PreflightInput {
  profile: ReferenceProfile;
  analysis: ManuscriptAnalysis;
  options: FormatOptions;
}

export function preflight(input: PreflightInput): PreflightReport {
  const { profile, analysis, options } = input;
  const checks: PreflightCheck[] = [];
  const add = (check: PreflightCheck): void => {
    checks.push(check);
  };

  const estimatedPages = estimatePageCount(profile, analysis);

  // --- chapters -----------------------------------------------------------
  const chapters = analysis.blocks.filter(
    (b) => (options.roleOverrides[b.index] ?? b.role) === 'chapterTitle',
  ).length;
  if (chapters === 0) {
    add({
      id: 'no-chapters',
      level: 'attention',
      title: 'No chapters were recognised',
      detail:
        'The whole book will run as one long chapter. Open “Your chapters and headings”, tick ' +
        '“Show every paragraph”, and mark your chapter titles so each one starts its own page.',
    });
  } else {
    add({
      id: 'chapters',
      level: 'ready',
      title: `${chapters} chapter${chapters === 1 ? '' : 's'} found`,
      detail: 'Each one will begin on its own page.',
    });
  }

  const unsureBlocks = analysis.blocks.filter(
    (b) => b.confidence < 0.6 && b.role !== 'body' && b.role !== 'empty',
  );
  if (unsureBlocks.length > 0) {
    add({
      id: 'uncertain',
      level: 'check',
      title: `${unsureBlocks.length} line${unsureBlocks.length === 1 ? ' was' : 's were'} hard to place`,
      detail:
        'These are marked “not sure” in the list of chapters and headings. Glance at them and ' +
        'correct any that were read as the wrong kind of line.',
      examples: examplesOf(unsureBlocks),
    });
  }

  // --- leftover template wording ------------------------------------------
  const placeholders = analysis.blocks.filter(
    (b) => !b.isEmpty && PLACEHOLDER_PATTERNS.some((re) => re.test(b.text)),
  );
  if (placeholders.length > 0) {
    add({
      id: 'placeholder-text',
      level: 'attention',
      title: `${placeholders.length} line${placeholders.length === 1 ? '' : 's'} still look like filler text`,
      detail:
        'This looks like wording from a template that was never replaced. Fix it in your ' +
        'manuscript, then load it again.',
      examples: examplesOf(placeholders),
    });
  }

  const headerPlaceholder = profile.headerFooterText.filter((text) =>
    PLACEHOLDER_PATTERNS.some((re) => re.test(text)),
  );
  if (headerPlaceholder.length > 0) {
    add({
      id: 'header-placeholder',
      level: 'attention',
      title: 'The page headers still say the template’s words',
      detail:
        `The top or bottom of every page reads “${headerPlaceholder[0]}”. Open the finished file ` +
        'in Word, double-click that area, and type your own book title and name.',
    });
  } else if (profile.headerFooterText.length > 0) {
    add({
      id: 'header-text',
      level: 'check',
      title: 'Check the words in your page headers',
      detail:
        `They currently read “${profile.headerFooterText.join('” / “')}”. If that is not your ` +
        'book title and name, change it in Word by double-clicking the top of a page.',
    });
  }

  // --- page numbers -------------------------------------------------------
  if (!profile.hasPageNumbers) {
    add({
      id: 'page-numbers',
      level: 'attention',
      title: 'Your design has no page numbers',
      detail:
        'Printed books need them. In Word, choose Insert → Page Number, or pick one of the ' +
        'app’s own designs under Quick Start, which include them.',
    });
  } else {
    add({
      id: 'page-numbers-ok',
      level: 'ready',
      title: 'Page numbers are set up',
      detail: 'They carry over from your design.',
    });
  }

  // --- margins against the likely page count ------------------------------
  const gutterIn = twipsToInches(profile.pageSetup.margins.gutter);
  const insideIn = twipsToInches(profile.pageSetup.margins.left) + gutterIn;
  if (estimatedPages !== null) {
    const needed = GUTTER_BY_PAGES.find((row) => estimatedPages <= row.upTo)!.inches;
    if (insideIn + 0.01 < needed) {
      add({
        id: 'gutter',
        level: 'check',
        title: 'The inside margin may be tight for a book this long',
        detail:
          `Your book looks like roughly ${estimatedPages} pages, and KDP wants at least ` +
          `${needed}" on the inside edge at that length. Yours is about ${round(insideIn)}". ` +
          'This is only an estimate — check it in KDP’s Print Previewer, which knows the real count.',
      });
    } else {
      add({
        id: 'gutter-ok',
        level: 'ready',
        title: 'Margins suit the likely length',
        detail:
          `About ${estimatedPages} pages, with roughly ${round(insideIn)}" on the inside edge. ` +
          'KDP’s Print Previewer has the final say once it knows the real page count.',
      });
    }
  }

  // --- things that will not fit -------------------------------------------
  const textWidth =
    profile.pageSetup.widthTwips -
    profile.pageSetup.margins.left -
    profile.pageSetup.margins.right -
    profile.pageSetup.margins.gutter;

  const wideImages = analysis.blocks.filter(
    (b) => b.imageWidthTwips !== null && b.imageWidthTwips > textWidth,
  );
  if (wideImages.length > 0) {
    add({
      id: 'wide-images',
      level: 'attention',
      title: `${wideImages.length} picture${wideImages.length === 1 ? ' is' : 's are'} wider than the page allows`,
      detail:
        `Your text is about ${round(twipsToInches(textWidth))}" wide. Those pictures will spill ` +
        'past the margin. Resize them in Word after formatting, or before you load the manuscript.',
      examples: examplesOf(wideImages),
    });
  }

  const wideTables = analysis.blocks.filter(
    (b) => b.tableWidthTwips !== null && b.tableWidthTwips > textWidth,
  );
  if (wideTables.length > 0) {
    add({
      id: 'wide-tables',
      level: 'check',
      title: `${wideTables.length} table${wideTables.length === 1 ? ' is' : 's are'} wider than the page`,
      detail:
        'Tables are copied at their original size. Check them in Word and narrow any that run ' +
        'past the margin.',
      examples: examplesOf(wideTables),
    });
  } else if (analysis.tableCount > 0) {
    add({
      id: 'tables',
      level: 'check',
      title: `${analysis.tableCount} table${analysis.tableCount === 1 ? '' : 's'} copied across`,
      detail: 'Tables are carried over as they are. Give them a look in Word.',
    });
  }

  // --- blank pages --------------------------------------------------------
  if (options.chapterStart === 'oddPage') {
    add({
      id: 'blank-pages',
      level: 'check',
      title: 'Some left-hand pages will be blank',
      detail:
        'You have chosen for chapters to start on a right-hand page, so a blank page is added ' +
        'whenever one is needed. That is normal in printed books, and KDP accepts it.',
    });
  }

  // --- duplicated opening pages -------------------------------------------
  const generating = (
    ['titlePage', 'copyrightPage', 'dedication'] as const
  ).some((key) => options.extraSections[key] && sectionHasContent(key, options.bookDetails));
  const manuscriptHasFront = analysis.blocks.some(
    (b) => b.role === 'frontMatterTitle' || b.role === 'copyright',
  );
  if (generating && manuscriptHasFront && !options.replaceFrontMatter && options.includeFrontMatter) {
    add({
      id: 'duplicate-front',
      level: 'attention',
      title: 'Your book may end up with two title pages',
      detail:
        'You have asked the app to build opening pages, and your manuscript already has some. ' +
        'Tick “Use these instead of the front pages already in my manuscript”, or turn the ' +
        'built pages off.',
    });
  }

  return { level: worstLevel(checks), checks: sortChecks(checks), estimatedPages };
}

function round(inches: number): number {
  return Math.round(inches * 100) / 100;
}

/** A few of the offending paragraphs, enough to recognise without a wall. */
function examplesOf(
  blocks: ManuscriptAnalysis['blocks'],
  limit = 4,
): Array<{ index: number; preview: string }> {
  return blocks.slice(0, limit).map((b) => ({
    index: b.index,
    preview: b.preview || (b.kind === 'table' ? '[table]' : '[picture]'),
  }));
}

const ORDER: Record<CheckLevel, number> = { attention: 0, check: 1, ready: 2 };

function sortChecks(checks: PreflightCheck[]): PreflightCheck[] {
  return [...checks].sort((a, b) => ORDER[a.level] - ORDER[b.level]);
}

function worstLevel(checks: PreflightCheck[]): CheckLevel {
  if (checks.some((c) => c.level === 'attention')) return 'attention';
  if (checks.some((c) => c.level === 'check')) return 'check';
  return 'ready';
}

/**
 * A rough page count from the size of the text block and the body type. Words
 * per line comes from average character width (about half the point size), and
 * lines per page from the type size and line spacing. Accurate enough to give
 * margin advice, and always presented as an estimate.
 */
export function estimatePageCount(
  profile: ReferenceProfile,
  analysis: ManuscriptAnalysis,
): number | null {
  const sizePt = profile.bodyFontSizePt;
  if (!sizePt) return null;

  const page = profile.pageSetup;
  return estimatePages({
    textWidthIn: twipsToInches(
      page.widthTwips - page.margins.left - page.margins.right - page.margins.gutter,
    ),
    textHeightIn: twipsToInches(page.heightTwips - page.margins.top - page.margins.bottom),
    fontSizePt: sizePt,
    lineSpacing: profile.bodyLineSpacing ? profile.bodyLineSpacing / 240 : 1.15,
    wordCount: analysis.wordCount,
  });
}
