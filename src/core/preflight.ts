import { twipsToInches } from './ooxml/ns.js';
import { estimatePages, lineSpacingMultiple } from './pageEstimate.js';
import { sectionHasContent } from './build/matter.js';
import { willReplaceRunningHead } from './build/headers.js';
import { checkChapterNumbers } from './analyze/chapterNumbers.js';
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

  // --- chapter numbers ----------------------------------------------------
  const chapterBlocks = analysis.blocks.filter(
    (b) => (options.roleOverrides[b.index] ?? b.role) === 'chapterTitle',
  );
  const numbering = checkChapterNumbers(chapterBlocks);
  if (numbering.numbered.length >= 2) {
    const byNumber = (n: number) =>
      chapterBlocks.filter((b) => numbering.numbered.some((x) => x.index === b.index && x.number === n));
    const wrong = [
      ...numbering.duplicates.flatMap(byNumber),
      ...numbering.outOfOrder.flatMap(byNumber),
    ];
    const problems: string[] = [];
    if (numbering.duplicates.length > 0) {
      problems.push(
        `${numbering.duplicates.length === 1 ? 'number' : 'numbers'} ${numbering.duplicates.join(', ')} ` +
          `${numbering.duplicates.length === 1 ? 'is' : 'are'} used twice`,
      );
    }
    if (numbering.gaps.length > 0) {
      problems.push(
        `${numbering.gaps.length === 1 ? 'number' : 'numbers'} ${numbering.gaps.join(', ')} ` +
          `${numbering.gaps.length === 1 ? 'is' : 'are'} missing`,
      );
    }
    if (numbering.outOfOrder.length > 0) {
      problems.push(`${numbering.outOfOrder.join(', ')} ${numbering.outOfOrder.length === 1 ? 'comes' : 'come'} out of order`);
    }
    if (options.renumberChapters) {
      add({
        id: 'chapter-numbers-renumbered',
        level: 'ready',
        title: `Chapters will be numbered 1 to ${numbering.numbered.length} in order`,
        detail:
          problems.length > 0
            ? `As written, ${problems.join(' and ')}. Renumbering fixes that.`
            : 'Every numbered chapter is written in sequence.',
      });
    } else if (problems.length > 0) {
      add({
        id: 'chapter-numbers',
        level: 'check',
        title: 'Chapter numbers do not run in sequence',
        detail:
          `As written, ${problems.join(' and ')}. If the chapters are in the right order, tick ` +
          '“Renumber chapters in order” under Adjustments and the app will number them 1, 2, 3… for you.',
        examples: examplesOf(wrong.length > 0 ? wrong : chapterBlocks),
      });
    }
    if (numbering.mixed && options.chapterNumberStyle === 'keep') {
      add({
        id: 'chapter-numbers-mixed',
        level: 'check',
        title: 'Chapter numbers are written in more than one way',
        detail:
          'Some chapters spell the number out, use figures or numerals, or carry the word ' +
          '“Chapter” where others do not. Choose a style under “Chapter numbers” in Adjustments ' +
          'to make every chapter opening match.',
        examples: examplesOf(chapterBlocks.filter((b) => numbering.numbered.some((x) => x.index === b.index))),
      });
    }
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

  // Headers the typed-in details will overwrite are not worth warning about.
  const remainingHeads = profile.headerFooterText.filter(
    (text) => !willReplaceRunningHead(text, options.bookDetails),
  );
  const replacedHeads = profile.headerFooterText.length - remainingHeads.length;
  if (replacedHeads > 0) {
    add({
      id: 'headers-updated',
      level: 'ready',
      title: 'Your title and name go into the page headers',
      detail:
        `The design's placeholder wording is replaced on ${replacedHeads} of them, using the ` +
        'details you typed in.',
    });
  }

  const headerPlaceholder = remainingHeads.filter((text) =>
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
  } else if (remainingHeads.length > 0) {
    add({
      id: 'header-text',
      level: 'check',
      title: 'Check the words in your page headers',
      detail:
        `They currently read “${remainingHeads.join('” / “')}”. If that is not your ` +
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

  // KDP will not print a paperback shorter than 24 pages or longer than 828.
  if (estimatedPages !== null && estimatedPages < 24) {
    add({
      id: 'too-short',
      level: 'attention',
      title: 'The book may be too short for KDP to print',
      detail:
        `KDP needs at least 24 pages and this looks like roughly ${estimatedPages}. A larger type ` +
        'size, more space between lines, or a smaller page all add pages. Only an estimate — ' +
        'KDP’s Print Previewer counts the real pages.',
    });
  } else if (estimatedPages !== null && estimatedPages > 828) {
    add({
      id: 'too-long',
      level: 'attention',
      title: 'The book may be too long for KDP to print in one volume',
      detail:
        `KDP’s paperback limit is 828 pages and this looks like roughly ${estimatedPages}. A larger ` +
        'page size or a slightly smaller type size brings it down; otherwise it needs splitting ' +
        'into volumes. Only an estimate — KDP’s Print Previewer counts the real pages.',
    });
  }

  // --- the ISBN on the copyright page ---------------------------------------
  const isbn = options.bookDetails.isbn.trim();
  if (isbn && options.extraSections.copyrightPage) {
    const verdict = checkIsbn(isbn);
    if (verdict === 'ok') {
      add({
        id: 'isbn-ok',
        level: 'ready',
        title: 'The ISBN checks out',
        detail: 'Its check digit is right, so it was copied correctly.',
      });
    } else {
      add({
        id: 'isbn',
        level: 'attention',
        title: 'The ISBN does not look right',
        detail:
          verdict === 'length'
            ? `“${isbn}” has the wrong number of digits. An ISBN has 13 (older ones have 10). Check it against the one KDP or your ISBN agency gave you.`
            : `“${isbn}” fails its own check digit, which usually means a typo. Check it against the one KDP or your ISBN agency gave you.`,
      });
    }
  }

  // --- chapters with nothing in them, and scene breaks with nothing between ---
  const roleAt = (b: ManuscriptAnalysis['blocks'][number]) => options.roleOverrides[b.index] ?? b.role;
  const emptyChapters: ManuscriptAnalysis['blocks'] = [];
  const strandedBreaks: ManuscriptAnalysis['blocks'] = [];
  {
    /** The last thing seen that was not blank: what kind of line it was. */
    let previous: 'title' | 'break' | 'text' | null = null;
    let previousTitle: ManuscriptAnalysis['blocks'][number] | null = null;
    let previousBreak: ManuscriptAnalysis['blocks'][number] | null = null;
    for (const b of analysis.blocks) {
      const role = b.kind === 'table' ? 'body' : roleAt(b);
      if (role === 'empty' || role === 'pageBreak' || role === 'chapterSubtitle') continue;
      if (role === 'chapterTitle' || role === 'partTitle') {
        if (previous === 'title' && previousTitle && roleAt(previousTitle) === 'chapterTitle') {
          emptyChapters.push(previousTitle);
        }
        if (previous === 'break' && previousBreak) strandedBreaks.push(previousBreak);
        previous = 'title';
        previousTitle = b;
        continue;
      }
      if (role === 'sceneBreak') {
        if (previous === 'title' || previous === 'break') strandedBreaks.push(b);
        previous = 'break';
        previousBreak = b;
        continue;
      }
      previous = 'text';
    }
    if (previous === 'title' && previousTitle && roleAt(previousTitle) === 'chapterTitle') {
      emptyChapters.push(previousTitle);
    }
    if (previous === 'break' && previousBreak) strandedBreaks.push(previousBreak);
  }
  if (emptyChapters.length > 0) {
    add({
      id: 'empty-chapters',
      level: 'check',
      title: `${emptyChapters.length} chapter${emptyChapters.length === 1 ? ' has' : 's have'} nothing under the title`,
      detail:
        'A chapter title is followed straight away by the next title, or by the end of the book. ' +
        'Either the text is missing or the line is not really a chapter title — change what it is ' +
        'in the list of chapters and headings.',
      examples: examplesOf(emptyChapters),
    });
  }
  if (strandedBreaks.length > 0) {
    add({
      id: 'stranded-scene-breaks',
      level: 'check',
      title: `${strandedBreaks.length} scene break${strandedBreaks.length === 1 ? ' has' : 's have'} nothing on one side`,
      detail:
        'A scene-break mark sits at the very start or end of a chapter, or right after another ' +
        'one. Printed books do not open or close a chapter with one, so these are probably left ' +
        'over from editing. Delete them in your manuscript, or mark them as ordinary lines.',
      examples: examplesOf(strandedBreaks),
    });
  }

  // --- tidy-ups the options can do, and what they would touch ---------------
  const habit = (
    id: string,
    count: number,
    on: boolean,
    what: string,
    onDetail: string,
    offDetail: string,
    onTitle: string,
  ): void => {
    if (count === 0) return;
    add(
      on
        ? { id: `${id}-on`, level: 'ready', title: onTitle, detail: onDetail }
        : { id, level: 'check', title: what, detail: offDetail },
    );
  };
  const n = (count: number, one: string, many: string): string =>
    `${count.toLocaleString()} ${count === 1 ? one : many}`;
  habit(
    'straight-quotes',
    analysis.straightQuoteCount,
    options.smartTypography,
    `${n(analysis.straightQuoteCount, 'straight quote or apostrophe', 'straight quotes and apostrophes')} will print as typed`,
    `${n(analysis.straightQuoteCount, 'straight quote or apostrophe is', 'straight quotes and apostrophes are')} being changed to the curly kind, and -- and ... to a dash and an ellipsis.`,
    'Printed books use curly quotes (“ ” ’). Tick “Fix straight quotes, dashes, and ellipses” under Adjustments to change them — unless your manuscript has already been through an editor and you want it left exactly as it is.',
    'Quotes and apostrophes are being set the printed way',
  );
  habit(
    'double-spaces',
    analysis.doubleSpaceCount,
    options.collapseMultipleSpaces,
    `${n(analysis.doubleSpaceCount, 'place has', 'places have')} two or more spaces in a row`,
    `${n(analysis.doubleSpaceCount, 'run of doubled spaces is', 'runs of doubled spaces are')} being reduced to one.`,
    'Usually two spaces after a full stop, a typewriter habit. Printed books use one. Tick “Reduce multiple spaces to one” under Adjustments to fix them all.',
    'Doubled spaces are being reduced to one',
  );
  habit(
    'underlining',
    analysis.underlinedRunCount,
    options.underlineToItalic && options.keepEmphasis,
    `${n(analysis.underlinedRunCount, 'underlined passage', 'underlined passages')} will print underlined`,
    `${n(analysis.underlinedRunCount, 'underlined passage is', 'underlined passages are')} being set in italics instead, as a printed book would.`,
    'Typed manuscripts underline what a printed book sets in italics. Tick “Set underlined words in italics” under Adjustments unless the underlining is meant to print.',
    'Underlined passages are being set in italics',
  );

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

  // KDP prints at 300 dots per inch. A picture set larger than its pixels
  // allow comes out soft; well below that it comes out visibly blocky.
  const blurry = analysis.blocks.filter((b) => b.imageMinDpi !== null && b.imageMinDpi < 300);
  const sharp = analysis.blocks.filter((b) => b.imageMinDpi !== null && b.imageMinDpi >= 300);
  if (blurry.length > 0) {
    const worst = Math.min(...blurry.map((b) => b.imageMinDpi ?? 300));
    add({
      id: 'image-resolution',
      level: worst < 150 ? 'attention' : 'check',
      title: `${blurry.length} picture${blurry.length === 1 ? ' may' : 's may'} print blurry`,
      detail:
        `KDP prints at 300 dots per inch. At the size ${blurry.length === 1 ? 'it is' : 'they are'} placed, ` +
        `${blurry.length === 1 ? 'this one comes' : 'the softest comes'} to about ${worst} dots per inch. ` +
        'Use a larger original, or make the picture smaller on the page.',
      examples: examplesOf(blurry),
    });
  } else if (sharp.length > 0) {
    add({
      id: 'image-resolution-ok',
      level: 'ready',
      title: `${sharp.length} picture${sharp.length === 1 ? ' has' : 's have'} enough detail to print sharply`,
      detail: 'Each has at least 300 pixels for every inch it takes up on the page.',
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

/**
 * Whether an ISBN's check digit holds. `length` when it is not 10 or 13
 * characters once hyphens and spaces are removed.
 */
export function checkIsbn(raw: string): 'ok' | 'checksum' | 'length' {
  const digits = raw.replace(/[-\s]/g, '').toUpperCase();
  if (/^\d{13}$/.test(digits)) {
    const sum = [...digits].reduce((acc, ch, i) => acc + Number(ch) * (i % 2 === 0 ? 1 : 3), 0);
    return sum % 10 === 0 ? 'ok' : 'checksum';
  }
  if (/^\d{9}[\dX]$/.test(digits)) {
    const sum = [...digits].reduce(
      (acc, ch, i) => acc + (ch === 'X' ? 10 : Number(ch)) * (10 - i),
      0,
    );
    return sum % 11 === 0 ? 'ok' : 'checksum';
  }
  return 'length';
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
    lineSpacing: lineSpacingMultiple(profile.bodyLineSpacing, profile.bodyLineRule, sizePt),
    wordCount: analysis.wordCount,
  });
}
