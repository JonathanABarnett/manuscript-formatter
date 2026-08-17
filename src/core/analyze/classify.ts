import type { BlockRole } from '../types.js';
import type { ParagraphFacts } from './paragraph.js';
import {
  endsLikeSentence,
  isSceneBreakText,
  labelKind,
  looksLikeCopyright,
  looksLikeHeadingCase,
  standaloneNumber,
} from './patterns.js';

export interface ClassifyContext {
  /** Nearest preceding paragraph with text, if any. */
  previous: ParagraphFacts | null;
  /** Nearest following paragraph with text, if any. */
  next: ParagraphFacts | null;
  /** A page or section break sits between `previous` and this paragraph. */
  brokenBefore: boolean;
  /** Blank paragraph(s) sit between `previous` and this paragraph. */
  blankBefore: boolean;
  /** Median word count of the manuscript's prose paragraphs. */
  medianProseWords: number;
  isFirstBlock: boolean;
}

export interface Classification {
  role: BlockRole;
  confidence: number;
  reasons: string[];
  /**
   * The paragraph names a book division outright — "Chapter 4", "Prologue", a
   * bare number, or an explicit Heading 1. Distinguishes a real chapter opening
   * from a short centred line on a title page, which merely looks like one.
   */
  structural: boolean;
}

/** Whether a paragraph's text or style declares it a book division. */
function isStructuralMarker(f: ParagraphFacts): boolean {
  if (/^heading\s*1$/i.test(f.styleName ?? '') || f.outlineLevel === 0) return true;
  if (labelKind(f.trimmed) !== null) return true;
  const num = standaloneNumber(f.trimmed);
  return num.match && num.confidence >= 0.7;
}

/**
 * Decide what a single paragraph is. Explicit Word styles are trusted first;
 * everything else falls back to weighted text-shape evidence so a manuscript
 * typed with no styles at all still gets sensible structure.
 */
export function classifyParagraph(f: ParagraphFacts, ctx: ClassifyContext): Classification {
  const result = classifyCore(f, ctx);
  return { ...result, structural: isStructuralMarker(f) };
}

function classifyCore(
  f: ParagraphFacts,
  ctx: ClassifyContext,
): Omit<Classification, 'structural'> {
  const reasons: string[] = [];

  if (f.isEmpty && !f.hasImage) {
    if (f.leadingPageBreak || f.trailingPageBreak || f.sectionBreakType) {
      return { role: 'pageBreak', confidence: 1, reasons: ['blank paragraph holding a page break'] };
    }
    return { role: 'empty', confidence: 1, reasons: ['blank paragraph'] };
  }

  if (isSceneBreakText(f.trimmed)) {
    return { role: 'sceneBreak', confidence: 0.95, reasons: [`ornament line "${f.trimmed}"`] };
  }

  if (f.hasNumbering) {
    return { role: 'listItem', confidence: 0.9, reasons: ['paragraph carries list numbering'] };
  }

  const styleName = f.styleName ?? '';
  const outline = f.outlineLevel;

  if (/^heading\s*1$/i.test(styleName) || outline === 0) {
    reasons.push(outline === 0 ? 'style is at outline level 1' : `styled "${styleName}"`);
    const kind = labelKind(f.trimmed);
    if (kind === 'part') return { role: 'partTitle', confidence: 0.9, reasons: [...reasons, 'text starts with "Part"'] };
    return { role: 'chapterTitle', confidence: 0.92, reasons };
  }
  if (/^heading\s*[2-9]$/i.test(styleName) || (outline !== null && outline >= 1)) {
    return {
      role: 'subheading',
      confidence: 0.85,
      reasons: [outline !== null ? `style is at outline level ${outline + 1}` : `styled "${styleName}"`],
    };
  }
  if (/^title$/i.test(styleName)) {
    return { role: 'frontMatterTitle', confidence: 0.8, reasons: ['styled "Title"'] };
  }
  if (/^subtitle$/i.test(styleName)) {
    return { role: 'chapterSubtitle', confidence: 0.75, reasons: ['styled "Subtitle"'] };
  }
  if (/quot|extract|epigraph|verse|poetry/i.test(styleName)) {
    return { role: 'blockQuote', confidence: 0.85, reasons: [`styled "${styleName}"`] };
  }

  // Indentation well beyond the body, with no first-line indent, reads as a
  // pulled quotation even when the manuscript uses no named style.
  if ((f.leftIndentTwips ?? 0) >= 540 && (f.firstLineIndentTwips ?? 0) <= 0 && f.wordCount >= 8) {
    return {
      role: 'blockQuote',
      confidence: 0.65,
      reasons: [`indented ${Math.round((f.leftIndentTwips ?? 0) / 14.4) / 100}" from the margin`],
    };
  }

  // A copyright notice is short, centred and unpunctuated — everything the
  // heading score rewards — but it belongs to the copyright page, never over
  // one. The structural pass promotes the whole page afterwards.
  if (looksLikeCopyright(f.trimmed) && f.wordCount <= 60) {
    return { role: 'body', confidence: 0.8, reasons: ['reads as a copyright notice'] };
  }

  const heading = scoreHeading(f, ctx);
  if (heading.score >= 0.55) {
    const kind = labelKind(f.trimmed);
    const role: BlockRole = kind === 'part' ? 'partTitle' : 'chapterTitle';
    return { role, confidence: Math.min(0.95, heading.score), reasons: heading.reasons };
  }
  if (heading.score >= 0.38 && f.wordCount <= 8 && !endsLikeSentence(f.trimmed)) {
    return {
      role: 'subheading',
      confidence: heading.score,
      reasons: [...heading.reasons, 'too weak for a chapter title, kept as a subheading'],
    };
  }

  return {
    role: 'body',
    confidence: heading.score > 0.3 ? 0.6 : 0.9,
    reasons: heading.score > 0.3 ? ['some heading-like traits, but read as prose'] : [],
  };
}

interface HeadingScore {
  score: number;
  reasons: string[];
}

/** Weighted evidence that a paragraph is a heading rather than prose. */
function scoreHeading(f: ParagraphFacts, ctx: ClassifyContext): HeadingScore {
  const reasons: string[] = [];
  let score = 0;
  const add = (points: number, why: string): void => {
    score += points;
    if (points > 0) reasons.push(why);
  };

  const kind = labelKind(f.trimmed);
  if (kind === 'chapter') add(0.5, 'starts with "Chapter"');
  else if (kind === 'part') add(0.5, 'starts with "Part"');
  else if (kind === 'frontMatter' || kind === 'backMatter') add(0.42, 'names a book division');

  const num = standaloneNumber(f.trimmed);
  if (num.match) add(0.4 * num.confidence + 0.12, 'the line is just a number');

  if (ctx.brokenBefore) add(0.24, 'starts on a new page');
  if (ctx.isFirstBlock) add(0.08, 'first paragraph in the document');

  if (f.wordCount <= 3) add(0.14, 'very short line');
  else if (f.wordCount <= 8) add(0.1, 'short line');
  else if (f.wordCount <= 12) add(0.02, 'fairly short line');
  else if (f.wordCount > 15) add(-0.6, '');
  else add(-0.2, '');

  if (f.centered) add(0.16, 'centered');
  if (f.allBold) add(0.12, 'entirely bold');
  if (f.allCaps && f.wordCount <= 12) add(0.1, 'set in capitals');
  if (!endsLikeSentence(f.trimmed)) add(0.1, 'no sentence-ending punctuation');
  else if (f.wordCount > 5) add(-0.28, '');

  if (looksLikeHeadingCase(f.trimmed) && f.wordCount <= 12) add(0.08, 'title case');
  if (ctx.blankBefore && !ctx.brokenBefore) add(0.05, 'preceded by a blank line');

  // A paragraph much shorter than the manuscript's typical prose stands out.
  if (ctx.medianProseWords > 0 && f.wordCount > 0 && f.wordCount < ctx.medianProseWords / 6) {
    add(0.06, 'far shorter than a typical paragraph');
  }

  // Dialogue and fragments are short but are not headings.
  if (/^["'“‘]/.test(f.trimmed)) add(-0.35, '');
  if (/[,;:]$/.test(f.trimmed)) add(-0.3, '');

  return { score: Math.max(0, score), reasons };
}
