import { DocxPackage } from '../ooxml/package.js';
import { NS, RELTYPE } from '../ooxml/ns.js';
import {
  attr,
  child,
  childVal,
  children,
  descendants,
  numAttr,
  textOf,
  toggleOn,
} from '../ooxml/xml.js';
import type { DocxInput, PageSetup, ReferenceProfile, StyleRole } from '../types.js';
import { STYLE_ROLES } from '../types.js';
import { StyleSheet } from './styles.js';
import { readParagraph, type ParagraphFacts } from './paragraph.js';
import { isSceneBreakText, labelKind, standaloneNumber } from './patterns.js';
import { describePageSize } from './pageSize.js';

/** The reference document, kept open so the composer can reuse its XML. */
export interface LoadedReference {
  pkg: DocxPackage;
  documentDoc: Document;
  body: Element;
  styles: StyleSheet;
  /** Section properties in document order; the last one is body-level. */
  sectPrs: Element[];
  bodySectPr: Element | null;
  /** Section for front matter when the reference is multi-section. */
  frontSectPr: Element | null;
  profile: ReferenceProfile;
}

export async function analyzeReference(input: DocxInput): Promise<LoadedReference> {
  const pkg = await DocxPackage.fromBuffer(input.data, input.name);
  const documentDoc = await pkg.readXml(pkg.documentPath);
  if (!documentDoc) throw new Error('The reference document has no readable body.');
  const body = child(documentDoc.documentElement, 'body');
  if (!body) throw new Error('The reference document has no <w:body> element.');

  const rels = await pkg.relsFor(pkg.documentPath);
  const stylesTarget = rels.firstTargetOfType(RELTYPE.styles) ?? 'styles.xml';
  const stylesDoc = await pkg.readXml(pkg.resolveTarget(pkg.documentPath, stylesTarget));
  const styles = new StyleSheet(stylesDoc);

  const settingsTarget = rels.firstTargetOfType(RELTYPE.settings);
  const settingsDoc = settingsTarget
    ? await pkg.readXml(pkg.resolveTarget(pkg.documentPath, settingsTarget))
    : null;

  const sectPrs = collectSectPrs(body);
  const bodySectPr = sectPrs.length > 0 ? sectPrs[sectPrs.length - 1] : null;
  const frontSectPr = sectPrs.length > 1 ? sectPrs[0] : null;

  const facts = children(body, 'p').map((p) => readParagraph(p, styles));
  const warnings: string[] = [];

  const pageSetup = readPageSetup(bodySectPr, settingsDoc, warnings);
  const usage = countStyleUsage(facts, styles.defaultParagraphStyleId);
  const detection = detectRoles(styles, facts, usage, warnings);

  const sectionLayout = readSectionLayout(
    body,
    facts,
    sectPrs,
    detection.roleStyles,
    styles.defaultParagraphStyleId,
  );
  const headerFooter = await inspectHeadersAndFooters(pkg, rels, sectPrs);
  const hasFootnotes = await documentHasFootnotes(pkg, rels);

  if (headerFooter.text.length > 0) {
    warnings.push(
      'Words in the design file\'s headers or footers will also appear in the new book. ' +
        'Make sure the names shown below belong to this book.',
    );
  }

  if (sectPrs.length > 1) {
    warnings.push(
      `The design file uses ${sectPrs.length} different page-layout sections. The first is used ` +
        'for the opening pages and the last is used for the main book.',
    );
  }
  if (!detection.roleStyles.chapterTitle) {
    warnings.push(
      'No chapter-title formatting was found in the design file. Chapter titles will use centered, ' +
        'bold body text unless you choose a different look below.',
    );
  }
  if (facts.filter((f) => !f.isEmpty).length < 5) {
    warnings.push(
      'The design file contains very little sample text, so some formatting choices were guessed ' +
        'from their names. Check the sample pages below.',
    );
  }

  const bodyStyleId = detection.roleStyles.body;
  const bodyProps = bodyStyleId ? styles.resolve(bodyStyleId) : null;

  const profile: ReferenceProfile = {
    fileName: input.name,
    pageSetup,
    pageSizeLabel: describePageSize(pageSetup.widthTwips, pageSetup.heightTwips),
    sectionCount: sectPrs.length,
    styles: styles.toStyleInfos(usage),
    roleStyles: detection.roleStyles,
    roleEvidence: detection.roleEvidence,
    chapterStartsOnNewPage: detection.chapterStartsOnNewPage,
    chapterStartsOnOddPage: sectPrs.some(
      (s) => (child(s, 'type') && attr(child(s, 'type'), 'val')) === 'oddPage',
    ),
    chapterTitleBlanksBefore: detection.chapterTitleBlanksBefore,
    chapterTitleBlanksAfter: detection.chapterTitleBlanksAfter,
    roleVAlign: sectionLayout.roleVAlign,
    bodySectionIndex: sectionLayout.bodySectionIndex,
    usesFirstParagraphNoIndent: detection.usesFirstParagraphNoIndent,
    bodyFirstLineIndentTwips: bodyProps?.firstLineIndentTwips ?? null,
    hasHeaders: headerFooter.hasHeaders,
    hasFooters: headerFooter.hasFooters,
    hasPageNumbers: headerFooter.hasPageNumbers,
    headerFooterText: headerFooter.text,
    hasFootnotes,
    hyphenates: toggleOn(settingsDoc?.documentElement ?? null, 'autoHyphenation'),
    bodyJustified: bodyProps?.alignment === 'both',
    defaultParagraphStyleId: styles.defaultParagraphStyleId,
    bodyFontName: bodyProps?.fontName ?? null,
    bodyFontSizePt: bodyProps?.fontSizePt ?? null,
    bodyLineSpacing: bodyProps?.lineSpacing ?? null,
    bodyLineRule: bodyProps?.lineRule ?? null,
    warnings,
  };

  return { pkg, documentDoc, body, styles, sectPrs, bodySectPr, frontSectPr, profile };
}

/**
 * Which section each body paragraph belongs to. A paragraph carrying a
 * `sectPr` is the *last* of its section, so the counter advances after it.
 */
function sectionIndexPerParagraph(body: Element): number[] {
  let section = 0;
  return children(body, 'p').map((p) => {
    const here = section;
    if (child(child(p, 'pPr'), 'sectPr')) section++;
    return here;
  });
}

/**
 * Vertical placement per role, read from the section each role's paragraphs
 * live in, plus the section the body matter starts in.
 */
function readSectionLayout(
  body: Element,
  facts: ParagraphFacts[],
  sectPrs: Element[],
  roleStyles: Record<StyleRole, string | null>,
  defaultStyleId: string | null,
): { roleVAlign: Partial<Record<StyleRole, string>>; bodySectionIndex: number } {
  const sectionOf = sectionIndexPerParagraph(body);
  const vAlignOf = (index: number): string | null =>
    index < sectPrs.length ? childVal(sectPrs[index], 'vAlign') : null;

  const roleVAlign: Partial<Record<StyleRole, string>> = {};
  const interesting: StyleRole[] = ['frontMatterTitle', 'copyright', 'frontMatter', 'partTitle'];
  for (const role of interesting) {
    const styleId = roleStyles[role];
    if (!styleId) continue;
    const counts = new Map<string, number>();
    facts.forEach((f, i) => {
      if (f.isEmpty || (f.styleId ?? defaultStyleId) !== styleId) return;
      const align = vAlignOf(sectionOf[i] ?? 0);
      if (align) counts.set(align, (counts.get(align) ?? 0) + 1);
    });
    let best: string | null = null;
    let bestCount = 0;
    for (const [align, count] of counts) {
      if (count > bestCount) {
        best = align;
        bestCount = count;
      }
    }
    if (best) roleVAlign[role] = best;
  }

  return { roleVAlign, bodySectionIndex: findBodySection(sectPrs) };
}

const ROMAN_OR_LETTER = /^(lower|upper)(Roman|Letter)$/;

/**
 * The section the body matter runs in. Books number their front matter in
 * roman numerals and restart at 1 in arabic where the body begins, so the
 * first section that restarts in arabic is the body's.
 *
 * Style is no help here: a template uses its chapter-title style for
 * "Dedication" and "Contents" as well, which are front matter.
 */
function findBodySection(sectPrs: Element[]): number {
  const last = Math.max(0, sectPrs.length - 1);
  for (let i = 0; i < sectPrs.length; i++) {
    const pgNumType = child(sectPrs[i], 'pgNumType');
    if (!pgNumType) continue;
    if (attr(pgNumType, 'start') === null) continue;
    const format = attr(pgNumType, 'fmt');
    if (format === null || !ROMAN_OR_LETTER.test(format)) return i;
  }
  return last;
}

/** Section properties, both paragraph-level and the final body-level one. */
export function collectSectPrs(body: Element): Element[] {
  const out: Element[] = [];
  for (const p of children(body, 'p')) {
    const sectPr = child(child(p, 'pPr'), 'sectPr');
    if (sectPr) out.push(sectPr);
  }
  const final = child(body, 'sectPr');
  if (final) out.push(final);
  return out;
}

function readPageSetup(
  sectPr: Element | null,
  settingsDoc: Document | null,
  warnings: string[],
): PageSetup {
  const pgSz = child(sectPr, 'pgSz');
  const pgMar = child(sectPr, 'pgMar');
  const width = numAttr(pgSz, 'w');
  const height = numAttr(pgSz, 'h');
  if (width === null || height === null) {
    warnings.push('The design file has no page size. The new book will use US Letter size.');
  }
  const settingsRoot = settingsDoc?.documentElement ?? null;
  const cols = child(sectPr, 'cols');

  return {
    widthTwips: width ?? 12240,
    heightTwips: height ?? 15840,
    orientation: attr(pgSz, 'orient') === 'landscape' ? 'landscape' : 'portrait',
    margins: {
      top: numAttr(pgMar, 'top') ?? 1440,
      right: numAttr(pgMar, 'right') ?? 1440,
      bottom: numAttr(pgMar, 'bottom') ?? 1440,
      left: numAttr(pgMar, 'left') ?? 1440,
      header: numAttr(pgMar, 'header') ?? 720,
      footer: numAttr(pgMar, 'footer') ?? 720,
      gutter: numAttr(pgMar, 'gutter') ?? 0,
    },
    mirrorMargins: child(settingsRoot, 'mirrorMargins') !== null,
    differentFirstPage: child(sectPr, 'titlePg') !== null,
    differentOddEven: child(settingsRoot, 'evenAndOddHeaders') !== null,
    sectionBreakType: attr(child(sectPr, 'type'), 'val'),
    columns: numAttr(cols, 'num') ?? 1,
  };
}

function countStyleUsage(
  facts: ParagraphFacts[],
  defaultStyleId: string | null,
): Map<string, number> {
  const usage = new Map<string, number>();
  for (const f of facts) {
    if (f.isEmpty) continue;
    const id = f.styleId ?? defaultStyleId;
    if (!id) continue;
    usage.set(id, (usage.get(id) ?? 0) + 1);
  }
  return usage;
}

interface RoleDetection {
  roleStyles: Record<StyleRole, string | null>;
  roleEvidence: Partial<Record<StyleRole, string>>;
  chapterStartsOnNewPage: boolean;
  usesFirstParagraphNoIndent: boolean;
  chapterTitleBlanksBefore: number;
  chapterTitleBlanksAfter: number;
}

function detectRoles(
  styles: StyleSheet,
  facts: ParagraphFacts[],
  usage: Map<string, number>,
  warnings: string[],
): RoleDetection {
  const roleStyles = Object.fromEntries(STYLE_ROLES.map((r) => [r, null])) as Record<
    StyleRole,
    string | null
  >;
  const evidence: Partial<Record<StyleRole, string>> = {};
  const effectiveId = (f: ParagraphFacts): string | null =>
    f.styleId ?? styles.defaultParagraphStyleId;

  /** Styles already given a role, so no two roles resolve to the same style. */
  const claimed = new Set<string>();
  const set = (role: StyleRole, id: string | null, why: string): void => {
    if (!id || roleStyles[role]) return;
    roleStyles[role] = id;
    evidence[role] = why;
    claimed.add(id);
  };

  // --- first paragraph after a break, by name ------------------------------
  // Detected before the body style so that a template's "First Paragraph"
  // style is not mistaken for the body style it opens.
  set(
    'bodyFirst',
    firstIdByPattern(
      styles,
      /first\s*para|body\s*text\s*first|text\s*first|no\s*indent|noindent|chapter\s*first|opening/i,
    ),
    'the design file names it as a first paragraph without an indent',
  );

  // --- body: the style carrying the bulk of the reference's prose ----------
  const prose = facts.filter((f) => !f.isEmpty && f.wordCount >= 25 && !f.hasNumbering);
  const proseUsage = new Map<string, number>();
  for (const f of prose) {
    const id = effectiveId(f);
    if (id) proseUsage.set(id, (proseUsage.get(id) ?? 0) + 1);
  }
  if (roleStyles.bodyFirst && proseUsage.size > 1) proseUsage.delete(roleStyles.bodyFirst);
  const topProse = pickBodyCandidate(styles, proseUsage);
  if (topProse) {
    set(
      'body',
      topProse[0],
      `used by ${topProse[1]} regular paragraph${topProse[1] === 1 ? '' : 's'} in the design file`,
    );
  }
  if (!roleStyles.body) {
    const byName = firstIdByNames(styles, ['Body Text', 'Body', 'Book Body', 'Text', 'Normal']);
    if (byName) set('body', byName, `matched the style name "${styles.nameOf(byName)}"`);
  }
  if (!roleStyles.body && styles.defaultParagraphStyleId) {
    set('body', styles.defaultParagraphStyleId, "the document's default paragraph style");
  }
  if (!roleStyles.body) {
    const first = styles.paragraphStyleIds()[0];
    if (first) set('body', first, 'the only paragraph style available');
    else warnings.push('The reference defines no paragraph styles.');
  }

  // --- title page: claimed first so a book title cannot become a chapter ---
  set(
    'frontMatterTitle',
    firstIdByPattern(styles, /book\s*title|half\s*title|front\s*matter\s*title|^title$/i, usage, claimed),
    'style name mentions a title page',
  );
  set(
    'copyright',
    firstIdByPattern(styles, /copyright|^colophon$/i, usage, claimed),
    'style name mentions the copyright page',
  );

  // --- headings ------------------------------------------------------------
  // A style whose *name* says "chapter title" is far stronger evidence than a
  // style that happened to be used once on something heading-shaped, so the
  // name match is tried first.
  set(
    'chapterTitle',
    firstIdByPattern(styles, /chapter\s*(title|head|name)|^chapter$/i, usage, claimed),
    'style name mentions "chapter"',
  );
  const headingFacts = facts.filter((f) => isHeadingLike(f));
  const chapterUsage = new Map<string, number>();
  for (const f of headingFacts) {
    if (f.outlineLevel !== null && f.outlineLevel > 0) continue;
    const id = effectiveId(f);
    if (id && id !== roleStyles.body && !claimed.has(id)) {
      chapterUsage.set(id, (chapterUsage.get(id) ?? 0) + 1);
    }
  }
  const topChapter = topEntry(chapterUsage);
  if (topChapter) {
    set(
      'chapterTitle',
      topChapter[0],
      `used by ${topChapter[1]} main heading${topChapter[1] === 1 ? '' : 's'} in the design file`,
    );
  }
  const outline0 = styles
    .paragraphStyleIds()
    .find((id) => styles.resolve(id).outlineLevel === 0 && id !== roleStyles.body && !claimed.has(id));
  set('chapterTitle', outline0 ?? null, 'the only style at outline level 1');
  set(
    'chapterTitle',
    firstIdByNames(styles, ['Heading 1']),
    'fell back to the built-in Heading 1 style',
  );

  set(
    'partTitle',
    firstIdByPattern(styles, /part\s*(title|head)|^part$/i, usage, claimed),
    'style name mentions "part"',
  );
  set(
    'chapterSubtitle',
    // Not "book subtitle": that belongs to the title page and is typically set
    // several times larger than anything inside a chapter. Word's built-in
    // "Subtitle" is the same trap, so it only counts if the template uses it.
    firstIdByPattern(styles, /chapter\s*subtitle|^subtitle$/i, usage, claimed, true),
    'style name mentions "subtitle"',
  );

  const outline1 = styles
    .paragraphStyleIds()
    .find((id) => styles.resolve(id).outlineLevel === 1 && !claimed.has(id));
  set(
    'subheading',
    firstIdByPattern(styles, /sub\s*head|heading\s*2/i, usage, claimed),
    'style name mentions "subhead"',
  );
  set('subheading', outline1 ?? null, 'style sits at outline level 2');
  set('subheading', firstIdByNames(styles, ['Heading 2']), 'fell back to the built-in Heading 2 style');

  // --- scene break --------------------------------------------------------
  const sceneUsage = new Map<string, number>();
  for (const f of facts) {
    if (f.isEmpty || !isSceneBreakText(f.trimmed)) continue;
    const id = effectiveId(f);
    if (id && id !== roleStyles.body) sceneUsage.set(id, (sceneUsage.get(id) ?? 0) + 1);
  }
  const topScene = topEntry(sceneUsage);
  if (topScene) {
    set(
      'sceneBreak',
      topScene[0],
      `used by ${topScene[1]} scene-break mark${topScene[1] === 1 ? '' : 's'} in the design file`,
    );
  }
  set(
    'sceneBreak',
    firstIdByPattern(
      styles,
      /scene\s*break|dinkus|ornament|separator|asterisk|space\s*break/i,
      usage,
      claimed,
    ),
    'style name mentions a scene break',
  );

  // --- block quote --------------------------------------------------------
  set(
    'blockQuote',
    firstIdByPattern(
      styles,
      /block\s*quote|^quote$|^quotation$|extract|epigraph|verse|poetry/i,
      usage,
      claimed,
    ),
    'style name mentions a quotation',
  );
  if (!roleStyles.blockQuote && roleStyles.body) {
    const bodyLeft = styles.resolve(roleStyles.body).leftIndentTwips ?? 0;
    const indented = [...usage.keys()].find((id) => {
      if (id === roleStyles.body) return false;
      const props = styles.resolve(id);
      return (props.leftIndentTwips ?? 0) >= bodyLeft + 180 && (props.outlineLevel ?? 9) > 1;
    });
    set('blockQuote', indented ?? null, 'indented further than body text in the reference');
  }

  set(
    'listItem',
    firstIdByPattern(styles, /list\s*paragraph|list\s*bullet|list\s*number|^list$/i, usage, claimed),
    'style name mentions a list',
  );

  // --- first paragraph after a break, by position in the reference ---------
  if (!roleStyles.bodyFirst && roleStyles.chapterTitle) {
    const afterHeading = new Map<string, number>();
    for (let i = 0; i < facts.length - 1; i++) {
      if (effectiveId(facts[i]) !== roleStyles.chapterTitle) continue;
      for (let j = i + 1; j < facts.length; j++) {
        if (facts[j].isEmpty) continue;
        const id = effectiveId(facts[j]);
        if (id && id !== roleStyles.body) afterHeading.set(id, (afterHeading.get(id) ?? 0) + 1);
        break;
      }
    }
    const top = topEntry(afterHeading);
    if (top && top[1] >= 2) {
      set('bodyFirst', top[0], `follows a chapter title ${top[1]} times in the reference`);
    }
  }

  // --- front matter -------------------------------------------------------
  set(
    'frontMatter',
    firstIdByPattern(styles, /front\s*matter|^no\s*indent$|dedication/i, usage, claimed),
    'style name mentions front matter',
  );
  if (!roleStyles.frontMatter) {
    set(
      'frontMatter',
      roleStyles.bodyFirst ?? roleStyles.body,
      'uses the same formatting as regular book text',
    );
  }
  if (!roleStyles.copyright) {
    set(
      'copyright',
      roleStyles.frontMatter,
      'uses the same formatting as the other opening-page text',
    );
  }

  // --- layout behaviour ---------------------------------------------------
  const chapterStyle = roleStyles.chapterTitle;
  let chapterStartsOnNewPage = false;
  if (chapterStyle && styles.resolve(chapterStyle).pageBreakBefore) {
    chapterStartsOnNewPage = true;
  } else {
    let starts = 0;
    let total = 0;
    for (let i = 0; i < facts.length; i++) {
      const f = facts[i];
      if (!chapterStyle || effectiveId(f) !== chapterStyle || f.isEmpty) continue;
      total++;
      const prev = previousNonEmpty(facts, i);
      if (
        f.pageBreakBeforeProp ||
        f.leadingPageBreak ||
        prev === null ||
        prev.trailingPageBreak ||
        prev.leadingPageBreak ||
        prev.sectionBreakType !== null
      ) {
        starts++;
      }
    }
    // With no evidence either way, a new page per chapter is the book default.
    chapterStartsOnNewPage = total === 0 ? true : starts / total >= 0.6;
  }

  const bodyIndent = roleStyles.body ? styles.resolve(roleStyles.body).firstLineIndentTwips ?? 0 : 0;
  const firstIndent = roleStyles.bodyFirst
    ? styles.resolve(roleStyles.bodyFirst).firstLineIndentTwips ?? 0
    : 0;
  const usesFirstParagraphNoIndent =
    bodyIndent > 0 && (roleStyles.bodyFirst === null || firstIndent <= 0);

  const sink = measureChapterSink(facts, roleStyles.chapterTitle, effectiveId);

  return {
    roleStyles,
    roleEvidence: evidence,
    chapterStartsOnNewPage,
    usesFirstParagraphNoIndent,
    chapterTitleBlanksBefore: sink.before,
    chapterTitleBlanksAfter: sink.after,
  };
}

function isHeadingLike(f: ParagraphFacts): boolean {
  if (f.isEmpty) return false;
  if (f.outlineLevel !== null && f.outlineLevel <= 2) return true;
  if (f.styleName && /^heading\s*[1-3]$/i.test(f.styleName)) return true;
  if (f.wordCount > 12) return false;
  // "Dedication", "Contents" and "Acknowledgments" head their own pages in the
  // same style a chapter uses, so they count as headings too.
  if (labelKind(f.trimmed) !== null) return true;
  return standaloneNumber(f.trimmed).match && (f.centered || f.allBold);
}

/**
 * How far a chapter opening is sunk down the page, measured in blank
 * paragraphs of the chapter-title style either side of the title. Templates
 * commonly use five or six to push the title a third of the way down; without
 * reproducing them the title lands hard against the top margin.
 */
function measureChapterSink(
  facts: ParagraphFacts[],
  chapterStyleId: string | null,
  effectiveId: (f: ParagraphFacts) => string | null,
): { before: number; after: number } {
  if (!chapterStyleId) return { before: 0, after: 0 };
  const befores: number[] = [];
  const afters: number[] = [];

  for (let i = 0; i < facts.length; i++) {
    const f = facts[i];
    if (f.isEmpty || effectiveId(f) !== chapterStyleId) continue;

    let before = 0;
    for (let j = i - 1; j >= 0 && facts[j].isEmpty && effectiveId(facts[j]) === chapterStyleId; j--) {
      before++;
    }
    let after = 0;
    for (
      let j = i + 1;
      j < facts.length && facts[j].isEmpty && effectiveId(facts[j]) === chapterStyleId;
      j++
    ) {
      after++;
    }
    befores.push(before);
    afters.push(after);
  }

  return { before: mode(befores), after: mode(afters) };
}

/** Most common value, preferring the larger on a tie. Zero when empty. */
function mode(values: number[]): number {
  if (values.length === 0) return 0;
  const counts = new Map<number, number>();
  for (const v of values) counts.set(v, (counts.get(v) ?? 0) + 1);
  let best = 0;
  let bestCount = 0;
  for (const [value, count] of counts) {
    if (count > bestCount || (count === bestCount && value > best)) {
      best = value;
      bestCount = count;
    }
  }
  return best;
}

function previousNonEmpty(facts: ParagraphFacts[], index: number): ParagraphFacts | null {
  for (let i = index - 1; i >= 0; i--) {
    if (!facts[i].isEmpty) return facts[i];
    // A blank paragraph can still carry the page break that starts a chapter.
    if (facts[i].leadingPageBreak || facts[i].trailingPageBreak || facts[i].sectionBreakType) {
      return facts[i];
    }
  }
  return null;
}

/**
 * The prose style that carries the book's running text. Where two styles are
 * used about equally, the `w:next` chain decides: a style whose "next
 * paragraph" is another candidate is an opener, so its target is the body.
 */
function pickBodyCandidate(
  styles: StyleSheet,
  usage: Map<string, number>,
): [string, number] | null {
  const top = topEntry(usage);
  if (!top) return null;
  const contenders = [...usage].filter(([, count]) => count >= top[1] - 1);
  if (contenders.length < 2) return top;
  for (const [id] of contenders) {
    const next = styles.nextStyleOf(id);
    if (next && next !== id && usage.has(next)) {
      return [next, usage.get(next) ?? 0];
    }
  }
  return top;
}

function topEntry(map: Map<string, number>): [string, number] | null {
  let best: [string, number] | null = null;
  for (const entry of map) {
    if (!best || entry[1] > best[1]) best = entry;
  }
  return best;
}

function firstIdByNames(styles: StyleSheet, names: string[]): string | null {
  for (const name of names) {
    const id = styles.idByName(name);
    if (id && styles.typeOf(id) === 'paragraph') return id;
  }
  return null;
}

/**
 * Best style whose id or name matches `pattern`, preferring the one the
 * reference actually uses. Word documents carry dozens of unused built-in
 * styles, so "heading 2" would otherwise beat a template's own "Subhead".
 * `exclude` keeps a style already claimed by another role from being reused.
 */
function firstIdByPattern(
  styles: StyleSheet,
  pattern: RegExp,
  usage: Map<string, number> = new Map(),
  exclude: ReadonlySet<string> = new Set(),
  /** Ignore styles the reference never uses. For roles where guessing wrong
   *  is visually severe, an unused built-in is worse than no match at all. */
  requireUsed = false,
): string | null {
  const matches = styles
    .findByPattern(pattern)
    .filter((id) => !exclude.has(id) && (!requireUsed || (usage.get(id) ?? 0) > 0));
  if (matches.length === 0) return null;
  return matches.sort((a, b) => (usage.get(b) ?? 0) - (usage.get(a) ?? 0))[0];
}

async function inspectHeadersAndFooters(
  pkg: DocxPackage,
  rels: Awaited<ReturnType<DocxPackage['relsFor']>>,
  sectPrs: Element[],
): Promise<{
  hasHeaders: boolean;
  hasFooters: boolean;
  hasPageNumbers: boolean;
  text: string[];
}> {
  const referenced = new Set<string>();
  for (const sectPr of sectPrs) {
    for (const kind of ['headerReference', 'footerReference'] as const) {
      for (const ref of children(sectPr, kind)) {
        const id = attr(ref, 'id', NS.r);
        if (id) referenced.add(id);
      }
    }
  }
  let hasHeaders = false;
  let hasFooters = false;
  let hasPageNumbers = false;
  const visibleText = new Set<string>();

  for (const id of referenced) {
    const rel = rels.byId(id);
    if (!rel) continue;
    if (rel.type === RELTYPE.header) hasHeaders = true;
    if (rel.type === RELTYPE.footer) hasFooters = true;
    const partPath = pkg.resolveTarget(pkg.documentPath, rel.target);
    const doc = await pkg.readXml(partPath);
    if (!doc) continue;
    if (containsPageField(doc.documentElement)) hasPageNumbers = true;
    const words = textOf(doc.documentElement).replace(/\s+/g, ' ').trim();
    // A cached page-field result such as "1" is not template wording the
    // author needs to replace. Keep any surrounding title or author text.
    if (words && !/^(?:page\s*)?\d+$/i.test(words)) visibleText.add(words);
  }
  return { hasHeaders, hasFooters, hasPageNumbers, text: [...visibleText] };
}

/** Detect a PAGE field, written either as `fldSimple` or an `instrText` run. */
function containsPageField(root: Element | null): boolean {
  if (!root) return false;
  for (const el of descendants(root, 'fldSimple')) {
    if (/\bPAGE\b/i.test(attr(el, 'instr') ?? '')) return true;
  }
  for (const el of descendants(root, 'instrText')) {
    if (/\bPAGE\b/i.test(textOf(el.parentNode) || el.textContent || '')) return true;
  }
  return false;
}

async function documentHasFootnotes(
  pkg: DocxPackage,
  rels: Awaited<ReturnType<DocxPackage['relsFor']>>,
): Promise<boolean> {
  const target = rels.firstTargetOfType(RELTYPE.footnotes);
  if (!target) return false;
  const doc = await pkg.readXml(pkg.resolveTarget(pkg.documentPath, target));
  if (!doc) return false;
  return children(doc.documentElement, 'footnote').some((fn) => {
    const id = Number(attr(fn, 'id') ?? '0');
    return Number.isFinite(id) && id > 0;
  });
}
