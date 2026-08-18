import { DocxPackage } from '../ooxml/package.js';
import { RELTYPE } from '../ooxml/ns.js';
import { NS } from '../ooxml/ns.js';
import { attr, child, childVal, descendants, textOf } from '../ooxml/xml.js';
import type { BlockRole, DocxInput, ManuscriptAnalysis, ManuscriptBlock } from '../types.js';
import { StyleSheet, underlineOn } from './styles.js';
import { readParagraph, type ParagraphFacts } from './paragraph.js';
import { classifyParagraph, type ClassifyContext } from './classify.js';
import { looksLikeCopyright } from './patterns.js';
import { detectBookDetails } from './details.js';
import { readPixelSize } from './imageSize.js';

/** The manuscript, kept open so the composer can copy its runs verbatim. */
export interface LoadedManuscript {
  pkg: DocxPackage;
  documentDoc: Document;
  body: Element;
  styles: StyleSheet;
  /** Source XML for each block, parallel to `analysis.blocks`. */
  nodes: Element[];
  /** Paragraph facts for each block; null where the block is a table. */
  facts: (ParagraphFacts | null)[];
  analysis: ManuscriptAnalysis;
}

const PREVIEW_LIMIT = 140;

/** DrawingML measures in EMUs: 914400 per inch against 1440 twips per inch. */
const EMU_PER_TWIP = 635;

/** Widest inline or floating picture in a paragraph, in twips. */
function imageWidthTwips(p: Element): number | null {
  let widest = 0;
  for (const extent of [...descendants(p, 'extent', NS.wp), ...descendants(p, 'ext', NS.a)]) {
    const cx = Number(extent.getAttribute('cx') ?? '');
    if (Number.isFinite(cx) && cx > 0) widest = Math.max(widest, cx / EMU_PER_TWIP);
  }
  return widest > 0 ? Math.round(widest) : null;
}

/** A table's declared width, when it gives one in twips rather than a share. */
function tableWidthTwips(tbl: Element): number | null {
  const w = child(child(tbl, 'tblPr'), 'tblW');
  if (!w) return null;
  const type = attr(w, 'type');
  if (type !== 'dxa' && type !== null) return null;
  const value = Number(attr(w, 'w') ?? '');
  return Number.isFinite(value) && value > 0 ? value : null;
}

export async function analyzeManuscript(input: DocxInput): Promise<LoadedManuscript> {
  const pkg = await DocxPackage.fromBuffer(input.data, input.name);
  const documentDoc = await pkg.readXml(pkg.documentPath);
  if (!documentDoc) throw new Error('The manuscript has no readable body.');
  const body = child(documentDoc.documentElement, 'body');
  if (!body) throw new Error('The manuscript has no <w:body> element.');

  const rels = await pkg.relsFor(pkg.documentPath);
  const stylesTarget = rels.firstTargetOfType(RELTYPE.styles) ?? 'styles.xml';
  const stylesDoc = await pkg.readXml(pkg.resolveTarget(pkg.documentPath, stylesTarget));
  const styles = new StyleSheet(stylesDoc);

  const nodes = flattenBody(body);
  const facts: (ParagraphFacts | null)[] = nodes.map((node) =>
    node.localName === 'p' ? readParagraph(node, styles) : null,
  );

  const warnings: string[] = [];
  const blocks = buildBlocks(nodes, facts, warnings);
  const habits = countHabits(nodes, blocks, styles);
  await measureImages(pkg, rels, nodes, blocks);

  const analysis: ManuscriptAnalysis = {
    fileName: input.name,
    blocks,
    wordCount: blocks.reduce((sum, b) => sum + b.wordCount, 0),
    paragraphCount: blocks.filter((b) => b.kind === 'paragraph' && !b.isEmpty).length,
    chapterCount: blocks.filter((b) => b.role === 'chapterTitle').length,
    partCount: blocks.filter((b) => b.role === 'partTitle').length,
    sceneBreakCount: blocks.filter((b) => b.role === 'sceneBreak').length,
    tableCount: blocks.filter((b) => b.kind === 'table').length,
    imageCount: blocks.filter((b) => b.hasImage).length,
    footnoteCount: countFootnoteReferences(body),
    bodyStartIndex: blocks.findIndex((b) => b.role === 'chapterTitle' || b.role === 'partTitle'),
    detectedDetails: detectBookDetails(blocks),
    hasContentsPage: blocks.some((b) => !b.isEmpty && looksLikeContentsHeading(b.text)),
    ...habits,
    language: detectLanguage(body, stylesDoc, styles),
    warnings,
  };

  if (analysis.chapterCount === 0 && analysis.paragraphCount > 40) {
    warnings.push(
      'No chapter titles were detected. Mark them on the review screen, or the whole ' +
        'manuscript will be formatted as one continuous chapter.',
    );
  }
  if (analysis.tableCount > 0) {
    warnings.push(
      `${analysis.tableCount} table${analysis.tableCount === 1 ? ' is' : 's are'} copied without resizing. ` +
        `${analysis.tableCount === 1 ? 'Check its' : 'Check their'} width in Word before uploading to KDP.`,
    );
  }

  return { pkg, documentDoc, body, styles, nodes, facts, analysis };
}

/**
 * Body children as a flat list of paragraphs and tables. Content controls
 * (`w:sdt`) are unwrapped so their paragraphs are treated like any other.
 */
export function flattenBody(body: Element): Element[] {
  const out: Element[] = [];
  const visit = (parent: Element): void => {
    for (let n = parent.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const el = n as Element;
      if (el.localName === 'p' || el.localName === 'tbl') out.push(el);
      else if (el.localName === 'sdt') {
        const content = child(el, 'sdtContent');
        if (content) visit(content);
      }
    }
  };
  visit(body);
  return out;
}

function buildBlocks(
  nodes: Element[],
  facts: (ParagraphFacts | null)[],
  warnings: string[],
): ManuscriptBlock[] {
  const proseWords = facts
    .filter((f): f is ParagraphFacts => f !== null && !f.isEmpty && f.wordCount >= 12)
    .map((f) => f.wordCount)
    .sort((a, b) => a - b);
  const medianProseWords = proseWords.length ? proseWords[Math.floor(proseWords.length / 2)] : 0;

  const blocks: ManuscriptBlock[] = nodes.map((_node, index) => {
    const f = facts[index];
    if (!f) {
      const block = baseBlock(index, 'table', 'table', 1, ['table copied unchanged']);
      block.tableWidthTwips = tableWidthTwips(nodes[index]);
      return block;
    }
    const ctx = contextFor(facts, index, medianProseWords);
    const { role, confidence, reasons, structural } = classifyParagraph(f, ctx);
    return {
      index,
      kind: 'paragraph',
      autoRole: role,
      role,
      text: f.text,
      preview: preview(f.trimmed),
      styleId: f.styleId,
      styleName: f.styleName,
      outlineLevel: f.outlineLevel,
      alignment: f.alignment,
      hasPageBreakBefore: f.pageBreakBeforeProp || f.leadingPageBreak,
      isEmpty: f.isEmpty,
      wordCount: f.wordCount,
      charCount: f.trimmed.length,
      hasNumbering: f.hasNumbering,
      leftIndentTwips: f.leftIndentTwips,
      firstLineIndentTwips: f.firstLineIndentTwips,
      allBold: f.allBold,
      allItalic: f.allItalic,
      hasImage: f.hasImage,
      hasFootnote: f.hasFootnote,
      hasHyperlink: f.hasHyperlink,
      structuralMarker: structural,
      imageWidthTwips: f.hasImage ? imageWidthTwips(nodes[index]) : null,
      imageMinDpi: null,
      tableWidthTwips: null,
      confidence,
      reasons,
    };
  });

  applyStructuralPasses(blocks, warnings);
  return blocks;
}

function baseBlock(
  index: number,
  kind: ManuscriptBlock['kind'],
  role: BlockRole,
  confidence: number,
  reasons: string[],
): ManuscriptBlock {
  return {
    index,
    kind,
    autoRole: role,
    role,
    text: '',
    preview: kind === 'table' ? '[table]' : '',
    styleId: null,
    styleName: null,
    outlineLevel: null,
    alignment: null,
    hasPageBreakBefore: false,
    isEmpty: false,
    wordCount: 0,
    charCount: 0,
    hasNumbering: false,
    leftIndentTwips: null,
    firstLineIndentTwips: null,
    allBold: false,
    allItalic: false,
    hasImage: false,
    hasFootnote: false,
    hasHyperlink: false,
    structuralMarker: false,
    imageWidthTwips: null,
    imageMinDpi: null,
    tableWidthTwips: null,
    confidence,
    reasons,
  };
}

/** EMUs per inch, DrawingML's unit for picture extents. */
const EMU_PER_INCH = 914400;

/**
 * Work out how sharply each picture will print: its pixels against the size
 * it is placed at. Reads only the image headers, so a manuscript full of
 * photographs costs little more to analyse than one without.
 */
async function measureImages(
  pkg: DocxPackage,
  rels: Awaited<ReturnType<DocxPackage['relsFor']>>,
  nodes: Element[],
  blocks: ManuscriptBlock[],
): Promise<void> {
  const sizes = new Map<string, { width: number; height: number } | null>();
  const pixelSize = async (relId: string) => {
    if (sizes.has(relId)) return sizes.get(relId) ?? null;
    const rel = rels.byId(relId);
    let size: { width: number; height: number } | null = null;
    if (rel && !rel.targetMode) {
      const bytes = await pkg.readBinary(pkg.resolveTarget(pkg.documentPath, rel.target));
      if (bytes) size = readPixelSize(bytes);
    }
    sizes.set(relId, size);
    return size;
  };

  for (const block of blocks) {
    if (!block.hasImage) continue;
    let minDpi: number | null = null;
    for (const drawing of descendants(nodes[block.index], 'drawing')) {
      const extent = descendants(drawing, 'extent', NS.wp)[0] ?? null;
      const cx = Number(extent?.getAttribute('cx') ?? '');
      const cy = Number(extent?.getAttribute('cy') ?? '');
      const blip = descendants(drawing, 'blip', NS.a)[0] ?? null;
      const relId = blip ? attr(blip, 'embed', NS.r) : null;
      if (!relId || !(cx > 0) || !(cy > 0)) continue;
      const size = await pixelSize(relId);
      if (!size || size.width === 0 || size.height === 0) continue;
      const dpi = Math.min(size.width / (cx / EMU_PER_INCH), size.height / (cy / EMU_PER_INCH));
      minDpi = minDpi === null ? dpi : Math.min(minDpi, dpi);
    }
    block.imageMinDpi = minDpi === null ? null : Math.round(minDpi);
  }
}

function contextFor(
  facts: (ParagraphFacts | null)[],
  index: number,
  medianProseWords: number,
): ClassifyContext {
  let previous: ParagraphFacts | null = null;
  let brokenBefore = facts[index]?.pageBreakBeforeProp === true || facts[index]?.leadingPageBreak === true;
  let blankBefore = false;
  let sawContentBefore = false;

  for (let i = index - 1; i >= 0; i--) {
    const f = facts[i];
    if (!f) {
      sawContentBefore = true;
      break;
    }
    if (f.isEmpty) {
      blankBefore = true;
      if (f.leadingPageBreak || f.trailingPageBreak || f.sectionBreakType) brokenBefore = true;
      continue;
    }
    if (f.trailingPageBreak || f.sectionBreakType) brokenBefore = true;
    previous = f;
    sawContentBefore = true;
    break;
  }

  let next: ParagraphFacts | null = null;
  for (let i = index + 1; i < facts.length; i++) {
    const f = facts[i];
    if (!f) break;
    if (f.isEmpty) continue;
    next = f;
    break;
  }

  return {
    previous,
    next,
    brokenBefore,
    blankBefore,
    medianProseWords,
    isFirstBlock: !sawContentBefore,
  };
}

/**
 * Second pass over the classified blocks: split front matter from body matter,
 * fold a title line that follows a chapter number into a subtitle, and mark the
 * first paragraph of every section.
 */
function applyStructuralPasses(blocks: ManuscriptBlock[], warnings: string[]): void {
  const isHeading = (b: ManuscriptBlock): boolean =>
    b.role === 'chapterTitle' || b.role === 'partTitle';
  const isSkippable = (b: ManuscriptBlock): boolean =>
    b.role === 'empty' || b.role === 'pageBreak';

  // --- front matter: everything before the first structural heading ---------
  // A title page's centred lines look like headings, so the body starts at the
  // first heading that actually names a division ("Chapter One", "Prologue").
  // Where no heading does — chapters titled only by name — the first heading
  // starts the body and there is no front matter to split off.
  const firstStructural = blocks.findIndex((b) => isHeading(b) && b.structuralMarker);
  const bodyStart = firstStructural !== -1 ? firstStructural : blocks.findIndex(isHeading);
  if (bodyStart > 0) {
    let titleAssigned = false;
    for (let i = 0; i < bodyStart; i++) {
      const b = blocks[i];
      if (isSkippable(b) || b.kind === 'table') continue;
      const wasHeadingish =
        b.role === 'frontMatterTitle' || b.role === 'subheading' || b.role === 'chapterSubtitle';
      if (!titleAssigned && (wasHeadingish || (b.wordCount <= 12 && i <= 2))) {
        b.role = 'frontMatterTitle';
        b.autoRole = 'frontMatterTitle';
        b.reasons = [...b.reasons, 'first title-like line before the body matter'];
        titleAssigned = true;
      } else {
        b.role = 'frontMatter';
        b.autoRole = 'frontMatter';
        b.reasons = [...b.reasons, 'sits before the first chapter'];
      }
    }
    markCopyrightPages(blocks, bodyStart);
  } else if (bodyStart === -1) {
    for (const b of blocks) {
      if (b.role === 'frontMatterTitle') {
        b.role = 'chapterTitle';
        b.autoRole = 'chapterTitle';
        b.reasons = [...b.reasons, 'no chapters found, so treated as a chapter title'];
      }
    }
  }

  // --- "Chapter 7" followed by "The Meeting" -------------------------------
  for (let i = 0; i < blocks.length; i++) {
    // A part title followed by "Chapter One" is two real divisions, not a
    // chapter title and subtitle. Subtitle folding only begins at a chapter.
    if (blocks[i].role !== 'chapterTitle') continue;
    let crossedPageBreak = false;
    for (let j = i + 1; j < blocks.length; j++) {
      const b = blocks[j];
      if (isSkippable(b)) {
        if (b.role === 'pageBreak') crossedPageBreak = true;
        continue;
      }
      // An explicitly styled Heading 2 is a real subheading and stays one; an
      // inferred short line under a chapter number is the chapter's title.
      const inferredSubhead = b.role === 'subheading' && b.confidence < 0.7;
      const sameOpening = !crossedPageBreak && !b.hasPageBreakBefore && !b.structuralMarker;
      if (sameOpening && (isHeading(b) || inferredSubhead) && b.wordCount <= 12) {
        b.role = 'chapterSubtitle';
        b.autoRole = 'chapterSubtitle';
        b.reasons = [...b.reasons, 'follows a chapter title on the same page'];
      }
      break;
    }
  }

  // --- first paragraph after any break gets the no-indent treatment ---------
  const opensSection = new Set<BlockRole>([
    'chapterTitle',
    'partTitle',
    'chapterSubtitle',
    'subheading',
    'sceneBreak',
    'pageBreak',
  ]);
  let pendingOpen = true;
  for (const b of blocks) {
    if (b.role === 'empty') continue;
    if (b.kind === 'table') {
      pendingOpen = false;
      continue;
    }
    if (opensSection.has(b.role)) {
      pendingOpen = true;
      continue;
    }
    if (b.role === 'body' && pendingOpen) {
      b.role = 'bodyFirst';
      b.autoRole = 'bodyFirst';
      b.reasons = [...b.reasons, 'opens a section'];
    }
    if (b.role === 'body' || b.role === 'bodyFirst' || b.role === 'frontMatter') pendingOpen = false;
  }

  const lowConfidence = blocks.filter(
    (b) => b.confidence < 0.6 && b.role !== 'body' && b.role !== 'empty',
  );
  if (lowConfidence.length > 0) {
    warnings.push(
      `${lowConfidence.length} paragraph${lowConfidence.length === 1 ? ' needs' : 's need'} a quick review. ` +
        `${lowConfidence.length === 1 ? 'It is' : 'They are'} marked "check this" in the chapter and heading list.`,
    );
  }
}

/**
 * Mark the copyright page within the front matter. Only some of its lines
 * carry a recognisable marker — a bare publisher name or printing history does
 * not — so a match promotes the whole run of paragraphs between page breaks,
 * which is what shares the page in a printed book.
 */
function markCopyrightPages(blocks: ManuscriptBlock[], bodyStart: number): void {
  let pageStart = 0;
  const pages: Array<[number, number]> = [];
  for (let i = 0; i < bodyStart; i++) {
    if (blocks[i].role === 'pageBreak' || blocks[i].hasPageBreakBefore) {
      if (i > pageStart) pages.push([pageStart, i]);
      pageStart = i;
    }
  }
  pages.push([pageStart, bodyStart]);

  for (const [from, to] of pages) {
    const page = blocks.slice(from, to);
    const hasMarker = page.some(
      (b) => b.kind === 'paragraph' && !b.isEmpty && looksLikeCopyright(b.text),
    );
    if (!hasMarker) continue;
    for (const b of page) {
      if (b.role !== 'frontMatter' && b.role !== 'frontMatterTitle') continue;
      b.role = 'copyright';
      b.autoRole = 'copyright';
      b.reasons = [...b.reasons, 'part of the copyright page'];
    }
  }
}

/**
 * A heading over a contents list, in English or the other languages the
 * classifier knows, allowing a stray colon or full stop after it.
 */
export function looksLikeContentsHeading(text: string): boolean {
  const t = text.replace(/\s+/g, ' ').trim().replace(/[:.]+$/, '').trim();
  if (t.length === 0 || t.length > 40) return false;
  return /^(table of contents|contents( page)?|[íi]ndice( de contenidos?)?|tabla de contenidos?|contenido|table des mati[èe]res|sommaire|inhalt|inhaltsverzeichnis|indice|sum[áa]rio|inhoud|inhoudsopgave)$/i.test(
    t,
  );
}

/** Roles whose runs are restyled wholesale, so their underlining never prints. */
const HEADING_LIKE = new Set<BlockRole>([
  'chapterTitle',
  'partTitle',
  'chapterSubtitle',
  'subheading',
  'frontMatterTitle',
  'sceneBreak',
]);

/**
 * How much of a typed manuscript's habit the text carries: underlined runs,
 * straight quotes and apostrophes, and doubled spaces. Each is something an
 * option can tidy, and the report quotes these counts so the author knows what
 * turning it on would actually touch.
 */
function countHabits(
  nodes: Element[],
  blocks: ManuscriptBlock[],
  styles: StyleSheet,
): Pick<ManuscriptAnalysis, 'underlinedRunCount' | 'straightQuoteCount' | 'doubleSpaceCount'> {
  let underlinedRunCount = 0;
  let straightQuoteCount = 0;
  let doubleSpaceCount = 0;
  for (const block of blocks) {
    if (block.kind !== 'paragraph' || block.isEmpty) continue;
    straightQuoteCount += (block.text.match(/["']/g) ?? []).length;
    doubleSpaceCount += (block.text.match(/ {2,}/g) ?? []).length;
    if (HEADING_LIKE.has(block.role)) continue;
    for (const run of descendants(nodes[block.index], 'r')) {
      if (textOf(run).trim().length === 0) continue;
      const rPr = child(run, 'rPr');
      if (underlineOn(rPr)) {
        underlinedRunCount++;
        continue;
      }
      const styleId = childVal(rPr, 'rStyle');
      if (styleId && styles.has(styleId) && styles.resolve(styleId).underline) underlinedRunCount++;
    }
  }
  return { underlinedRunCount, straightQuoteCount, doubleSpaceCount };
}

/**
 * The language most of the text is tagged with. Word writes `w:lang` on a run
 * only when it differs from the style, so untagged text counts towards the
 * document's default (the Normal style, then the document defaults). Weighted
 * by text length so a stray foreign phrase does not decide it.
 */
function detectLanguage(body: Element, stylesDoc: Document | null, styles: StyleSheet): string | null {
  const stylesRoot = stylesDoc?.documentElement ?? null;
  const defaultLang =
    (styles.defaultParagraphStyleId
      ? styleLang(stylesRoot, styles.defaultParagraphStyleId)
      : null) ??
    attr(child(child(child(child(stylesRoot, 'docDefaults'), 'rPrDefault'), 'rPr'), 'lang'), 'val');

  const weights = new Map<string, number>();
  for (const run of descendants(body, 'r')) {
    const length = textOf(run).length;
    if (length === 0) continue;
    const lang = attr(child(child(run, 'rPr'), 'lang'), 'val') ?? defaultLang;
    if (!lang) continue;
    weights.set(lang, (weights.get(lang) ?? 0) + length);
  }
  let best: string | null = null;
  let bestWeight = 0;
  for (const [lang, weight] of weights) {
    if (weight > bestWeight) {
      best = lang;
      bestWeight = weight;
    }
  }
  return best;
}

function styleLang(stylesRoot: Element | null, styleId: string): string | null {
  if (!stylesRoot) return null;
  for (let n = stylesRoot.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    if (el.localName !== 'style' || attr(el, 'styleId') !== styleId) continue;
    return attr(child(child(el, 'rPr'), 'lang'), 'val');
  }
  return null;
}

function preview(text: string): string {
  const clean = text.replace(/\s+/g, ' ').trim();
  return clean.length > PREVIEW_LIMIT ? `${clean.slice(0, PREVIEW_LIMIT - 1)}…` : clean;
}

function countFootnoteReferences(body: Element): number {
  return (
    descendants(body, 'footnoteReference').length + descendants(body, 'endnoteReference').length
  );
}
