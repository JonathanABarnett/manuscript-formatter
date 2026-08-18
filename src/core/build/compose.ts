import { NS, RELTYPE } from '../ooxml/ns.js';
import type { DocxPackage, Relationships } from '../ooxml/package.js';
import { attr, child, children, clearChildren, descendants, importNode, wEl } from '../ooxml/xml.js';
import type { BlockRole, FormatOptions, FormatStats, ManuscriptBlock, StyleRole } from '../types.js';
import type { LoadedReference } from '../analyze/reference.js';
import { collectSectPrs } from '../analyze/reference.js';
import type { LoadedManuscript } from '../analyze/manuscript.js';
import { styleForRole } from '../roles.js';
import {
  buildBackMatter,
  buildContentsPage,
  buildFrontMatter,
  needsFieldUpdate,
} from './matter.js';
import { applyDetailsToRunningHeads } from './headers.js';
import { ResourceMigrator } from './resources.js';
import { NumberingMerger } from './numbering.js';
import { NoteMerger } from './notes.js';
import { TextTransformer, copyParagraphContent } from './runs.js';

export interface ComposeResult {
  pkg: DocxPackage;
  stats: FormatStats;
  warnings: string[];
}

/** Roles that start a fresh page/section when the option calls for it. */
const CHAPTER_ROLES = new Set<BlockRole>(['chapterTitle', 'partTitle']);
const FRONT_MATTER_ROLES = new Set<BlockRole>(['frontMatterTitle', 'frontMatter', 'copyright']);
/** Roles whose appearance the reference's own heading style should own. */
const HEADING_ROLES = new Set<BlockRole>([
  'chapterTitle',
  'partTitle',
  'chapterSubtitle',
  'subheading',
  'frontMatterTitle',
  'sceneBreak',
]);

/**
 * Build the output document: the reference package with its body replaced by
 * the manuscript's content, mapped onto the reference's styles.
 */
export async function composeDocument(
  reference: LoadedReference,
  manuscript: LoadedManuscript,
  options: FormatOptions,
): Promise<ComposeResult> {
  const warnings: string[] = [];
  const out = await reference.pkg.clone();

  const outDoc = await out.readXml(out.documentPath);
  const outBody = outDoc ? child(outDoc.documentElement, 'body') : null;
  if (!outDoc || !outBody) throw new Error('The reference document could not be rebuilt.');

  // Section properties must be captured before the body is emptied.
  const outSectPrs = collectSectPrs(outBody);
  // Not necessarily the last: a template gives each chapter its own section,
  // and it is the first body section that restarts the page numbering.
  const bodySectPr =
    outSectPrs[reference.profile.bodySectionIndex] ?? outSectPrs[outSectPrs.length - 1] ?? null;
  const frontSectPrSource = outSectPrs.length > 1 ? outSectPrs[0] : null;
  const bodySectPrClone = bodySectPr ? (bodySectPr.cloneNode(true) as Element) : null;
  const frontSectPrClone = frontSectPrSource
    ? (frontSectPrSource.cloneNode(true) as Element)
    : null;
  clearChildren(outBody);

  const outRels = await out.relsFor(out.documentPath);
  const msRels = await manuscript.pkg.relsFor(manuscript.pkg.documentPath);
  const migrator = new ResourceMigrator(
    manuscript.pkg,
    manuscript.pkg.documentPath,
    msRels,
    out,
    out.documentPath,
    outRels,
  );

  const msNumberingTarget = msRels.firstTargetOfType(RELTYPE.numbering);
  const numbering = new NumberingMerger(
    msNumberingTarget
      ? await manuscript.pkg.readXml(
          manuscript.pkg.resolveTarget(manuscript.pkg.documentPath, msNumberingTarget),
        )
      : null,
    out,
    outRels,
  );

  const footnotes = new NoteMerger('footnote', manuscript.pkg, msRels, out, outRels, reference.styles);
  const endnotes = new NoteMerger('endnote', manuscript.pkg, msRels, out, outRels, reference.styles);

  const roleStyles = resolveRoleStyles(reference, options, warnings);
  const transformer = new TextTransformer(options.smartTypography, options.collapseMultipleSpaces);
  const copyOptions = {
    targetDoc: outDoc,
    manuscriptStyles: manuscript.styles,
    referenceStyles: reference.styles,
    keepEmphasis: options.keepEmphasis,
    removeManualIndents: options.removeManualIndents,
    transformer,
  };

  const bodyIndent = roleStyles.body
    ? reference.styles.resolve(roleStyles.body).firstLineIndentTwips ?? 0
    : 0;

  // The reviewer may tune the chapter sink; otherwise follow the template.
  const chapterBlanksBefore = Math.max(
    0,
    options.chapterSpaceBefore ?? reference.profile.chapterTitleBlanksBefore,
  );
  const chapterBlanksAfter = Math.max(
    0,
    options.chapterSpaceAfter ?? reference.profile.chapterTitleBlanksAfter,
  );

  const stats: FormatStats = {
    paragraphsWritten: 0,
    chapters: 0,
    parts: 0,
    sceneBreaks: 0,
    tables: 0,
    imagesCopied: 0,
    footnotesCopied: 0,
    blanksRemoved: 0,
    runningHeadsUpdated: 0,
    wordCount: 0,
  };

  // Generated opening pages come first, before anything from the manuscript.
  const matterContext = {
    doc: outDoc,
    roles: roleStyles,
    details: options.bookDetails,
    sections: options.extraSections,
  };
  const frontTemplateSectPr = frontSectPrClone ?? bodySectPrClone;
  const generatedFront = buildFrontMatter(matterContext);
  let contentsPage = buildContentsPage(matterContext);
  let lastParagraph: Element | null = null;

  /**
   * Close a run of front-matter paragraphs with its own section, so the page
   * can sit where the template puts it — a copyright notice at the foot of
   * the page, a title page centred — rather than all falling to the top.
   */
  const applyVAlign = (paragraph: Element | null, role: BlockRole): boolean => {
    if (!paragraph || !frontTemplateSectPr) return false;
    // Printed books set a copyright notice at the foot of its page. A design
    // that says nothing is not asking for the top, it simply has no opinion.
    const wanted =
      reference.profile.roleVAlign[role as StyleRole] ?? (role === 'copyright' ? 'bottom' : undefined);
    if (!wanted) return false;
    const sectPr = frontTemplateSectPr.cloneNode(true) as Element;
    setSectionType(sectPr, 'nextPage', outDoc);
    setVerticalAlignment(sectPr, wanted, outDoc);
    stripPageNumberRestart(sectPr);
    attachSectPr(paragraph, sectPr, outDoc);
    return true;
  };

  for (const page of generatedFront) {
    for (const p of page.paragraphs) outBody.appendChild(p);
    stats.paragraphsWritten += page.paragraphs.length;
    lastParagraph = page.paragraphs.at(-1) ?? lastParagraph;
    applyVAlign(lastParagraph, page.role);
  }
  /**
   * Sections cloned from the reference's body section. Only the first may keep
   * the reference's `pgNumType` start value — otherwise every chapter would
   * restart page numbering at 1.
   */
  let bodySectionsEmitted = 0;
  let pendingPageBreak = false;
  let anyContentEmitted = generatedFront.length > 0;
  /** The front-matter role currently being copied, for its section break. */
  let lastFrontRole: BlockRole | null = null;
  let frontSectionPending = frontSectPrClone !== null && options.includeFrontMatter;
  let inFrontMatter = true;

  for (const block of manuscript.analysis.blocks) {
    const role = effectiveRole(block, options);

    if (role === 'empty') {
      if (options.removeEmptyParagraphs) {
        stats.blanksRemoved++;
        continue;
      }
      const p = newParagraph(outDoc, roleStyles.body);
      outBody.appendChild(p);
      lastParagraph = p;
      stats.paragraphsWritten++;
      continue;
    }

    if (role === 'pageBreak') {
      pendingPageBreak = true;
      continue;
    }

    // The manuscript's own opening pages are dropped when they have been
    // turned off, or when generated pages are standing in for them.
    if (FRONT_MATTER_ROLES.has(role) && (!options.includeFrontMatter || options.replaceFrontMatter)) {
      continue;
    }

    // A copied front-matter page gets the same vertical placement the
    // template gives that kind of page, closing the run when the role changes.
    if (
      inFrontMatter &&
      lastFrontRole !== null &&
      role !== lastFrontRole &&
      FRONT_MATTER_ROLES.has(lastFrontRole)
    ) {
      applyVAlign(lastParagraph, lastFrontRole);
    }
    if (FRONT_MATTER_ROLES.has(role)) lastFrontRole = role;

    const leavingFrontMatter = inFrontMatter && !FRONT_MATTER_ROLES.has(role);
    if (leavingFrontMatter) inFrontMatter = false;

    // A contents list is the last thing in the front matter, so it goes in
    // once the manuscript's own opening pages are behind us — never ahead of
    // the author's title page.
    if (leavingFrontMatter && contentsPage) {
      applyVAlign(lastParagraph, lastFrontRole ?? 'frontMatter');
      for (const p of contentsPage.paragraphs) outBody.appendChild(p);
      stats.paragraphsWritten += contentsPage.paragraphs.length;
      lastParagraph = contentsPage.paragraphs.at(-1) ?? lastParagraph;
      anyContentEmitted = true;
      contentsPage = null;
    }

    // Close the front-matter section on the last paragraph that belongs to it,
    // so the reference's own front/body section split is reproduced.
    let sectionBreakApplied = false;
    if (leavingFrontMatter && frontSectionPending && lastParagraph && frontSectPrClone) {
      attachSectPr(lastParagraph, frontSectPrClone, outDoc);
      frontSectionPending = false;
      sectionBreakApplied = true;
    }

    const node = manuscript.nodes[block.index];
    if (block.kind === 'table') {
      const tbl = importNode(outDoc, node);
      await migrator.migrate(tbl);
      await remapNotes(tbl, footnotes, endnotes);
      outBody.appendChild(tbl);
      stats.tables++;
      anyContentEmitted = true;
      lastParagraph = null;
      pendingPageBreak = false;
      continue;
    }

    const startsChapter = CHAPTER_ROLES.has(role);
    // A manual page break sitting before a chapter opening is an artefact of
    // the manuscript's own pagination, so the template's chapter rule replaces
    // it. Breaks anywhere else were deliberate and are kept.
    let breakBefore = pendingPageBreak && !CHAPTER_ROLES.has(block.autoRole);
    pendingPageBreak = false;

    if (startsChapter && anyContentEmitted && !sectionBreakApplied) {
      if (options.chapterStart === 'oddPage' && lastParagraph && bodySectPrClone) {
        const sectPr = bodySectPrClone.cloneNode(true) as Element;
        setSectionType(sectPr, 'oddPage', outDoc);
        if (bodySectionsEmitted > 0) stripPageNumberRestart(sectPr);
        bodySectionsEmitted++;
        attachSectPr(lastParagraph, sectPr, outDoc);
        breakBefore = false;
      } else if (options.chapterStart !== 'continuous') {
        breakBefore = true;
      }
    }

    // Reproduce the template's chapter sink: blank paragraphs in the chapter
    // style that push the title down the page. These are generated, so the
    // "drop blank paragraphs" option must not remove them.
    if (startsChapter && roleStyles.chapterTitle) {
      for (let i = 0; i < chapterBlanksBefore; i++) {
        const filler = newParagraph(outDoc, roleStyles.chapterTitle);
        if (i === 0 && breakBefore) {
          const pPr = child(filler, 'pPr');
          pPr?.appendChild(wEl(outDoc, 'pageBreakBefore'));
          breakBefore = false;
        }
        outBody.appendChild(filler);
        lastParagraph = filler;
        stats.paragraphsWritten++;
      }
    }

    const p = buildParagraph({
      outDoc,
      role,
      block,
      source: node,
      roleStyles,
      reference,
      options,
      copyOptions,
      breakBefore,
      bodyIndent,
    });

    const numPr = child(child(node, 'pPr'), 'numPr');
    if (role === 'listItem' && numPr) {
      const mapped = await mapNumbering(numPr, numbering, outDoc);
      if (mapped) insertIntoPPr(p, mapped, outDoc);
    }

    await migrator.migrate(p);
    await remapNotes(p, footnotes, endnotes);
    outBody.appendChild(p);

    lastParagraph = p;
    anyContentEmitted = true;
    stats.paragraphsWritten++;
    stats.wordCount += block.wordCount;

    if (startsChapter && roleStyles.chapterTitle) {
      for (let i = 0; i < chapterBlanksAfter; i++) {
        const filler = newParagraph(outDoc, roleStyles.chapterTitle);
        outBody.appendChild(filler);
        lastParagraph = filler;
        stats.paragraphsWritten++;
      }
    }
    if (role === 'chapterTitle') stats.chapters++;
    else if (role === 'partTitle') stats.parts++;
    if (role === 'sceneBreak') stats.sceneBreaks++;
  }

  const generatedBack = buildBackMatter(matterContext);
  for (const p of generatedBack) outBody.appendChild(p);
  stats.paragraphsWritten += generatedBack.length;

  // A contents table is an empty field until Word evaluates it, so ask Word to
  // offer that when the document opens.
  if (needsFieldUpdate(options.extraSections)) {
    await requestFieldUpdate(out, outRels);
  }

  if (contentsPage) {
    for (const p of contentsPage.paragraphs) outBody.appendChild(p);
    stats.paragraphsWritten += contentsPage.paragraphs.length;
  }

  if (bodySectPrClone) {
    // The final section only restarts numbering when it is the only one.
    if (bodySectionsEmitted > 0) stripPageNumberRestart(bodySectPrClone);
    outBody.appendChild(bodySectPrClone);
  }

  // Put the author's own title and name into the running heads, replacing a
  // template's placeholder wording. Skipped entirely when nothing was typed.
  const headsChanged = await applyDetailsToRunningHeads(out, outRels, options.bookDetails);
  if (headsChanged.changed.length > 0) stats.runningHeadsUpdated = headsChanged.changed.length;

  numbering.save();
  footnotes.save();
  endnotes.save();
  migrator.save();
  out.writeXml(out.documentPath, outDoc);

  stats.imagesCopied = migrator.imagesCopied;
  stats.footnotesCopied = footnotes.copied + endnotes.copied;

  warnings.push(...migrator.warnings, ...numbering.warnings, ...footnotes.warnings, ...endnotes.warnings);
  if (stats.paragraphsWritten === 0) {
    warnings.push('No content was written. Check the manuscript and the front-matter option.');
  }
  return { pkg: out, stats, warnings: dedupe(warnings) };
}

function effectiveRole(block: ManuscriptBlock, options: FormatOptions): BlockRole {
  return options.roleOverrides[block.index] ?? block.role;
}

/** Merge detected role styles with user overrides, dropping unknown ids. */
function resolveRoleStyles(
  reference: LoadedReference,
  options: FormatOptions,
  warnings: string[],
): Record<StyleRole, string | null> {
  const merged = { ...reference.profile.roleStyles };
  for (const [role, id] of Object.entries(options.roleStyles) as Array<[StyleRole, string | null]>) {
    if (id === undefined) continue;
    if (id !== null && !reference.styles.has(id)) {
      warnings.push(`Style "${id}" is not defined in the reference and was ignored for ${role}.`);
      continue;
    }
    merged[role] = id;
  }
  return merged;
}

interface BuildParagraphArgs {
  outDoc: Document;
  role: BlockRole;
  block: ManuscriptBlock;
  source: Element;
  roleStyles: Record<StyleRole, string | null>;
  reference: LoadedReference;
  options: FormatOptions;
  copyOptions: Parameters<typeof copyParagraphContent>[1];
  breakBefore: boolean;
  bodyIndent: number;
}

function buildParagraph(args: BuildParagraphArgs): Element {
  const { outDoc, role, block, source, roleStyles, options, breakBefore, bodyIndent } = args;
  const p = outDoc.createElementNS(NS.w, 'w:p');
  const pPr = wEl(outDoc, 'pPr');

  const styleId = styleForRole(role, roleStyles);
  const direct = directFormatting(role, styleId, block, roleStyles, options, bodyIndent);

  // `w:pPr` children must follow the schema's sequence or Word rejects the
  // file: pStyle, keepNext, pageBreakBefore, numPr, spacing, ind, jc, sectPr.
  if (styleId) pPr.appendChild(wEl(outDoc, 'pStyle', { val: styleId }));
  if (direct.keepNext) pPr.appendChild(wEl(outDoc, 'keepNext'));
  if (breakBefore) pPr.appendChild(wEl(outDoc, 'pageBreakBefore'));
  if (direct.spacing) {
    pPr.appendChild(wEl(outDoc, 'spacing', direct.spacing));
  }
  if (direct.indent) {
    pPr.appendChild(wEl(outDoc, 'ind', direct.indent));
  }
  if (direct.alignment) {
    pPr.appendChild(wEl(outDoc, 'jc', { val: direct.alignment }));
  }
  p.appendChild(pPr);

  const content = copyParagraphContent(source, args.copyOptions);
  for (const node of content) p.appendChild(node);
  if (direct.bold) applyBold(p, outDoc);

  // Authors bold and underline their headings as a stand-in for a real style.
  // Where the reference supplies one, that styling wins. Italic and small caps
  // stay: inside a heading they usually mark a title, not decoration.
  if (HEADING_ROLES.has(role) && styleId && styleId !== roleStyles.body && !direct.bold) {
    stripRunToggles(p, ['b', 'bCs', 'u', 'caps']);
  }

  // A scene break can be replaced with the reference's own ornament text.
  if (role === 'sceneBreak' && options.sceneBreakText !== null) {
    replaceText(p, outDoc, options.sceneBreakText);
  }
  return p;
}

interface DirectFormatting {
  alignment: string | null;
  indent: Record<string, string | number> | null;
  spacing: Record<string, string | number> | null;
  keepNext: boolean;
  bold: boolean;
}

/**
 * Direct formatting applied on top of the style. Used only where the reference
 * offers no dedicated style, or where an option asks for a specific effect.
 */
function directFormatting(
  role: BlockRole,
  styleId: string | null,
  block: ManuscriptBlock,
  roles: Record<StyleRole, string | null>,
  options: FormatOptions,
  bodyIndent: number,
): DirectFormatting {
  const none: DirectFormatting = {
    alignment: null,
    indent: null,
    spacing: null,
    keepNext: false,
    bold: false,
  };

  switch (role) {
    case 'chapterTitle':
    case 'partTitle':
    case 'frontMatterTitle': {
      if (styleId && styleId !== roles.body) return none;
      // No heading style in the reference: centre it and set it apart.
      return {
        alignment: 'center',
        indent: { firstLine: 0, left: 0 },
        spacing: { before: 480, after: 360 },
        keepNext: true,
        bold: true,
      };
    }
    case 'chapterSubtitle':
    case 'subheading': {
      if (styleId && styleId !== roles.body) return none;
      return {
        alignment: role === 'chapterSubtitle' ? 'center' : null,
        indent: { firstLine: 0 },
        spacing: { before: 240, after: 120 },
        keepNext: true,
        bold: true,
      };
    }
    case 'sceneBreak': {
      if (styleId && styleId !== roles.body) return none;
      return {
        alignment: 'center',
        indent: { firstLine: 0, left: 0 },
        spacing: { before: 240, after: 240 },
        keepNext: false,
        bold: false,
      };
    }
    case 'blockQuote': {
      if (styleId && styleId !== roles.body) return none;
      return {
        alignment: null,
        indent: { left: 720, right: 720, firstLine: 0 },
        spacing: { before: 120, after: 120 },
        keepNext: false,
        bold: false,
      };
    }
    case 'bodyFirst': {
      // Only needed when the reference has no dedicated first-paragraph style.
      const needsFlush =
        options.firstParagraphNoIndent && bodyIndent > 0 && (roles.bodyFirst ?? roles.body) === roles.body;
      const centred = imageOrCentred(block);
      return {
        alignment: centred ? 'center' : null,
        indent: needsFlush || centred ? { firstLine: 0 } : null,
        spacing: null,
        keepNext: false,
        bold: false,
      };
    }
    case 'body': {
      const centred = imageOrCentred(block);
      return {
        alignment: centred ? 'center' : null,
        indent: centred ? { firstLine: 0 } : null,
        spacing: null,
        keepNext: false,
        bold: false,
      };
    }
    default:
      return none;
  }
}

/**
 * Short centred paragraphs and image-only paragraphs keep their centring: an
 * epigraph, a letter or a figure is deliberate, not stray formatting.
 */
function imageOrCentred(block: ManuscriptBlock): boolean {
  if (block.hasImage) return true;
  return block.alignment === 'center' && block.wordCount > 0 && block.wordCount <= 25;
}

function newParagraph(doc: Document, styleId: string | null): Element {
  const p = doc.createElementNS(NS.w, 'w:p');
  if (styleId) {
    const pPr = wEl(doc, 'pPr');
    pPr.appendChild(wEl(doc, 'pStyle', { val: styleId }));
    p.appendChild(pPr);
  }
  return p;
}

/** Force bold on every run, for headings with no reference style to lean on. */
function applyBold(p: Element, doc: Document): void {
  for (const run of descendants(p, 'r')) {
    let rPr = child(run, 'rPr');
    if (!rPr) {
      rPr = wEl(doc, 'rPr');
      run.insertBefore(rPr, run.firstChild);
    }
    if (!child(rPr, 'b')) rPr.insertBefore(wEl(doc, 'b'), rPr.firstChild);
  }
}

/** Remove named toggles from every run, and drop the `w:rPr` left empty. */
function stripRunToggles(p: Element, names: string[]): void {
  for (const run of descendants(p, 'r')) {
    const rPr = child(run, 'rPr');
    if (!rPr) continue;
    for (const name of names) {
      const el = child(rPr, name);
      if (el) rPr.removeChild(el);
    }
    if (childElementCount(rPr) === 0) run.removeChild(rPr);
  }
}

/** Replace a paragraph's text with a single run, keeping its properties. */
function replaceText(p: Element, doc: Document, text: string): void {
  for (const node of [...children(p, 'r'), ...children(p, 'hyperlink')]) p.removeChild(node);
  const run = doc.createElementNS(NS.w, 'w:r');
  const t = doc.createElementNS(NS.w, 'w:t');
  t.appendChild(doc.createTextNode(text));
  run.appendChild(t);
  p.appendChild(run);
}

async function mapNumbering(
  numPr: Element,
  numbering: NumberingMerger,
  doc: Document,
): Promise<Element | null> {
  const sourceNumId = attr(child(numPr, 'numId'), 'val');
  if (!sourceNumId) return null;
  const mapped = await numbering.mapNumId(sourceNumId);
  if (!mapped) return null;
  const copy = wEl(doc, 'numPr');
  const ilvl = attr(child(numPr, 'ilvl'), 'val');
  if (ilvl !== null) copy.appendChild(wEl(doc, 'ilvl', { val: ilvl }));
  copy.appendChild(wEl(doc, 'numId', { val: mapped }));
  return copy;
}

/** `numPr` must follow `pageBreakBefore` and precede `spacing` in `w:pPr`. */
function insertIntoPPr(p: Element, numPr: Element, doc: Document): void {
  let pPr = child(p, 'pPr');
  if (!pPr) {
    pPr = wEl(doc, 'pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  const anchor =
    child(pPr, 'spacing') ?? child(pPr, 'ind') ?? child(pPr, 'jc') ?? null;
  if (anchor) pPr.insertBefore(numPr, anchor);
  else pPr.appendChild(numPr);
}

async function remapNotes(
  container: Element,
  footnotes: NoteMerger,
  endnotes: NoteMerger,
): Promise<void> {
  const jobs: Array<[Element, NoteMerger]> = [
    ...descendants(container, 'footnoteReference').map(
      (el) => [el, footnotes] as [Element, NoteMerger],
    ),
    ...descendants(container, 'endnoteReference').map((el) => [el, endnotes] as [Element, NoteMerger]),
  ];
  for (const [el, merger] of jobs) {
    const sourceId = attr(el, 'id');
    const mapped = sourceId ? await merger.mapId(sourceId) : null;
    if (mapped) {
      el.setAttributeNS(NS.w, 'w:id', mapped);
      continue;
    }
    const run = el.parentNode as Element | null;
    el.parentNode?.removeChild(el);
    // Drop a run left holding nothing but its properties.
    if (run && run.localName === 'r' && children(run, 'rPr').length === childElementCount(run)) {
      run.parentNode?.removeChild(run);
    }
  }
}

function childElementCount(el: Element): number {
  let count = 0;
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) count++;
  return count;
}

/** Put a section break at the end of a paragraph. */
function attachSectPr(p: Element, sectPr: Element, doc: Document): void {
  let pPr = child(p, 'pPr');
  if (!pPr) {
    pPr = wEl(doc, 'pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  if (child(pPr, 'sectPr')) return;
  pPr.appendChild(importNode(doc, sectPr));
}

/**
 * Where a section's text sits vertically on its page. `w:vAlign` follows the
 * column settings and precedes `w:titlePg` in the schema's sequence.
 */
function setVerticalAlignment(sectPr: Element, value: string, doc: Document): void {
  const existing = child(sectPr, 'vAlign');
  if (existing) {
    existing.setAttributeNS(NS.w, 'w:val', value);
    return;
  }
  const vAlign = wEl(doc, 'vAlign', { val: value });
  const anchor = child(sectPr, 'titlePg') ?? child(sectPr, 'docGrid') ?? null;
  if (anchor) sectPr.insertBefore(vAlign, anchor);
  else sectPr.appendChild(vAlign);
}

/** `w:type` sits after the note properties and before `w:pgSz`. */
function setSectionType(sectPr: Element, value: string, doc: Document): void {
  const existing = child(sectPr, 'type');
  if (existing) {
    existing.setAttributeNS(NS.w, 'w:val', value);
    return;
  }
  const type = wEl(doc, 'type', { val: value });
  const anchor = child(sectPr, 'pgSz') ?? child(sectPr, 'pgMar') ?? null;
  if (anchor) sectPr.insertBefore(type, anchor);
  else sectPr.appendChild(type);
}

/**
 * Repeated sections must not each restart page numbering, so the explicit
 * start value is removed from the cloned section properties.
 */
function stripPageNumberRestart(sectPr: Element): void {
  const pgNumType = child(sectPr, 'pgNumType');
  if (!pgNumType) return;
  pgNumType.removeAttributeNS(NS.w, 'start');
  pgNumType.removeAttribute('w:start');
  if (pgNumType.attributes.length === 0) sectPr.removeChild(pgNumType);
}

function dedupe(items: string[]): string[] {
  return [...new Set(items)];
}

/**
 * Set `w:updateFields` so Word offers to fill in the contents table on open.
 * The settings part is created if the reference has none.
 */
async function requestFieldUpdate(out: DocxPackage, rels: Relationships): Promise<void> {
  let target = rels.firstTargetOfType(RELTYPE.settings);
  if (!target) {
    target = 'settings.xml';
    rels.add(RELTYPE.settings, target);
    out.writeText(
      `${out.documentDir}settings.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:settings xmlns:w="${NS.w}"/>`,
    );
    await out.ensureContentType({
      partName: `${out.documentDir}settings.xml`,
      partType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    });
  }
  const path = out.resolveTarget(out.documentPath, target);
  const doc = await out.readXml(path);
  const root = doc?.documentElement;
  if (!doc || !root) return;
  if (child(root, 'updateFields')) return;
  // `w:updateFields` belongs near the top of the settings sequence.
  root.insertBefore(wEl(doc, 'updateFields', { val: 'true' }), root.firstChild);
  out.writeXml(path, doc);
}
