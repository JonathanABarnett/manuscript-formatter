import { NS, RELTYPE } from '../ooxml/ns.js';
import type { DocxPackage } from '../ooxml/package.js';
import {
  attr,
  child,
  children,
  clearChildren,
  descendants,
  importNode,
  textOf,
  wEl,
} from '../ooxml/xml.js';
import type { BlockRole, FormatOptions, FormatStats, ManuscriptBlock, StyleRole } from '../types.js';
import type { LoadedReference } from '../analyze/reference.js';
import { collectSectPrs } from '../analyze/reference.js';
import type { LoadedManuscript } from '../analyze/manuscript.js';
import { chapterTitleTexts, leadInLength, styleForRole } from '../roles.js';
import { parseChapterTitle } from '../analyze/chapterNumbers.js';
import {
  buildBackMatter,
  buildContentsPage,
  buildFrontMatter,
  needsFieldUpdate,
} from './matter.js';
import { applyDetailsToRunningHeads } from './headers.js';
import { applyLanguage, applyPrintSettings, requestFieldUpdate } from './settings.js';
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
  // Which page side each header serves, read before the body is emptied.
  const headerSides = new Map<string, string>();
  for (const sectPr of outSectPrs) {
    for (const ref of children(sectPr, 'headerReference')) {
      const id = attr(ref, 'id', NS.r);
      const type = attr(ref, 'type');
      if (id && type && !headerSides.has(id)) headerSides.set(id, type);
    }
  }

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
    underlineToItalic: options.underlineToItalic,
    removeManualIndents: options.removeManualIndents,
    transformer,
  };

  const bodyIndent = roleStyles.body
    ? reference.styles.resolve(roleStyles.body).firstLineIndentTwips ?? 0
    : 0;

  // Chapter titles rewritten to a uniform numbering, where that was asked for.
  const chapterTitles = chapterTitleTexts(
    manuscript.analysis.blocks,
    (b) => effectiveRole(b, options),
    options,
  );

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

  // The chapter titles as the contents page will list them until Word fills
  // the table in for real: as written, or as the numbering options rewrite them.
  const contentsEntries = manuscript.analysis.blocks
    .filter((b) => {
      const role = effectiveRole(b, options);
      return (role === 'chapterTitle' || role === 'partTitle') && b.text.trim().length > 0;
    })
    .map((b) => chapterTitles.get(b.index) ?? b.text.replace(/\s+/g, ' ').trim());

  // Generated opening pages come first, before anything from the manuscript.
  const matterContext = {
    doc: outDoc,
    roles: roleStyles,
    details: options.bookDetails,
    sections: options.extraSections,
    chapterBlanksBefore,
    chapterBlanksAfter,
    contentsEntries,
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
  /**
   * Chapters become sections of their own when they must open on a recto, or
   * when their opening page is to carry no running head — Word can only
   * leave a header off the first page of a section.
   */
  const headerlessOpeners = options.chapterOpenerNoHeader && options.chapterStart !== 'continuous';
  const chapterSections = options.chapterStart === 'oddPage' || headerlessOpeners;
  let pendingPageBreak = false;
  let anyContentEmitted = generatedFront.length > 0;
  /** The front-matter role currently being copied, for its section break. */
  let lastFrontRole: BlockRole | null = null;
  let frontSectionPending = frontSectPrClone !== null && options.includeFrontMatter;
  let inFrontMatter = true;
  /** The next text paragraph opens a chapter, so it may take the lead-in. */
  let openerPending = false;

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

    let sectionBreakApplied = false;

    // A contents list is the last thing in the front matter, so it goes in
    // once the manuscript's own opening pages are behind us — never ahead of
    // the author's title page. It gets a section of its own, aligned to the
    // top of the page: a design that centres its title page or sets its
    // copyright notice at the foot must not drag the contents down with it.
    if (leavingFrontMatter && contentsPage) {
      applyVAlign(lastParagraph, lastFrontRole ?? 'frontMatter');
      if (frontSectionPending && lastParagraph && frontSectPrClone) {
        attachSectPr(lastParagraph, frontSectPrClone, outDoc);
        frontSectionPending = false;
      }
      for (const p of contentsPage.paragraphs) outBody.appendChild(p);
      stats.paragraphsWritten += contentsPage.paragraphs.length;
      lastParagraph = contentsPage.paragraphs.at(-1) ?? lastParagraph;
      anyContentEmitted = true;
      const template = frontSectPrClone ?? bodySectPrClone;
      if (template && lastParagraph) {
        const sectPr = template.cloneNode(true) as Element;
        setSectionType(sectPr, 'nextPage', outDoc);
        setVerticalAlignment(sectPr, 'top', outDoc);
        stripPageNumberRestart(sectPr);
        attachSectPr(lastParagraph, sectPr, outDoc);
        sectionBreakApplied = true;
      }
      contentsPage = null;
    }

    // Close the front-matter section on the last paragraph that belongs to it,
    // so the reference's own front/body section split is reproduced.
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
      openerPending = false;
      continue;
    }

    const startsChapter = CHAPTER_ROLES.has(role);
    // A manual page break sitting before a chapter opening is an artefact of
    // the manuscript's own pagination, so the template's chapter rule replaces
    // it. Breaks anywhere else were deliberate and are kept.
    let breakBefore = pendingPageBreak && !CHAPTER_ROLES.has(block.autoRole);
    pendingPageBreak = false;

    if (startsChapter && anyContentEmitted && !sectionBreakApplied) {
      if (chapterSections && lastParagraph && bodySectPrClone) {
        const sectPr = bodySectPrClone.cloneNode(true) as Element;
        setSectionType(sectPr, options.chapterStart === 'oddPage' ? 'oddPage' : 'nextPage', outDoc);
        if (bodySectionsEmitted > 0) stripPageNumberRestart(sectPr);
        if (headerlessOpeners) markFirstPageDifferent(sectPr, outDoc);
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
      chapterTitles,
    });

    const numPr = child(child(node, 'pPr'), 'numPr');
    if (role === 'listItem' && numPr) {
      const mapped = await mapNumbering(numPr, numbering, outDoc);
      if (mapped) insertIntoPPr(p, mapped, outDoc);
    }

    await migrator.migrate(p);
    await remapNotes(p, footnotes, endnotes);
    outBody.appendChild(p);

    // The opening words of a chapter's first paragraph, in small capitals.
    if (CHAPTER_ROLES.has(role) || role === 'chapterSubtitle') {
      openerPending = true;
    } else if (openerPending) {
      if (options.leadInSmallCaps && (role === 'bodyFirst' || role === 'body') && !block.hasImage) {
        applySmallCapsPrefix(p, outDoc, leadInLength(textOf(p)));
      }
      openerPending = false;
    }

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
  // Settings that make the file print well: hyphenation as chosen, fonts
  // embedded on the author's next save, and the manuscript's own language.
  await applyPrintSettings(out, outRels, { hyphenate: options.hyphenate });
  if (manuscript.analysis.language) {
    await applyLanguage(out, outRels, manuscript.analysis.language);
  }

  if (contentsPage) {
    for (const p of contentsPage.paragraphs) outBody.appendChild(p);
    stats.paragraphsWritten += contentsPage.paragraphs.length;
  }

  if (bodySectPrClone) {
    // The final section only restarts numbering when it is the only one.
    if (bodySectionsEmitted > 0) stripPageNumberRestart(bodySectPrClone);
    // The last chapter's opening page is the first page of this section.
    if (headerlessOpeners && stats.chapters + stats.parts > 0) {
      markFirstPageDifferent(bodySectPrClone, outDoc);
    }
    outBody.appendChild(bodySectPrClone);
  }

  // Put the author's own title and name into the running heads, replacing a
  // template's placeholder wording. Skipped entirely when nothing was typed.
  const headsChanged = await applyDetailsToRunningHeads(
    out,
    outRels,
    options.bookDetails,
    options.runningHeads,
    headerSides,
  );
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
  /** Replacement text for chapter titles whose numbering is being changed. */
  chapterTitles: Map<number, string>;
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

  // A chapter number written to the chosen style. Only the label and number
  // change, so a title after them keeps whatever emphasis it carried.
  const retitled = role === 'chapterTitle' ? args.chapterTitles.get(block.index) : undefined;
  if (retitled !== undefined) {
    replaceLeadingText(p, outDoc, retitled, parseChapterTitle(block.text)?.rest ?? '');
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

/**
 * Replace everything before `keepTail` with the start of `newText`, keeping
 * the runs that carry the tail. Falls back to a plain replacement when the
 * paragraph's text does not end with the tail as expected.
 */
function replaceLeadingText(p: Element, doc: Document, newText: string, keepTail: string): void {
  // Trailing whitespace in the paragraph sits after the tail and is left alone.
  const raw = textOf(p).replace(/\s+$/, '');
  if (!keepTail || !raw.endsWith(keepTail) || !newText.endsWith(keepTail)) {
    replaceText(p, doc, newText);
    return;
  }
  const prefixLen = raw.length - keepTail.length;
  const newPrefix = newText.slice(0, newText.length - keepTail.length);

  const runs = descendants(p, 'r');
  let consumed = 0;
  // The new prefix takes the formatting of the run it replaces.
  const firstProps: Element | null = runs.length > 0 ? child(runs[0], 'rPr') : null;
  let anchor: Element | null = null;
  for (const run of runs) {
    const length = textOf(run).length;
    if (consumed >= prefixLen) {
      anchor = run;
      break;
    }
    if (consumed + length <= prefixLen) {
      consumed += length;
      // Keep a run that holds a picture or a note; drop one that only had text.
      if (descendants(run, 'drawing').length === 0 && descendants(run, 'footnoteReference').length === 0) {
        run.parentNode?.removeChild(run);
      }
      continue;
    }
    // The boundary falls inside this run: trim its leading characters.
    let remaining = prefixLen - consumed;
    for (let n = run.firstChild; n && remaining > 0; ) {
      const next = n.nextSibling;
      if (n.nodeType === 1) {
        const el = n as Element;
        if (el.namespaceURI === NS.w && el.localName === 't') {
          const text = el.textContent ?? '';
          if (text.length <= remaining) {
            remaining -= text.length;
            run.removeChild(el);
          } else {
            el.textContent = text.slice(remaining);
            remaining = 0;
          }
        } else if (
          el.namespaceURI === NS.w &&
          (el.localName === 'tab' || el.localName === 'br' || el.localName === 'cr' || el.localName === 'noBreakHyphen')
        ) {
          remaining -= 1;
          run.removeChild(el);
        }
      }
      n = next;
    }
    consumed = prefixLen;
    anchor = run;
    break;
  }

  const run = doc.createElementNS(NS.w, 'w:r');
  if (firstProps) run.appendChild(importNode(doc, firstProps));
  const t = doc.createElementNS(NS.w, 'w:t');
  if (newPrefix !== newPrefix.trim()) t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
  t.appendChild(doc.createTextNode(newPrefix));
  run.appendChild(t);
  if (anchor?.parentNode) anchor.parentNode.insertBefore(run, anchor);
  else p.appendChild(run);
}

/** Put the first `length` characters of a paragraph in small capitals. */
function applySmallCapsPrefix(p: Element, doc: Document, length: number): void {
  let remaining = length;
  for (const run of descendants(p, 'r')) {
    if (remaining <= 0) break;
    const runLength = textOf(run).length;
    if (runLength === 0) continue;
    if (runLength > remaining) splitRun(run, remaining, doc);
    let rPr = child(run, 'rPr');
    if (!rPr) {
      rPr = wEl(doc, 'rPr');
      run.insertBefore(rPr, run.firstChild);
    }
    if (!child(rPr, 'smallCaps')) {
      // `w:smallCaps` follows `w:caps` and precedes `w:strike` in `w:rPr`.
      const anchor =
        child(rPr, 'strike') ?? child(rPr, 'dstrike') ?? child(rPr, 'vanish') ?? child(rPr, 'u') ??
        child(rPr, 'vertAlign') ?? child(rPr, 'lang') ?? null;
      if (anchor) rPr.insertBefore(wEl(doc, 'smallCaps'), anchor);
      else rPr.appendChild(wEl(doc, 'smallCaps'));
    }
    remaining -= Math.min(runLength, remaining);
  }
}

/**
 * Split a run after `offset` characters of its text, moving the rest into a
 * new run with the same properties immediately after it.
 */
function splitRun(run: Element, offset: number, doc: Document): void {
  const tail = doc.createElementNS(NS.w, 'w:r');
  const rPr = child(run, 'rPr');
  if (rPr) tail.appendChild(rPr.cloneNode(true));
  const moving: Node[] = [];
  let consumed = 0;
  for (let n = run.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    if (el.namespaceURI === NS.w && el.localName === 'rPr') continue;
    if (consumed >= offset) {
      moving.push(el);
      continue;
    }
    if (el.namespaceURI === NS.w && el.localName === 't') {
      const text = el.textContent ?? '';
      if (consumed + text.length <= offset) {
        consumed += text.length;
        continue;
      }
      const cut = offset - consumed;
      const rest = text.slice(cut);
      el.textContent = text.slice(0, cut);
      el.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      const t = doc.createElementNS(NS.w, 'w:t');
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(rest));
      moving.push(t);
      consumed = offset;
      continue;
    }
    if (
      el.namespaceURI === NS.w &&
      (el.localName === 'tab' || el.localName === 'br' || el.localName === 'cr' || el.localName === 'noBreakHyphen')
    ) {
      consumed += 1;
    }
  }
  for (const node of moving) tail.appendChild(node);
  run.parentNode?.insertBefore(tail, run.nextSibling);
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

/**
 * Give the section a distinct first page with no header. Word then uses the
 * design's first-page header and footer for it — usually none — so a page
 * number that lives in the footer is carried over as the first-page footer
 * when the design has not said otherwise. `w:titlePg` follows `w:vAlign` and
 * precedes `w:textDirection`, `w:bidi`, `w:rtlGutter` and `w:docGrid`.
 */
function markFirstPageDifferent(sectPr: Element, doc: Document): void {
  if (!child(sectPr, 'titlePg')) {
    const titlePg = wEl(doc, 'titlePg');
    const anchor =
      child(sectPr, 'textDirection') ??
      child(sectPr, 'bidi') ??
      child(sectPr, 'rtlGutter') ??
      child(sectPr, 'docGrid') ??
      null;
    if (anchor) sectPr.insertBefore(titlePg, anchor);
    else sectPr.appendChild(titlePg);
  }
  const footers = children(sectPr, 'footerReference');
  const hasFirstFooter = footers.some((f) => attr(f, 'type') === 'first');
  const defaultFooter = footers.find((f) => attr(f, 'type') === 'default');
  if (!hasFirstFooter && defaultFooter) {
    const first = defaultFooter.cloneNode(true) as Element;
    first.setAttributeNS(NS.w, 'w:type', 'first');
    // Reference elements open the section: after the last footer reference.
    const last = footers[footers.length - 1];
    if (last.nextSibling) sectPr.insertBefore(first, last.nextSibling);
    else sectPr.appendChild(first);
  }
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
