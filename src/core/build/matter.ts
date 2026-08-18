import { NS } from '../ooxml/ns.js';
import { wEl } from '../ooxml/xml.js';
import { styleForRole, type RoleStyleMap } from '../roles.js';
import type { BlockRole, BookDetails, ExtraSections } from '../types.js';

/**
 * Builds the pages a manuscript usually lacks — a title page, a copyright
 * page, a dedication, a contents list — from details the author types in,
 * using the chosen design's own styles.
 *
 * Nothing here invents content. A section is only produced when the author has
 * supplied the words for it, so no page can appear with a heading and a blank
 * space under it.
 */

export interface MatterContext {
  doc: Document;
  roles: RoleStyleMap;
  details: BookDetails;
  sections: ExtraSections;
}

/** Whether a section will actually produce anything, given what was typed. */
export function sectionHasContent(
  section: keyof ExtraSections,
  details: BookDetails,
): boolean {
  switch (section) {
    case 'titlePage':
      return details.title.trim().length > 0;
    case 'copyrightPage':
      return (
        details.copyrightYear.trim().length > 0 ||
        details.author.trim().length > 0 ||
        details.publisher.trim().length > 0 ||
        details.isbn.trim().length > 0
      );
    case 'alsoBy':
      return details.alsoBy.trim().length > 0;
    case 'dedication':
      return details.dedication.trim().length > 0;
    case 'epigraph':
      return details.epigraph.trim().length > 0;
    case 'acknowledgments':
      return details.acknowledgments.trim().length > 0;
    case 'aboutTheAuthor':
      return details.aboutTheAuthor.trim().length > 0;
    case 'bibliography':
      return details.bibliography.trim().length > 0;
    case 'contents':
      // The table is a Word field, so it needs nothing typed in.
      return true;
    default:
      return false;
  }
}

function para(
  ctx: MatterContext,
  role: BlockRole,
  text: string,
  opts: { breakBefore?: boolean; align?: string; italic?: boolean } = {},
): Element {
  const p = ctx.doc.createElementNS(NS.w, 'w:p');
  const pPr = wEl(ctx.doc, 'pPr');
  const styleId = styleForRole(role, ctx.roles);
  if (styleId) pPr.appendChild(wEl(ctx.doc, 'pStyle', { val: styleId }));
  if (opts.breakBefore) pPr.appendChild(wEl(ctx.doc, 'pageBreakBefore'));
  if (opts.align) pPr.appendChild(wEl(ctx.doc, 'jc', { val: opts.align }));
  p.appendChild(pPr);

  if (text.length > 0) {
    const run = ctx.doc.createElementNS(NS.w, 'w:r');
    if (opts.italic) {
      const rPr = wEl(ctx.doc, 'rPr');
      rPr.appendChild(wEl(ctx.doc, 'i'));
      run.appendChild(rPr);
    }
    const t = ctx.doc.createElementNS(NS.w, 'w:t');
    t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
    t.appendChild(ctx.doc.createTextNode(text));
    run.appendChild(t);
    p.appendChild(run);
  }
  return p;
}

/** Split a typed block into paragraphs on blank lines, keeping the author's breaks. */
function lines(text: string): string[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
}

/**
 * A Word table-of-contents field. Word fills in the entries and page numbers
 * when the document's fields are updated, which the composer asks it to do on
 * open — a static list would go stale the moment anything reflowed.
 */
function contentsField(ctx: MatterContext): Element {
  const p = ctx.doc.createElementNS(NS.w, 'w:p');
  const pPr = wEl(ctx.doc, 'pPr');
  const styleId = styleForRole('body', ctx.roles);
  if (styleId) pPr.appendChild(wEl(ctx.doc, 'pStyle', { val: styleId }));
  pPr.appendChild(wEl(ctx.doc, 'ind', { firstLine: 0 }));
  p.appendChild(pPr);

  const run = (build: (r: Element) => void): void => {
    const r = ctx.doc.createElementNS(NS.w, 'w:r');
    build(r);
    p.appendChild(r);
  };

  run((r) => r.appendChild(wEl(ctx.doc, 'fldChar', { fldCharType: 'begin' })));
  run((r) => {
    const instr = ctx.doc.createElementNS(NS.w, 'w:instrText');
    instr.setAttributeNS(NS.xml, 'xml:space', 'preserve');
    // Chapter titles sit at outline level 1, so only those are listed.
    instr.appendChild(ctx.doc.createTextNode(' TOC \\o "1-1" \\h \\z \\u '));
    r.appendChild(instr);
  });
  run((r) => r.appendChild(wEl(ctx.doc, 'fldChar', { fldCharType: 'separate' })));
  run((r) => {
    const t = ctx.doc.createElementNS(NS.w, 'w:t');
    t.appendChild(
      ctx.doc.createTextNode('Right-click here in Word and choose “Update Field” to fill this in.'),
    );
    r.appendChild(t);
  });
  run((r) => r.appendChild(wEl(ctx.doc, 'fldChar', { fldCharType: 'end' })));
  return p;
}

/** A headed section: its title, then the author's own words beneath. */
function headedSection(ctx: MatterContext, heading: string, body: string): Element[] {
  const out: Element[] = [para(ctx, 'chapterTitle', heading, { breakBefore: true })];
  for (const line of lines(body)) out.push(para(ctx, 'frontMatter', line));
  return out;
}

/** One generated page, kept together so it can be given its own section. */
export interface MatterPage {
  /** The role whose vertical placement this page should follow. */
  role: BlockRole;
  paragraphs: Element[];
}

/** The pages that open the book, in the order a printed book sets them. */
export function buildFrontMatter(ctx: MatterContext): MatterPage[] {
  const { details, sections } = ctx;
  const pages: MatterPage[] = [];
  let out: Element[] = [];
  let first = true;
  const startPage = (): boolean => {
    const breakBefore = !first;
    first = false;
    return breakBefore;
  };
  const endPage = (role: BlockRole): void => {
    if (out.length > 0) pages.push({ role, paragraphs: out });
    out = [];
  };

  // Earlier books come first, ahead of the title page, as publishers set them.
  if (sections.alsoBy && sectionHasContent('alsoBy', details)) {
    const breakBefore = startPage();
    const author = details.author.trim();
    out.push(
      para(ctx, 'frontMatter', author ? `Also by ${author}` : 'Also by the author', {
        breakBefore,
        align: 'center',
      }),
    );
    for (const line of lines(details.alsoBy)) {
      out.push(para(ctx, 'frontMatter', line, { align: 'center', italic: true }));
    }
    endPage('frontMatter');
  }

  if (sections.titlePage && sectionHasContent('titlePage', details)) {
    const breakBefore = startPage();
    out.push(para(ctx, 'frontMatterTitle', details.title.trim(), { breakBefore }));
    if (details.subtitle.trim()) {
      out.push(para(ctx, 'chapterSubtitle', details.subtitle.trim(), { align: 'center' }));
    }
    if (details.author.trim()) {
      out.push(para(ctx, 'frontMatter', details.author.trim(), { align: 'center' }));
    }
    endPage('frontMatterTitle');
  }

  if (sections.copyrightPage && sectionHasContent('copyrightPage', details)) {
    const breakBefore = startPage();
    const year = details.copyrightYear.trim();
    const author = details.author.trim();
    const notice =
      year && author ? `Copyright © ${year} by ${author}` : year ? `Copyright © ${year}` : `Copyright © ${author}`;
    out.push(para(ctx, 'copyright', notice, { breakBefore }));
    out.push(para(ctx, 'copyright', 'All rights reserved.'));
    out.push(
      para(
        ctx,
        'copyright',
        'No part of this book may be reproduced in any form without written permission ' +
          'from the copyright holder, except brief quotations used in a review.',
      ),
    );
    if (details.publisher.trim()) out.push(para(ctx, 'copyright', details.publisher.trim()));
    if (details.isbn.trim()) out.push(para(ctx, 'copyright', `ISBN: ${details.isbn.trim()}`));
    endPage('copyright');
  }

  if (sections.dedication && sectionHasContent('dedication', details)) {
    const breakBefore = startPage();
    const [head, ...rest] = lines(details.dedication);
    out.push(para(ctx, 'frontMatter', head, { breakBefore, align: 'center' }));
    for (const line of rest) out.push(para(ctx, 'frontMatter', line, { align: 'center' }));
    endPage('frontMatter');
  }

  // An epigraph: the quotation in italics, its source on the last line set
  // to the right, in the design's quotation style where it has one.
  if (sections.epigraph && sectionHasContent('epigraph', details)) {
    const breakBefore = startPage();
    const all = lines(details.epigraph);
    const source = all.length > 1 ? all[all.length - 1] : null;
    const quote = source ? all.slice(0, -1) : all;
    quote.forEach((line, i) =>
      out.push(para(ctx, 'blockQuote', line, { breakBefore: breakBefore && i === 0, italic: true })),
    );
    if (source) {
      out.push(para(ctx, 'blockQuote', source.replace(/^[—–-]\s*/, '— '), { align: 'right' }));
    }
    endPage('frontMatter');
  }

  // Contents is deliberately not built here. It belongs at the end of the
  // front matter, after any title and copyright pages the manuscript carries
  // of its own, so the composer places it when the body begins.
  return pages;
}

/**
 * The contents page, which a printed book sets last in the front matter,
 * immediately before the body begins. Null when it was not asked for.
 */
export function buildContentsPage(ctx: MatterContext): MatterPage | null {
  if (!ctx.sections.contents) return null;
  return {
    role: 'frontMatter',
    paragraphs: [
      para(ctx, 'chapterTitle', 'Contents', { breakBefore: true }),
      contentsField(ctx),
    ],
  };
}

/** The sections that close the book. */
export function buildBackMatter(ctx: MatterContext): Element[] {
  const { details, sections } = ctx;
  const out: Element[] = [];
  if (sections.acknowledgments && sectionHasContent('acknowledgments', details)) {
    out.push(...headedSection(ctx, 'Acknowledgments', details.acknowledgments));
  }
  if (sections.aboutTheAuthor && sectionHasContent('aboutTheAuthor', details)) {
    out.push(...headedSection(ctx, 'About the Author', details.aboutTheAuthor));
  }
  if (sections.bibliography && sectionHasContent('bibliography', details)) {
    out.push(...headedSection(ctx, 'Bibliography', details.bibliography));
  }
  return out;
}

/** True when a contents table was added and Word should offer to fill it in. */
export function needsFieldUpdate(sections: ExtraSections): boolean {
  return sections.contents;
}
