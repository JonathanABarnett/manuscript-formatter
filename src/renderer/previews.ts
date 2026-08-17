import { twipsToInches } from '../core/ooxml/ns.js';
import { styleForRole, type RoleStyleMap } from '../core/roles.js';
import { sectionHasContent } from '../core/build/matter.js';
import {
  type BlockRole,
  type FormatOptions,
  type ManuscriptBlock,
  type ReferenceProfile,
  type StyleInfo,
  type StyleRole,
} from '../core/types.js';

/**
 * Renders a proof of each kind of page the formatter will produce, laid out at
 * the template's real trim size using the template's own styles and the
 * manuscript's own words. It is a layout proof, not a typesetting engine:
 * line breaks will differ from Word, but page proportions, margins, type
 * sizes, indents, alignment and the chapter sink are all true to the output.
 */

export interface PreviewContext {
  profile: ReferenceProfile;
  blocks: ManuscriptBlock[];
  options: FormatOptions;
  /** Rendered width of a page in CSS pixels; drives every other dimension. */
  pageWidthPx: number;
}

interface PreviewLine {
  role: BlockRole;
  text: string;
  /** A spacer paragraph carrying no text, as used for the chapter sink. */
  blank?: boolean;
  /** The paragraph holds a picture rather than text. */
  image?: boolean;
}

/**
 * Paragraphs worth drawing. A paragraph with no text still counts when it
 * carries a picture — a title page is often nothing but artwork.
 */
function isDrawable(block: ManuscriptBlock): boolean {
  return block.kind === 'paragraph' && (!block.isEmpty || block.hasImage);
}

function lineFrom(block: ManuscriptBlock, options: FormatOptions): PreviewLine {
  return {
    role: roleOf(block, options),
    text: block.text.trim(),
    image: block.isEmpty && block.hasImage,
  };
}

interface SectionSpec {
  key: string;
  label: string;
  caption: string;
  build: (ctx: PreviewContext) => PreviewLine[] | null;
}

/** Longest run of text shown per paragraph; the page clips the rest anyway. */
const MAX_CHARS = 520;

const SECTIONS: SectionSpec[] = [
  {
    key: 'titlePage',
    label: 'Your title page',
    caption: 'Your title, author name, and other opening-page text.',
    build: (ctx) => generatedTitlePage(ctx) ?? linesFor(ctx, ['frontMatterTitle', 'frontMatter'], 8),
  },
  {
    key: 'copyright',
    label: 'Your copyright page',
    caption: 'Found from the © line, ISBN, edition details, or rights wording.',
    build: (ctx) => generatedCopyrightPage(ctx) ?? linesFor(ctx, ['copyright'], 10),
  },
  {
    key: 'partTitle',
    label: 'A part title',
    caption: 'A major divider such as “Part One.”',
    build: (ctx) => linesFor(ctx, ['partTitle'], 2),
  },
  {
    key: 'chapter',
    label: 'How a chapter opens',
    caption:
      'Includes the space that moves the chapter title down the page.',
    build: buildChapterOpening,
  },
  {
    key: 'body',
    label: 'An ordinary page',
    caption: 'The way most of your book will look.',
    build: (ctx) => linesFor(ctx, ['body'], 9),
  },
  {
    key: 'subheading',
    label: 'A heading inside a chapter',
    caption: 'With the paragraphs either side of it, so you can judge the spacing.',
    build: (ctx) => aroundRole(ctx, 'subheading'),
  },
  {
    key: 'sceneBreak',
    label: 'A break between scenes',
    caption: 'The mark itself, and the paragraphs either side of it.',
    build: (ctx) => aroundRole(ctx, 'sceneBreak'),
  },
  {
    key: 'blockQuote',
    label: 'An indented quotation',
    caption: 'A longer quotation set apart from the main text.',
    build: (ctx) => aroundRole(ctx, 'blockQuote'),
  },
  {
    key: 'listItem',
    label: 'A list',
    caption: 'Bulleted and numbered lists.',
    build: (ctx) => aroundRole(ctx, 'listItem'),
  },
];

export function renderPreviews(container: HTMLElement, ctx: PreviewContext): void {
  container.replaceChildren();
  container.style.setProperty('--preview-width', `${ctx.pageWidthPx}px`);

  let rendered = 0;
  for (const section of SECTIONS) {
    const lines = section.build(ctx);
    if (!lines || lines.length === 0) continue;
    container.appendChild(card(section, lines, ctx));
    rendered++;
  }

  if (rendered === 0) {
    const note = document.createElement('p');
    note.className = 'preview-empty';
    note.textContent =
      'Nothing to preview yet. Sample pages will appear when the app can identify parts of the manuscript.';
    container.appendChild(note);
  }
}

// --- gathering real text from the manuscript --------------------------------

/** Effective role of a block, honouring any override the reviewer has set. */
function roleOf(block: ManuscriptBlock, options: FormatOptions): BlockRole {
  return options.roleOverrides[block.index] ?? block.role;
}

/**
 * When the author has asked the app to build a title or copyright page, the
 * proof must show that page rather than whatever the manuscript happens to
 * open with — otherwise the preview contradicts the finished book.
 */
function generatedTitlePage(ctx: PreviewContext): PreviewLine[] | null {
  const { bookDetails: d, extraSections: s } = ctx.options;
  if (!s.titlePage || !sectionHasContent('titlePage', d)) return null;
  const lines: PreviewLine[] = [{ role: 'frontMatterTitle', text: d.title.trim() }];
  if (d.subtitle.trim()) lines.push({ role: 'chapterSubtitle', text: d.subtitle.trim() });
  if (d.author.trim()) lines.push({ role: 'frontMatter', text: d.author.trim() });
  return lines;
}

function generatedCopyrightPage(ctx: PreviewContext): PreviewLine[] | null {
  const { bookDetails: d, extraSections: s } = ctx.options;
  if (!s.copyrightPage || !sectionHasContent('copyrightPage', d)) return null;
  const year = d.copyrightYear.trim();
  const author = d.author.trim();
  const notice =
    year && author ? `Copyright © ${year} by ${author}` : year ? `Copyright © ${year}` : `Copyright © ${author}`;
  const lines: PreviewLine[] = [
    { role: 'copyright', text: notice },
    { role: 'copyright', text: 'All rights reserved.' },
    {
      role: 'copyright',
      text:
        'No part of this book may be reproduced in any form without written permission ' +
        'from the copyright holder, except brief quotations used in a review.',
    },
  ];
  if (d.publisher.trim()) lines.push({ role: 'copyright', text: d.publisher.trim() });
  if (d.isbn.trim()) lines.push({ role: 'copyright', text: `ISBN: ${d.isbn.trim()}` });
  return lines;
}

/** The first `limit` blocks whose role is one of `roles`, in document order. */
function linesFor(ctx: PreviewContext, roles: BlockRole[], limit: number): PreviewLine[] | null {
  const wanted = new Set(roles);
  const out: PreviewLine[] = [];
  for (const block of ctx.blocks) {
    if (!isDrawable(block)) continue;
    if (!wanted.has(roleOf(block, ctx.options))) continue;
    out.push(lineFrom(block, ctx.options));
    if (out.length >= limit) break;
  }
  return out.length > 0 ? out : null;
}

/** A target paragraph with its neighbours, so the spacing reads in context. */
function aroundRole(ctx: PreviewContext, role: BlockRole): PreviewLine[] | null {
  const index = ctx.blocks.findIndex(
    (b) => isDrawable(b) && roleOf(b, ctx.options) === role,
  );
  if (index === -1) return null;

  return [
    ...contentNear(ctx, index, -1, 2).reverse(),
    lineFrom(ctx.blocks[index], ctx.options),
    ...contentNear(ctx, index, 1, role === 'listItem' ? 3 : 2),
  ];
}

/** Walk outward from `from`, collecting up to `count` paragraphs with text. */
function contentNear(
  ctx: PreviewContext,
  from: number,
  step: number,
  count: number,
): PreviewLine[] {
  const out: PreviewLine[] = [];
  for (let i = from + step; i >= 0 && i < ctx.blocks.length && out.length < count; i += step) {
    const block = ctx.blocks[i];
    if (!isDrawable(block)) continue;
    const role = roleOf(block, ctx.options);
    if (role === 'empty' || role === 'pageBreak') continue;
    out.push(lineFrom(block, ctx.options));
  }
  return out;
}

/** The chapter sink, the title, any subtitle, then the opening paragraphs. */
function buildChapterOpening(ctx: PreviewContext): PreviewLine[] | null {
  const index = ctx.blocks.findIndex((b) => {
    const role = roleOf(b, ctx.options);
    return isDrawable(b) && (role === 'chapterTitle' || role === 'partTitle');
  });
  if (index === -1) return null;

  const before = Math.max(
    0,
    ctx.options.chapterSpaceBefore ?? ctx.profile.chapterTitleBlanksBefore,
  );
  const after = Math.max(0, ctx.options.chapterSpaceAfter ?? ctx.profile.chapterTitleBlanksAfter);

  const lines: PreviewLine[] = [];
  for (let i = 0; i < before; i++) lines.push({ role: 'chapterTitle', text: '', blank: true });
  lines.push(lineFrom(ctx.blocks[index], ctx.options));
  for (let i = 0; i < after; i++) lines.push({ role: 'chapterTitle', text: '', blank: true });
  lines.push(...contentNear(ctx, index, 1, 5));
  return lines;
}

// --- drawing ----------------------------------------------------------------

function card(section: SectionSpec, lines: PreviewLine[], ctx: PreviewContext): HTMLElement {
  const wrapper = document.createElement('div');
  wrapper.className = 'preview-card';

  const head = document.createElement('div');
  head.className = 'preview-head';
  const title = document.createElement('span');
  title.className = 'preview-title';
  title.textContent = section.label;
  head.appendChild(title);
  wrapper.appendChild(head);

  wrapper.appendChild(page(lines, ctx));

  const caption = document.createElement('p');
  caption.className = 'preview-caption';
  caption.textContent = section.caption;
  wrapper.appendChild(caption);
  return wrapper;
}

function page(lines: PreviewLine[], ctx: PreviewContext): HTMLElement {
  const setup = ctx.profile.pageSetup;
  const widthIn = twipsToInches(setup.widthTwips);
  const heightIn = twipsToInches(setup.heightTwips);
  const pxPerInch = ctx.pageWidthPx / widthIn;

  const sheet = document.createElement('div');
  sheet.className = 'preview-page';
  // Both dimensions are set explicitly so the sheet keeps the trim ratio no
  // matter how wide its grid column happens to be.
  sheet.style.width = `${Math.round(ctx.pageWidthPx)}px`;
  sheet.style.height = `${Math.round(heightIn * pxPerInch)}px`;

  const inner = document.createElement('div');
  inner.className = 'preview-sheet';
  const m = setup.margins;
  // A gutter is bound into the inside edge, which is the left of a recto page.
  const left = (m.left + m.gutter) / 1440;
  inner.style.padding = `${(m.top / 1440) * pxPerInch}px ${(m.right / 1440) * pxPerInch}px ${
    (m.bottom / 1440) * pxPerInch
  }px ${left * pxPerInch}px`;

  const guide = document.createElement('div');
  guide.className = 'preview-guide';
  guide.style.inset = inner.style.padding;
  sheet.appendChild(guide);

  // Templates park a copyright notice at the foot of its page and centre a
  // title page, using section properties. Show that, or the proof lies.
  const pageRole = lines.find((l) => !l.blank)?.role;
  const vAlign = pageRole
    ? (ctx.profile.roleVAlign[pageRole as StyleRole] ??
      (pageRole === 'copyright' ? 'bottom' : undefined))
    : undefined;
  if (vAlign === 'bottom' || vAlign === 'center') {
    inner.style.display = 'flex';
    inner.style.flexDirection = 'column';
    inner.style.justifyContent = vAlign === 'bottom' ? 'flex-end' : 'center';
  }

  const roles = effectiveRoleStyles(ctx);
  for (const line of lines) {
    inner.appendChild(paragraph(line, roles, ctx, pxPerInch));
  }
  sheet.appendChild(inner);
  return sheet;
}

/** Detected role→style map with the reviewer's overrides applied. */
function effectiveRoleStyles(ctx: PreviewContext): RoleStyleMap {
  const merged = { ...ctx.profile.roleStyles };
  for (const [role, id] of Object.entries(ctx.options.roleStyles)) {
    if (id !== undefined) merged[role as StyleRole] = id;
  }
  return merged;
}

function paragraph(
  line: PreviewLine,
  roles: RoleStyleMap,
  ctx: PreviewContext,
  pxPerInch: number,
): HTMLElement {
  const styleId = styleForRole(line.role, roles);
  const info = ctx.profile.styles.find((s) => s.id === styleId) ?? null;

  const node = document.createElement('p');
  applyStyle(node, info, pxPerInch);
  applyRoleAdjustments(node, line.role, styleId, roles, ctx, pxPerInch);

  if (line.blank) {
    node.innerHTML = '&nbsp;';
    return node;
  }
  if (line.image) {
    node.className = 'preview-image';
    node.textContent = line.text || 'image';
    return node;
  }
  const text = line.text.length > MAX_CHARS ? `${line.text.slice(0, MAX_CHARS)}…` : line.text;
  node.textContent = text;
  return node;
}

const ALIGNMENT: Record<string, string> = {
  both: 'justify',
  distribute: 'justify',
  center: 'center',
  right: 'right',
  end: 'right',
  left: 'left',
  start: 'left',
};

function applyStyle(node: HTMLElement, info: StyleInfo | null, pxPerInch: number): void {
  const pt = (points: number): string => `${(points / 72) * pxPerInch}px`;
  const tw = (twips: number): string => `${(twips / 1440) * pxPerInch}px`;

  node.style.fontFamily = info?.fontName
    ? `"${info.fontName}", "Iowan Old Style", Georgia, serif`
    : '"Iowan Old Style", Georgia, serif';
  node.style.fontSize = pt(info?.fontSizePt ?? 11);
  node.style.textAlign = ALIGNMENT[info?.alignment ?? 'left'] ?? 'left';
  node.style.fontWeight = info?.bold ? '700' : '400';
  node.style.fontStyle = info?.italic ? 'italic' : 'normal';
  if (info?.allCaps) node.style.textTransform = 'uppercase';
  if (info?.smallCaps) node.style.fontVariantCaps = 'small-caps';

  const left = info?.leftIndentTwips ?? 0;
  if (left) node.style.marginLeft = tw(left);
  const first = info?.firstLineIndentTwips ?? 0;
  const hanging = info?.hangingIndentTwips ?? 0;
  if (hanging) node.style.textIndent = `-${(hanging / 1440) * pxPerInch}px`;
  else if (first) node.style.textIndent = tw(first);

  if (info?.spaceBeforeTwips) node.style.marginTop = tw(info.spaceBeforeTwips);
  if (info?.spaceAfterTwips) node.style.marginBottom = tw(info.spaceAfterTwips);

  // `auto` line spacing is in 240ths of a line; the other rules are twips.
  if (info?.lineSpacing) {
    node.style.lineHeight =
      info.lineRule === 'exact' || info.lineRule === 'atLeast'
        ? tw(info.lineSpacing)
        : String(info.lineSpacing / 240);
  } else {
    node.style.lineHeight = '1.3';
  }
}

/**
 * Mirrors the direct formatting the composer adds when the template has no
 * dedicated style for a role, so the proof matches the document.
 */
function applyRoleAdjustments(
  node: HTMLElement,
  role: BlockRole,
  styleId: string | null,
  roles: RoleStyleMap,
  ctx: PreviewContext,
  pxPerInch: number,
): void {
  const fellBackToBody = styleId !== null && styleId === roles.body;

  if (fellBackToBody) {
    if (role === 'chapterTitle' || role === 'partTitle' || role === 'frontMatterTitle') {
      node.style.textAlign = 'center';
      node.style.fontWeight = '700';
      node.style.textIndent = '0';
      node.style.marginTop = `${(480 / 1440) * pxPerInch}px`;
      node.style.marginBottom = `${(360 / 1440) * pxPerInch}px`;
    } else if (role === 'chapterSubtitle' || role === 'subheading') {
      node.style.fontWeight = '700';
      node.style.textIndent = '0';
      if (role === 'chapterSubtitle') node.style.textAlign = 'center';
    } else if (role === 'sceneBreak') {
      node.style.textAlign = 'center';
      node.style.textIndent = '0';
      node.style.marginTop = `${(240 / 1440) * pxPerInch}px`;
      node.style.marginBottom = `${(240 / 1440) * pxPerInch}px`;
    } else if (role === 'blockQuote') {
      node.style.marginLeft = `${(720 / 1440) * pxPerInch}px`;
      node.style.marginRight = `${(720 / 1440) * pxPerInch}px`;
      node.style.textIndent = '0';
    }
  }

  // The composer flushes an opening paragraph left when the template has no
  // dedicated first-paragraph style to do it.
  if (
    role === 'bodyFirst' &&
    ctx.options.firstParagraphNoIndent &&
    (roles.bodyFirst ?? roles.body) === roles.body
  ) {
    node.style.textIndent = '0';
  }

  if (role === 'sceneBreak' && ctx.options.sceneBreakText) {
    node.textContent = ctx.options.sceneBreakText;
  }
}
