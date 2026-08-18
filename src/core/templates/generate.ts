import JSZip from 'jszip';
import { NS } from '../ooxml/ns.js';
import {
  GUTTER_IN,
  findLook,
  findTrim,
  type BookLook,
  type TrimSize,
} from './design.js';

/**
 * Builds a book design as a real .docx, in memory. The result is fed straight
 * into the same analyzer that reads a template an author supplies, so a
 * built-in design and a KDP template travel identical code paths.
 */

const TW = (inches: number): number => Math.round(inches * 1440);
const HALF_PT = (points: number): number => Math.round(points * 2);
/** Word stores `auto` line spacing in 240ths of a line. */
const LINE = (multiple: number): number => Math.round(multiple * 240);

function esc(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

interface StyleSpec {
  id: string;
  name: string;
  basedOn?: string;
  next?: string;
  isDefault?: boolean;
  font?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  align?: string;
  firstLineIn?: number;
  leftIn?: number;
  rightIn?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: number;
  keepNext?: boolean;
  outlineLevel?: number;
}

function styleXml(s: StyleSpec): string {
  const pPr: string[] = [];
  if (s.keepNext) pPr.push('<w:keepNext/>');
  if (s.spaceBeforePt !== undefined || s.spaceAfterPt !== undefined || s.lineSpacing !== undefined) {
    const parts = [
      s.spaceBeforePt !== undefined ? ` w:before="${Math.round(s.spaceBeforePt * 20)}"` : '',
      s.spaceAfterPt !== undefined ? ` w:after="${Math.round(s.spaceAfterPt * 20)}"` : '',
      s.lineSpacing !== undefined ? ` w:line="${LINE(s.lineSpacing)}" w:lineRule="auto"` : '',
    ].join('');
    pPr.push(`<w:spacing${parts}/>`);
  }
  if (s.firstLineIn !== undefined || s.leftIn !== undefined || s.rightIn !== undefined) {
    const parts = [
      s.leftIn !== undefined ? ` w:left="${TW(s.leftIn)}"` : '',
      s.rightIn !== undefined ? ` w:right="${TW(s.rightIn)}"` : '',
      s.firstLineIn !== undefined ? ` w:firstLine="${TW(s.firstLineIn)}"` : '',
    ].join('');
    pPr.push(`<w:ind${parts}/>`);
  }
  if (s.align) pPr.push(`<w:jc w:val="${s.align}"/>`);
  if (s.outlineLevel !== undefined) pPr.push(`<w:outlineLvl w:val="${s.outlineLevel}"/>`);

  const rPr: string[] = [];
  if (s.font) rPr.push(`<w:rFonts w:ascii="${esc(s.font)}" w:hAnsi="${esc(s.font)}" w:cs="${esc(s.font)}"/>`);
  if (s.bold) rPr.push('<w:b/>');
  if (s.italic) rPr.push('<w:i/>');
  if (s.caps) rPr.push('<w:caps/>');
  if (s.sizePt) rPr.push(`<w:sz w:val="${HALF_PT(s.sizePt)}"/><w:szCs w:val="${HALF_PT(s.sizePt)}"/>`);

  return (
    `<w:style w:type="paragraph"${s.isDefault ? ' w:default="1"' : ''} w:styleId="${esc(s.id)}">` +
    `<w:name w:val="${esc(s.name)}"/>` +
    (s.basedOn ? `<w:basedOn w:val="${esc(s.basedOn)}"/>` : '') +
    (s.next ? `<w:next w:val="${esc(s.next)}"/>` : '') +
    '<w:qFormat/>' +
    (pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : '') +
    (rPr.length ? `<w:rPr>${rPr.join('')}</w:rPr>` : '') +
    '</w:style>'
  );
}

/** The named designs each part of a book maps onto. */
function stylesFor(look: BookLook): StyleSpec[] {
  return [
    {
      id: 'Normal',
      name: 'Normal',
      isDefault: true,
      font: look.bodyFont,
      sizePt: look.bodySizePt,
      lineSpacing: look.lineSpacing,
      spaceAfterPt: 0,
    },
    {
      id: 'BodyText',
      name: 'Body Text',
      basedOn: 'Normal',
      next: 'BodyText',
      firstLineIn: look.indentIn,
      align: look.justified ? 'both' : 'left',
    },
    {
      id: 'FirstParagraph',
      name: 'First Paragraph',
      basedOn: 'BodyText',
      next: 'BodyText',
      firstLineIn: 0,
    },
    {
      id: 'ChapterTitle',
      name: 'Chapter Title',
      basedOn: 'Normal',
      next: 'FirstParagraph',
      font: look.headingFont,
      sizePt: look.chapterSizePt,
      bold: look.chapterBold,
      caps: look.chapterCaps,
      align: look.chapterAlign,
      firstLineIn: 0,
      spaceAfterPt: 18,
      keepNext: true,
      outlineLevel: 0,
    },
    {
      id: 'PartTitle',
      name: 'Part Title',
      basedOn: 'ChapterTitle',
      next: 'FirstParagraph',
      sizePt: look.chapterSizePt + 4,
      align: 'center',
    },
    {
      id: 'Subhead',
      name: 'Heading 2',
      basedOn: 'Normal',
      next: 'FirstParagraph',
      font: look.headingFont,
      sizePt: look.bodySizePt + 1,
      bold: true,
      align: 'left',
      firstLineIn: 0,
      spaceBeforePt: 14,
      spaceAfterPt: 6,
      keepNext: true,
      outlineLevel: 1,
    },
    {
      id: 'SceneBreak',
      name: 'Scene Break',
      basedOn: 'Normal',
      next: 'FirstParagraph',
      align: 'center',
      firstLineIn: 0,
      spaceBeforePt: 12,
      spaceAfterPt: 12,
    },
    {
      id: 'BlockQuote',
      name: 'Block Quote',
      basedOn: 'Normal',
      next: 'BodyText',
      leftIn: 0.35,
      rightIn: 0.35,
      firstLineIn: 0,
      sizePt: look.bodySizePt - 0.5,
      spaceBeforePt: 8,
      spaceAfterPt: 8,
    },
    {
      id: 'BookTitle',
      name: 'Title',
      basedOn: 'Normal',
      font: look.headingFont,
      sizePt: look.chapterSizePt + 12,
      align: 'center',
      firstLineIn: 0,
      spaceAfterPt: 12,
    },
    {
      id: 'BookSubtitle',
      name: 'Subtitle',
      basedOn: 'Normal',
      font: look.headingFont,
      sizePt: look.chapterSizePt,
      align: 'center',
      firstLineIn: 0,
      spaceAfterPt: 24,
    },
    {
      id: 'AuthorName',
      name: 'Author Name',
      basedOn: 'Normal',
      font: look.headingFont,
      sizePt: look.bodySizePt + 3,
      align: 'center',
      firstLineIn: 0,
    },
    {
      id: 'CopyrightPage',
      name: 'Copyright Page',
      basedOn: 'Normal',
      sizePt: look.bodySizePt - 1.5,
      align: 'center',
      firstLineIn: 0,
      spaceAfterPt: 6,
      lineSpacing: 1,
    },
    {
      id: 'FrontMatterText',
      name: 'Front Matter Text',
      basedOn: 'Normal',
      align: 'center',
      firstLineIn: 0,
      spaceAfterPt: 8,
    },
  ];
}

function sectPr(trim: TrimSize, headerRefs: string, footerRefs: string): string {
  const m = trim.margins;
  return (
    headerRefs +
    footerRefs +
    `<w:pgSz w:w="${TW(trim.widthIn)}" w:h="${TW(trim.heightIn)}"/>` +
    `<w:pgMar w:top="${TW(m.top)}" w:right="${TW(m.outside)}" w:bottom="${TW(m.bottom)}"` +
    ` w:left="${TW(m.inside)}" w:header="${TW(Math.max(0.3, m.top - 0.35))}"` +
    ` w:footer="${TW(Math.max(0.3, m.bottom - 0.35))}" w:gutter="${TW(GUTTER_IN)}"/>` +
    '<w:pgNumType w:start="1"/>' +
    // No folio on a chapter's opening page, as printed books set them.
    '<w:titlePg/>' +
    '<w:docGrid w:linePitch="360"/>'
  );
}

/** A centred page number. Word evaluates the PAGE field when the file opens. */
function footerXml(look: BookLook): string {
  return (
    XML_DECL +
    `<w:ftr xmlns:w="${NS.w}"><w:p><w:pPr><w:jc w:val="center"/>` +
    `<w:rPr><w:rFonts w:ascii="${esc(look.bodyFont)}" w:hAnsi="${esc(look.bodyFont)}"/>` +
    `<w:sz w:val="${HALF_PT(look.bodySizePt - 1)}"/></w:rPr></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="${esc(look.bodyFont)}" w:hAnsi="${esc(look.bodyFont)}"/>` +
    `<w:sz w:val="${HALF_PT(look.bodySizePt - 1)}"/></w:rPr>` +
    '<w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>1</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:p></w:ftr>'
  );
}

interface Para {
  style: string;
  text?: string;
  sectBreak?: boolean;
}

function paraXml(p: Para, trim: TrimSize, refs: { header: string; footer: string }): string {
  const inner = p.sectBreak ? `<w:sectPr>${sectPr(trim, refs.header, refs.footer)}</w:sectPr>` : '';
  const pPr = `<w:pPr><w:pStyle w:val="${esc(p.style)}"/>${inner}</w:pPr>`;
  const run = p.text ? `<w:r><w:t xml:space="preserve">${esc(p.text)}</w:t></w:r>` : '';
  return `<w:p>${pPr}${run}</w:p>`;
}

/**
 * The template's own body. It carries no placeholder prose: what it holds is
 * the *shape* of a book — a chapter opening with its blank sink lines, one
 * scene break, one subhead — which is exactly what the analyzer reads to learn
 * how this design lays a chapter out.
 */
function templateBody(look: BookLook, trim: TrimSize, refs: { header: string; footer: string }): string {
  const paras: Para[] = [];
  const blanks = (style: string, count: number): void => {
    for (let i = 0; i < count; i++) paras.push({ style });
  };

  // A chapter opening, twice, so the sink is measurable rather than guessed.
  for (const title of ['Chapter One', 'Chapter Two']) {
    blanks('ChapterTitle', look.chapterBlanksBefore);
    paras.push({ style: 'ChapterTitle', text: title });
    blanks('ChapterTitle', look.chapterBlanksAfter);
    paras.push({ style: 'FirstParagraph', text: SPECIMEN_OPENING });
    paras.push({ style: 'BodyText', text: SPECIMEN_BODY });
    paras.push({ style: 'SceneBreak', text: look.sceneMark });
    paras.push({ style: 'FirstParagraph', text: SPECIMEN_AFTER_BREAK });
    paras.push({ style: 'BodyText', text: SPECIMEN_BODY_2 });
    paras.push({ style: 'Subhead', text: 'A heading inside a chapter' });
    paras.push({ style: 'FirstParagraph', text: SPECIMEN_AFTER_BREAK });
    paras.push({ style: 'BodyText', text: SPECIMEN_BODY });
  }

  return paras.map((p) => paraXml(p, trim, refs)).join('');
}

/**
 * Specimen sentences describing the design itself. They exist so the analyzer
 * has prose to measure and are replaced wholesale by the author's own words,
 * so nothing here can survive into a finished book.
 */
const SPECIMEN_OPENING =
  'This first paragraph of a chapter sits flush against the margin, with no indent, ' +
  'which is the convention printed books follow. Your own opening line replaces it.';
const SPECIMEN_BODY =
  'Every paragraph after the first is indented and set in the body face chosen for ' +
  'this design, so that the page has an even colour from top to bottom. Nothing of ' +
  'this specimen text reaches your finished book; it is here only so that the page ' +
  'proportions and spacing can be measured accurately before your words arrive.';
const SPECIMEN_AFTER_BREAK =
  'After a break the text resumes without an indent, exactly as it does at the start ' +
  'of a chapter, and then continues in the usual way until the section ends.';
const SPECIMEN_BODY_2 =
  'A second paragraph keeps the specimen honest and gives the analyzer enough running ' +
  'text to work out which design carries the bulk of the book rather than its headings.';

/** Build a book design as a .docx. */
export async function buildTemplate(trimId: string, lookId: string): Promise<Uint8Array> {
  const trim = findTrim(trimId);
  const look = findLook(lookId);
  const zip = new JSZip();

  const refs = {
    header: '',
    footer: '<w:footerReference w:type="default" r:id="rIdFooter"/>',
  };

  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      `<Types xmlns="${NS.contentTypes}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '</Types>',
  );

  zip.file(
    '_rels/.rels',
    XML_DECL +
      `<Relationships xmlns="${NS.pkgRel}">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  );

  zip.file(
    'word/_rels/document.xml.rels',
    XML_DECL +
      `<Relationships xmlns="${NS.pkgRel}">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
      '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
      '</Relationships>',
  );

  zip.file(
    'word/styles.xml',
    XML_DECL +
      `<w:styles xmlns:w="${NS.w}">` +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      `<w:rFonts w:ascii="${esc(look.bodyFont)}" w:hAnsi="${esc(look.bodyFont)}"/>` +
      `<w:sz w:val="${HALF_PT(look.bodySizePt)}"/>` +
      '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
      `<w:spacing w:after="0" w:line="${LINE(look.lineSpacing)}" w:lineRule="auto"/>` +
      '</w:pPr></w:pPrDefault></w:docDefaults>' +
      stylesFor(look).map(styleXml).join('') +
      '</w:styles>',
  );

  // Mirrored margins make the gutter swap sides on facing pages. Justified
  // text is hyphenated so the word spacing stays even, with capitals left
  // whole and no more than two hyphenated line ends in a row.
  zip.file(
    'word/settings.xml',
    XML_DECL +
      `<w:settings xmlns:w="${NS.w}"><w:mirrorMargins/><w:defaultTabStop w:val="720"/>` +
      (look.justified
        ? '<w:autoHyphenation/><w:consecutiveHyphenLimit w:val="2"/><w:doNotHyphenateCaps/>'
        : '') +
      '<w:characterSpacingControl w:val="compressPunctuation"/></w:settings>',
  );

  zip.file('word/footer1.xml', footerXml(look));

  zip.file(
    'word/document.xml',
    XML_DECL +
      `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body>` +
      templateBody(look, trim, refs) +
      `<w:sectPr>${sectPr(trim, refs.header, refs.footer)}</w:sectPr>` +
      '</w:body></w:document>',
  );

  const buffer = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  return buffer;
}

/**
 * A short manuscript in the shape authors actually deliver — no styles, tabbed
 * indents, capitalised chapter headings, a hash for a scene break — so that
 * "try it with a sample book" exercises the real detection path.
 */
export async function buildSampleManuscript(): Promise<Uint8Array> {
  const zip = new JSZip();
  const p = (text: string, opts: { center?: boolean; bold?: boolean; sizePt?: number; break?: boolean } = {}): string => {
    const pPr: string[] = [];
    if (opts.break) pPr.push('<w:pageBreakBefore/>');
    if (opts.center) pPr.push('<w:jc w:val="center"/>');
    const rPr: string[] = ['<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'];
    if (opts.bold) rPr.push('<w:b/>');
    rPr.push(`<w:sz w:val="${HALF_PT(opts.sizePt ?? 12)}"/>`);
    return (
      `<w:p>${pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}` +
      `<w:r><w:rPr>${rPr.join('')}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`
    );
  };

  const body: string[] = [
    p('The Cartographer of Small Hours', { center: true, bold: true, sizePt: 22 }),
    p('by A. N. Author', { center: true, sizePt: 14 }),
    p('Copyright © 2026 by A. N. Author', { center: true, break: true }),
    p('All rights reserved. No part of this book may be reproduced without permission.', { center: true }),
    p('This is a work of fiction. Any resemblance to real persons is coincidental.', { center: true }),
    p('First edition.', { center: true }),
  ];

  const chapters: Array<[string, string[]]> = [
    [
      'CHAPTER ONE',
      [
        'The morning came in grey and unhurried, the way mornings did in that part of the country, and Hollis had long since stopped expecting anything else of them.',
        'He pushed the shop door open with his shoulder, because the handle had never worked properly, and the small bell above it made its usual complaint about being disturbed so early in the day.',
        '#',
        'By nine the rain had settled into the steady, unremarkable sort that could go on for a week without anyone remarking on it.',
        'A woman came in asking for a map of somewhere that no longer existed, and Hollis, who kept several, did not ask her why she wanted it.',
      ],
    ],
    [
      'CHAPTER TWO',
      [
        'The corridor above the shop smelled of wet wool and yesterday rain, and the notice board at the far end had not been changed since the spring.',
        'Three announcements curled away from their pins. A fourth had been torn neatly in half and left there, which Hollis found more interesting than the other three together.',
        '#',
        'Outside, past the glass, the square was emptying in the way squares do when the light goes thin and everyone decides at once that the afternoon has ended.',
        'He turned the sign, put the kettle on, and began, without much hope, to look for the half that was missing.',
      ],
    ],
  ];

  for (const [title, paras] of chapters) {
    body.push(p(title, { center: true, bold: true, sizePt: 14, break: true }));
    body.push(p(''));
    for (const text of paras) {
      if (text === '#') body.push(p('#', { center: true }));
      else body.push(p(`\t${text}`));
    }
  }

  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      `<Types xmlns="${NS.contentTypes}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>',
  );
  zip.file(
    '_rels/.rels',
    XML_DECL +
      `<Relationships xmlns="${NS.pkgRel}">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    XML_DECL +
      `<Relationships xmlns="${NS.pkgRel}">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  );
  zip.file(
    'word/styles.xml',
    XML_DECL +
      `<w:styles xmlns:w="${NS.w}">` +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '</w:styles>',
  );
  zip.file(
    'word/document.xml',
    XML_DECL +
      `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body>` +
      body.join('') +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
      '</w:body></w:document>',
  );

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
