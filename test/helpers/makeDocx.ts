import { readFile, writeFile } from 'node:fs/promises';
import { basename } from 'node:path';
import JSZip from 'jszip';
import { DocxPackage } from '../../src/core/ooxml/package.js';
import type { DocxInput } from '../../src/core/types.js';

/** Minimal .docx builder for fixtures — just enough OOXML to be a real file. */

export interface StyleSpec {
  id: string;
  name: string;
  type?: 'paragraph' | 'character';
  basedOn?: string;
  next?: string;
  isDefault?: boolean;
  outlineLevel?: number;
  fontSizePt?: number;
  fontName?: string;
  alignment?: string;
  firstLineIndent?: number;
  leftIndent?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  bold?: boolean;
  italic?: boolean;
}

export interface RunSpec {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  smallCaps?: boolean;
  superscript?: boolean;
  fontSizePt?: number;
  color?: string;
  fontName?: string;
  rStyle?: string;
}

export interface ParaSpec {
  text?: string;
  runs?: RunSpec[];
  /** Append a run carrying a reference to this footnote id. */
  footnoteRef?: number;
  /** Append a run holding an inline picture of the packaged test image. */
  image?: boolean;
  style?: string;
  alignment?: string;
  firstLineIndent?: number;
  leftIndent?: number;
  pageBreakBefore?: boolean;
  /** Emit `<w:br w:type="page"/>` before the text. */
  leadingPageBreak?: boolean;
  numId?: number;
  ilvl?: number;
  outlineLevel?: number;
  /** Raw `<w:sectPr>` inner XML attached to this paragraph. */
  sectPr?: string;
}

export interface DocSpec {
  styles?: StyleSpec[];
  paragraphs: ParaSpec[];
  /** Inner XML of the body-level `<w:sectPr>`. */
  sectPr?: string;
  numbering?: boolean;
  /** Footnote bodies to package, keyed by the id paragraphs refer to. */
  footnotes?: Array<{ id: number; text: string }>;
  /** Package a 1x1 PNG at word/media/test.png for `image` paragraphs. */
  image?: boolean;
}

/** Smallest valid PNG, so image migration has real bytes to move. */
export const TEST_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const LETTER_SECTPR =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

/** 5.5 x 8.5 inch trim with mirrored book margins and a gutter. */
export const DIGEST_SECTPR =
  '<w:pgSz w:w="7920" w:h="12240"/>' +
  '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="360"/>' +
  '<w:pgNumType w:start="1"/>' +
  '<w:titlePg/>';

export function esc(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function styleXml(s: StyleSpec): string {
  const pPr: string[] = [];
  if (s.keepNext) pPr.push('<w:keepNext/>');
  if (s.pageBreakBefore) pPr.push('<w:pageBreakBefore/>');
  if (s.spaceBefore !== undefined || s.spaceAfter !== undefined) {
    pPr.push(
      `<w:spacing${s.spaceBefore !== undefined ? ` w:before="${s.spaceBefore}"` : ''}` +
        `${s.spaceAfter !== undefined ? ` w:after="${s.spaceAfter}"` : ''}/>`,
    );
  }
  if (s.firstLineIndent !== undefined || s.leftIndent !== undefined) {
    pPr.push(
      `<w:ind${s.leftIndent !== undefined ? ` w:left="${s.leftIndent}"` : ''}` +
        `${s.firstLineIndent !== undefined ? ` w:firstLine="${s.firstLineIndent}"` : ''}/>`,
    );
  }
  if (s.alignment) pPr.push(`<w:jc w:val="${s.alignment}"/>`);
  if (s.outlineLevel !== undefined) pPr.push(`<w:outlineLvl w:val="${s.outlineLevel}"/>`);

  const rPr: string[] = [];
  if (s.fontName) rPr.push(`<w:rFonts w:ascii="${esc(s.fontName)}" w:hAnsi="${esc(s.fontName)}"/>`);
  if (s.bold) rPr.push('<w:b/>');
  if (s.italic) rPr.push('<w:i/>');
  if (s.fontSizePt) rPr.push(`<w:sz w:val="${Math.round(s.fontSizePt * 2)}"/>`);

  return (
    `<w:style w:type="${s.type ?? 'paragraph'}"${s.isDefault ? ' w:default="1"' : ''} w:styleId="${esc(s.id)}">` +
    `<w:name w:val="${esc(s.name)}"/>` +
    (s.basedOn ? `<w:basedOn w:val="${esc(s.basedOn)}"/>` : '') +
    (s.next ? `<w:next w:val="${esc(s.next)}"/>` : '') +
    (pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : '') +
    (rPr.length ? `<w:rPr>${rPr.join('')}</w:rPr>` : '') +
    '</w:style>'
  );
}

function runXml(r: RunSpec): string {
  const rPr: string[] = [];
  if (r.rStyle) rPr.push(`<w:rStyle w:val="${esc(r.rStyle)}"/>`);
  if (r.fontName) rPr.push(`<w:rFonts w:ascii="${esc(r.fontName)}" w:hAnsi="${esc(r.fontName)}"/>`);
  if (r.bold) rPr.push('<w:b/>');
  if (r.italic) rPr.push('<w:i/>');
  if (r.smallCaps) rPr.push('<w:smallCaps/>');
  if (r.color) rPr.push(`<w:color w:val="${esc(r.color)}"/>`);
  if (r.fontSizePt) rPr.push(`<w:sz w:val="${Math.round(r.fontSizePt * 2)}"/>`);
  if (r.underline) rPr.push('<w:u w:val="single"/>');
  if (r.superscript) rPr.push('<w:vertAlign w:val="superscript"/>');
  const props = rPr.length ? `<w:rPr>${rPr.join('')}</w:rPr>` : '';
  return `<w:r>${props}<w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`;
}

function paraXml(p: ParaSpec): string {
  const pPr: string[] = [];
  if (p.style) pPr.push(`<w:pStyle w:val="${esc(p.style)}"/>`);
  if (p.pageBreakBefore) pPr.push('<w:pageBreakBefore/>');
  if (p.numId !== undefined) {
    pPr.push(`<w:numPr><w:ilvl w:val="${p.ilvl ?? 0}"/><w:numId w:val="${p.numId}"/></w:numPr>`);
  }
  if (p.firstLineIndent !== undefined || p.leftIndent !== undefined) {
    pPr.push(
      `<w:ind${p.leftIndent !== undefined ? ` w:left="${p.leftIndent}"` : ''}` +
        `${p.firstLineIndent !== undefined ? ` w:firstLine="${p.firstLineIndent}"` : ''}/>`,
    );
  }
  if (p.alignment) pPr.push(`<w:jc w:val="${p.alignment}"/>`);
  if (p.outlineLevel !== undefined) pPr.push(`<w:outlineLvl w:val="${p.outlineLevel}"/>`);
  if (p.sectPr) pPr.push(`<w:sectPr>${p.sectPr}</w:sectPr>`);

  const lead = p.leadingPageBreak ? '<w:r><w:br w:type="page"/></w:r>' : '';
  const runs = p.runs ?? (p.text !== undefined && p.text !== '' ? [{ text: p.text }] : []);
  const body = runs.map(runXml).join('');
  const note =
    p.footnoteRef !== undefined
      ? '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr>' +
        `<w:footnoteReference w:id="${p.footnoteRef}"/></w:r>`
      : '';
  const picture = p.image ? IMAGE_RUN : '';
  const props = pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : '';
  return `<w:p>${props}${lead}${body}${note}${picture}</w:p>`;
}

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

/** An inline picture pointing at the packaged media relationship. */
const IMAGE_RUN =
  '<w:r><w:drawing>' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>' +
  `<a:graphic><a:graphicData uri="${PIC}">` +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="test.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline>' +
  '</w:drawing></w:r>';

function footnotesXml(notes: Array<{ id: number; text: string }>): string {
  const fixed =
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  const body = notes
    .map(
      (n) =>
        `<w:footnote w:id="${n.id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
        `<w:r><w:t xml:space="preserve"> ${esc(n.text)}</w:t></w:r></w:p></w:footnote>`,
    )
    .join('');
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${W}">` +
    fixed +
    body +
    '</w:footnotes>'
  );
}

const NUMBERING_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}">` +
  '<w:abstractNum w:abstractNumId="0"><w:nsid w:val="11111111"/><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>';

export async function buildDocx(spec: DocSpec): Promise<Buffer> {
  const zip = new JSZip();

  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
  ];
  if (spec.numbering) {
    overrides.push(
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    );
  }
  if (spec.footnotes) {
    overrides.push(
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    );
  }

  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      (spec.image ? '<Default Extension="png" ContentType="image/png"/>' : '') +
      overrides.join('') +
      '</Types>',
  );

  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  );

  const docRels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
  ];
  if (spec.numbering) {
    docRels.push(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    );
  }
  if (spec.footnotes) {
    docRels.push(
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
    );
    zip.file('word/footnotes.xml', footnotesXml(spec.footnotes));
  }
  if (spec.image) {
    docRels.push(
      '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/test.png"/>',
    );
    zip.file('word/media/test.png', TEST_PNG);
  }
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      docRels.join('') +
      '</Relationships>',
  );

  const styles = spec.styles ?? [{ id: 'Normal', name: 'Normal', isDefault: true }];
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}">` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      styles.map(styleXml).join('') +
      '</w:styles>',
  );

  if (spec.numbering) zip.file('word/numbering.xml', NUMBERING_XML);

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:a="${A}" xmlns:wp="${WP}" xmlns:pic="${PIC}"><w:body>` +
      spec.paragraphs.map(paraXml).join('') +
      `<w:sectPr>${spec.sectPr ?? LETTER_SECTPR}</w:sectPr>` +
      '</w:body></w:document>',
  );

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function writeDocx(path: string, spec: DocSpec): Promise<string> {
  await writeFile(path, await buildDocx(spec));
  return path;
}

/** Read a .docx from disk into the engine's input shape. */
export async function openDocx(path: string): Promise<DocxInput> {
  return { data: await readFile(path), name: basename(path) };
}

/** Load a written .docx back as a package, for asserting on the output. */
export async function loadDocx(path: string): Promise<DocxPackage> {
  return DocxPackage.fromBuffer(await readFile(path), basename(path));
}

/** A realistic book template: styles, trim size, headers-free, indented body. */
export const BOOK_TEMPLATE: DocSpec = {
  sectPr: DIGEST_SECTPR,
  styles: [
    { id: 'Normal', name: 'Normal', isDefault: true, fontName: 'Garamond', fontSizePt: 11 },
    {
      id: 'BodyText',
      name: 'Body Text',
      basedOn: 'Normal',
      firstLineIndent: 288,
      alignment: 'both',
      fontSizePt: 11,
    },
    {
      id: 'FirstParagraph',
      name: 'First Paragraph',
      basedOn: 'BodyText',
      next: 'BodyText',
      firstLineIndent: 0,
    },
    {
      id: 'ChapterTitle',
      name: 'Chapter Title',
      basedOn: 'Normal',
      next: 'FirstParagraph',
      outlineLevel: 0,
      alignment: 'center',
      pageBreakBefore: true,
      keepNext: true,
      spaceBefore: 1440,
      spaceAfter: 720,
      fontSizePt: 18,
      bold: true,
    },
    {
      id: 'SceneBreak',
      name: 'Scene Break',
      basedOn: 'Normal',
      alignment: 'center',
      spaceBefore: 240,
      spaceAfter: 240,
      firstLineIndent: 0,
    },
    { id: 'BlockQuote', name: 'Block Quote', basedOn: 'Normal', leftIndent: 720, firstLineIndent: 0 },
    { id: 'BookTitle', name: 'Title', basedOn: 'Normal', alignment: 'center', fontSizePt: 28 },
    { id: 'Subhead', name: 'Heading 2', basedOn: 'Normal', outlineLevel: 1, bold: true },
  ],
  paragraphs: [
    { text: 'The Sample Book', style: 'BookTitle' },
    { text: 'Chapter One', style: 'ChapterTitle' },
    {
      text: 'This opening paragraph of the template shows how the first paragraph of a chapter is set flush left with no indent, which is the usual convention in printed books everywhere.',
      style: 'FirstParagraph',
    },
    {
      text: 'Every following paragraph is indented by a quarter inch, justified, and set in the body face chosen for this book so the page has an even colour throughout the whole text block.',
      style: 'BodyText',
    },
    { text: '* * *', style: 'SceneBreak' },
    {
      text: 'After a scene break the text resumes, again without an indent on the first line, and continues in the same measured way until the chapter reaches its end.',
      style: 'FirstParagraph',
    },
    {
      text: 'A further paragraph of body text keeps the sample honest and gives the analyzer enough prose to work out which style carries the bulk of the book.',
      style: 'BodyText',
    },
  ],
};
