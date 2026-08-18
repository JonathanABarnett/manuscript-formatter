import { describe, expect, it } from 'vitest';
import { formatToBuffer } from '../src/core/format.js';
import { DocxPackage } from '../src/core/ooxml/package.js';
import { child, children, textOf } from '../src/core/ooxml/xml.js';
import { RELTYPE } from '../src/core/ooxml/ns.js';
import { willReplaceRunningHead } from '../src/core/build/headers.js';
import { buildSampleManuscript, buildTemplate } from '../src/core/templates/generate.js';
import { EMPTY_BOOK_DETAILS, NO_EXTRA_SECTIONS, type BookDetails, type DocxInput } from '../src/core/types.js';

/**
 * A template ships its running heads filled with wording like "BOOK TITLE".
 * Copying it across verbatim leaves the author's book headed by someone
 * else's placeholder on every page.
 */

const asInput = (data: Uint8Array, name: string): DocxInput => ({ data, name });

const DETAILS: BookDetails = {
  ...EMPTY_BOOK_DETAILS,
  title: 'The Cartographer of Small Hours',
  author: 'A. N. Author',
};

describe('recognising placeholder wording in running heads', () => {
  it('spots the wording a template ships with', () => {
    for (const text of ['BOOK TITLE', 'Book Title', 'Title', '[Your Title]', 'TITLE GOES HERE']) {
      expect(willReplaceRunningHead(text, DETAILS)).toBe(true);
    }
    for (const text of ['AUTHOR NAME', 'Author Name', 'Author', '[Your Name]']) {
      expect(willReplaceRunningHead(text, DETAILS)).toBe(true);
    }
  });

  it('leaves wording an author put there themselves alone', () => {
    for (const text of ['A Cottage in Yorkshire', 'Part Two', 'Chapter 4', 'Hollis & Sons']) {
      expect(willReplaceRunningHead(text, DETAILS)).toBe(false);
    }
  });

  it('does nothing when the author has typed no details', () => {
    expect(willReplaceRunningHead('BOOK TITLE', EMPTY_BOOK_DETAILS)).toBe(false);
  });
});

/** A design whose headers carry the usual placeholder wording. */
async function templateWithHeads(): Promise<Uint8Array> {
  const base = await buildTemplate('6x9', 'classic');
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(base);
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const head = (text: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}">` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;

  zip.file('word/header1.xml', head('BOOK TITLE'));
  zip.file('word/header2.xml', head('AUTHOR NAME'));
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdH1" Type="${RELTYPE.header}" Target="header1.xml"/>` +
        `<Relationship Id="rIdH2" Type="${RELTYPE.header}" Target="header2.xml"/></Relationships>`,
    ),
  );
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  );
  const doc = await zip.file('word/document.xml')!.async('string');
  zip.file(
    'word/document.xml',
    doc.replace(
      '<w:sectPr>',
      '<w:sectPr><w:headerReference w:type="even" r:id="rIdH1"/><w:headerReference w:type="default" r:id="rIdH2"/>',
    ),
  );
  return zip.generateAsync({ type: 'uint8array' });
}

async function headTexts(data: Uint8Array): Promise<string[]> {
  const pkg = await DocxPackage.fromBuffer(Buffer.from(data), 'out.docx');
  const rels = await pkg.relsFor(pkg.documentPath);
  const out: string[] = [];
  for (const rel of rels.all().filter((r) => r.type === RELTYPE.header)) {
    const doc = await pkg.readXml(pkg.resolveTarget(pkg.documentPath, rel.target));
    const text = textOf(doc?.documentElement ?? null).replace(/\s+/g, ' ').trim();
    if (text) out.push(text);
  }
  return out.sort();
}

describe('putting the author into the running heads', () => {
  it('replaces the placeholders, keeping the design’s capitals', async () => {
    const out = await formatToBuffer({
      reference: asInput(await templateWithHeads(), 'design.docx'),
      manuscript: asInput(await buildSampleManuscript(), 'sample.docx'),
      options: { bookDetails: DETAILS, extraSections: NO_EXTRA_SECTIONS },
    });

    // The template shouted; so should the replacement, or the page changes look.
    expect(await headTexts(out.data)).toEqual([
      'A. N. AUTHOR',
      'THE CARTOGRAPHER OF SMALL HOURS',
    ]);
    expect(out.stats.runningHeadsUpdated).toBe(2);
  });

  it('leaves the headers untouched when no details were typed', async () => {
    const out = await formatToBuffer({
      reference: asInput(await templateWithHeads(), 'design.docx'),
      manuscript: asInput(await buildSampleManuscript(), 'sample.docx'),
      options: { bookDetails: EMPTY_BOOK_DETAILS, extraSections: NO_EXTRA_SECTIONS },
    });

    expect(await headTexts(out.data)).toEqual(['AUTHOR NAME', 'BOOK TITLE']);
    expect(out.stats.runningHeadsUpdated).toBe(0);
  });
});

describe('where the contents page goes', () => {
  it('sits after the book’s own opening pages, not before them', async () => {
    const out = await formatToBuffer({
      reference: asInput(await buildTemplate('6x9', 'classic'), 'design.docx'),
      manuscript: asInput(await buildSampleManuscript(), 'sample.docx'),
      options: {
        bookDetails: EMPTY_BOOK_DETAILS,
        extraSections: { ...NO_EXTRA_SECTIONS, contents: true },
      },
    });

    const pkg = await DocxPackage.fromBuffer(Buffer.from(out.data), 'out.docx');
    const doc = await pkg.readXml(pkg.documentPath);
    const rows = children(child(doc?.documentElement ?? null, 'body'), 'p')
      .map((p) => textOf(p).replace(/\s+/g, ' ').trim())
      .filter(Boolean);

    const contents = rows.findIndex((t) => t === 'Contents');
    const copyright = rows.findIndex((t) => t.startsWith('Copyright ©'));
    const firstChapter = rows.findIndex((t) => t === 'CHAPTER ONE');

    expect(contents).toBeGreaterThan(-1);
    // After the manuscript's own copyright page, and before the book begins.
    expect(contents).toBeGreaterThan(copyright);
    expect(contents).toBeLessThan(firstChapter);
  });
});

describe('setting the headers by hand', () => {
  it('puts the author’s own wording on each side of the book', async () => {
    const out = await formatToBuffer({
      reference: asInput(await templateWithHeads(), 'design.docx'),
      manuscript: asInput(await buildSampleManuscript(), 'sample.docx'),
      options: {
        bookDetails: DETAILS,
        extraSections: NO_EXTRA_SECTIONS,
        runningHeads: { mode: 'custom', verso: 'Hollis & Sons', recto: 'Part One' },
      },
    });

    // header1 is the even (left) page, header2 the default (right) one.
    expect(await headTexts(out.data)).toEqual(['Hollis & Sons', 'Part One']);
  });

  it('leaves the design alone when asked to', async () => {
    const out = await formatToBuffer({
      reference: asInput(await templateWithHeads(), 'design.docx'),
      manuscript: asInput(await buildSampleManuscript(), 'sample.docx'),
      options: {
        bookDetails: DETAILS,
        extraSections: NO_EXTRA_SECTIONS,
        runningHeads: { mode: 'leave', verso: '', recto: '' },
      },
    });

    expect(await headTexts(out.data)).toEqual(['AUTHOR NAME', 'BOOK TITLE']);
    expect(out.stats.runningHeadsUpdated).toBe(0);
  });
});

describe('reading the details out of the manuscript', () => {
  it('finds the author and year an author already typed', async () => {
    const { analyzeManuscript } = await import('../src/core/format.js');
    const { analysis } = await analyzeManuscript(
      asInput(await buildSampleManuscript(), 'sample.docx'),
    );

    expect(analysis.detectedDetails.author).toBe('A. N. Author');
    expect(analysis.detectedDetails.copyrightYear).toBe('2026');
    expect(analysis.detectedDetails.title).toBe('The Cartographer of Small Hours');
  });
});
