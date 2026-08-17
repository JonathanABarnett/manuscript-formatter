import { describe, expect, it } from 'vitest';
import { formatToBuffer } from '../src/core/format.js';
import { DocxPackage } from '../src/core/ooxml/package.js';
import { attr, child, children, descendants, textOf } from '../src/core/ooxml/xml.js';
import { buildSampleManuscript, buildTemplate } from '../src/core/templates/generate.js';
import { EMPTY_BOOK_DETAILS, NO_EXTRA_SECTIONS, type DocxInput, type FormatOptions } from '../src/core/types.js';

/**
 * The generated title, copyright and contents pages. The rule these enforce is
 * that a section only appears when the author supplied the words for it — a
 * heading over an empty page is worse than no page at all.
 */

const asInput = (data: Uint8Array, name: string): DocxInput => ({ data, name });

const DETAILS = {
  ...EMPTY_BOOK_DETAILS,
  title: 'The Cartographer of Small Hours',
  subtitle: 'A Novel',
  author: 'A. N. Author',
  copyrightYear: '2026',
  publisher: 'Small Hours Press',
  isbn: '978-1-234567-89-0',
  dedication: 'For everyone who waited.',
  acknowledgments: 'Thanks to the usual suspects.\n\nAnd to the unusual ones.',
  aboutTheAuthor: 'A. N. Author lives by the sea.',
};

async function build(options: Partial<FormatOptions>) {
  const [template, manuscript] = await Promise.all([
    buildTemplate('6x9', 'classic'),
    buildSampleManuscript(),
  ]);
  const out = await formatToBuffer({
    reference: asInput(template, 'design.docx'),
    manuscript: asInput(manuscript, 'sample.docx'),
    options,
  });
  const pkg = await DocxPackage.fromBuffer(Buffer.from(out.data), 'out.docx');
  const doc = await pkg.readXml(pkg.documentPath);
  const body = child(doc?.documentElement ?? null, 'body')!;
  return {
    pkg,
    body,
    rows: children(body, 'p').map((p) => ({
      style: attr(child(child(p, 'pPr'), 'pStyle'), 'val'),
      text: textOf(p).trim(),
      breaks: child(child(p, 'pPr'), 'pageBreakBefore') !== null,
    })),
    text: textOf(body),
  };
}

describe('the generated opening pages', () => {
  it('adds nothing at all when no sections are switched on', async () => {
    const { text } = await build({ bookDetails: DETAILS, extraSections: NO_EXTRA_SECTIONS });
    // Assert on wording only the generator produces: the sample manuscript
    // carries a copyright page of its own, which must survive untouched.
    expect(text).not.toContain('Small Hours Press');
    expect(text).not.toContain('ISBN:');
    expect(text).not.toContain('Contents');
    expect(text).toContain('All rights reserved');
  });

  it('builds a title page from the details typed in', async () => {
    const { rows } = await build({
      bookDetails: DETAILS,
      extraSections: { ...NO_EXTRA_SECTIONS, titlePage: true },
      replaceFrontMatter: true,
    });

    expect(rows[0]).toMatchObject({ style: 'BookTitle', text: 'The Cartographer of Small Hours' });
    expect(rows[1].text).toBe('A Novel');
    expect(rows[2].text).toBe('A. N. Author');
    // Replacing the front matter must not leave the manuscript's own title too.
    expect(rows.filter((r) => r.text === 'The Cartographer of Small Hours')).toHaveLength(1);
  });

  it("builds a copyright page using the design's own copyright style", async () => {
    const { rows } = await build({
      bookDetails: DETAILS,
      extraSections: { ...NO_EXTRA_SECTIONS, copyrightPage: true },
      replaceFrontMatter: true,
    });

    const copyright = rows.filter((r) => r.style === 'CopyrightPage');
    expect(copyright.length).toBeGreaterThanOrEqual(5);
    expect(copyright[0].text).toBe('Copyright © 2026 by A. N. Author');
    const joined = copyright.map((r) => r.text).join(' | ');
    expect(joined).toContain('All rights reserved.');
    expect(joined).toContain('Small Hours Press');
    expect(joined).toContain('ISBN: 978-1-234567-89-0');
  });

  it('skips a section the author left blank rather than heading an empty page', async () => {
    const { text } = await build({
      // Dedication and acknowledgments are on, but nothing was typed for them.
      bookDetails: { ...EMPTY_BOOK_DETAILS, title: 'A Book' },
      extraSections: {
        ...NO_EXTRA_SECTIONS,
        titlePage: true,
        dedication: true,
        acknowledgments: true,
        aboutTheAuthor: true,
      },
      replaceFrontMatter: true,
    });

    expect(text).toContain('A Book');
    expect(text).not.toContain('Acknowledgments');
    expect(text).not.toContain('About the Author');
  });

  it('adds a contents table Word can fill in, and asks Word to do it', async () => {
    const { pkg, body, rows } = await build({
      bookDetails: DETAILS,
      extraSections: { ...NO_EXTRA_SECTIONS, contents: true },
    });

    expect(rows.some((r) => r.text === 'Contents' && r.style === 'ChapterTitle')).toBe(true);
    const instructions = descendants(body, 'instrText').map((el) => el.textContent ?? '');
    expect(instructions.some((i) => i.includes('TOC'))).toBe(true);

    // Word only offers to build the table if the document asks it to.
    const rels = await pkg.relsFor(pkg.documentPath);
    const target = rels.all().find((r) => r.type.endsWith('/settings'))?.target;
    expect(target).toBeDefined();
    const settings = await pkg.readXml(pkg.resolveTarget(pkg.documentPath, target!));
    expect(child(settings?.documentElement ?? null, 'updateFields')).not.toBeNull();
  });

  it('puts back matter after the book, each starting a new page', async () => {
    const { rows } = await build({
      bookDetails: DETAILS,
      extraSections: { ...NO_EXTRA_SECTIONS, acknowledgments: true, aboutTheAuthor: true },
    });

    const ack = rows.findIndex((r) => r.text === 'Acknowledgments');
    const about = rows.findIndex((r) => r.text === 'About the Author');
    expect(ack).toBeGreaterThan(0);
    expect(about).toBeGreaterThan(ack);
    expect(rows[ack].breaks).toBe(true);
    expect(rows[about].breaks).toBe(true);

    // Both sit after the last of the manuscript's own chapters.
    const lastChapter = rows.map((r) => r.style).lastIndexOf('ChapterTitle');
    expect(lastChapter).toBeGreaterThanOrEqual(about);

    // A blank line between typed paragraphs becomes two paragraphs, not one.
    expect(rows[ack + 1].text).toBe('Thanks to the usual suspects.');
    expect(rows[ack + 2].text).toBe('And to the unusual ones.');
  });

  it("keeps the manuscript's own front matter when not replacing it", async () => {
    const { text } = await build({
      bookDetails: DETAILS,
      extraSections: { ...NO_EXTRA_SECTIONS, titlePage: true },
      replaceFrontMatter: false,
    });

    // Both the generated title and the sample's own copyright line survive.
    expect(text).toContain('The Cartographer of Small Hours');
    expect(text).toContain('All rights reserved');
  });
});
