import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { analyzeReference } from '../src/core/analyze/reference.js';
import { analyzeManuscript } from '../src/core/analyze/manuscript.js';
import { formatManuscript } from '../src/core/platform/node.js';
import { attr, child, children, textOf } from '../src/core/ooxml/xml.js';
import {
  DIGEST_SECTPR,
  loadDocx,
  openDocx,
  type DocSpec,
  type ParaSpec,
  writeDocx,
} from './helpers/makeDocx.js';

/**
 * Regressions found against Amazon's real KDP interior template, whose style
 * names ("Endure - Chapter Title") and chapter sink defeated earlier detection.
 */

let dir = '';
let counter = 0;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'mf-kdp-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const PROSE = 'Insert chapter one text here. '.repeat(8).trim();

/** A template shaped like Amazon's: prefixed names plus unused built-ins. */
const KDP_LIKE: DocSpec = {
  sectPr: DIGEST_SECTPR,
  styles: [
    { id: 'Normal', name: 'Normal', isDefault: true, fontSizePt: 9 },
    { id: 'ChapBody', name: 'Endure - Chapter Body Text', firstLineIndent: 180, fontSizePt: 9 },
    { id: 'FirstPara', name: 'Endure - First Paragraph Body Text', next: 'ChapBody', firstLineIndent: 0 },
    { id: 'ChapTitle', name: 'Endure - Chapter Title', alignment: 'center', fontSizePt: 13 },
    { id: 'BookTitle', name: 'Endure - Book Title', alignment: 'center', fontSizePt: 36 },
    { id: 'CopyPage', name: 'Endure - Copyright Page', fontSizePt: 8.5 },
    { id: 'Subhead', name: 'Endure - Subhead', fontSizePt: 9.5 },
    // Unused Word built-ins that must not win over the template's own styles.
    { id: 'Heading2', name: 'heading 2', outlineLevel: 1, fontSizePt: 13 },
    { id: 'Subtitle', name: 'Subtitle', alignment: 'center', fontSizePt: 18 },
    { id: 'Title', name: 'Title', alignment: 'center', fontSizePt: 36 },
  ],
  paragraphs: [
    { text: 'Book Title', style: 'BookTitle' },
    { text: 'Copyright © 2025 Author Name', style: 'CopyPage' },
    { text: 'All rights reserved.', style: 'CopyPage' },
    { text: 'ISBN: xxx-xxx-xxxx-xx-x', style: 'CopyPage' },
    // Like Amazon's, this template starts each division with a section break
    // rather than a page-break-before on the style.
    { text: '', style: 'CopyPage', sectPr: `<w:type w:val="nextPage"/>${DIGEST_SECTPR}` },
    // A front-matter division set in the chapter style, with its own sink.
    { text: 'DEDICATION', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: 'Insert dedication text here.', style: 'FirstPara' },
    { text: '', sectPr: `<w:type w:val="nextPage"/>${DIGEST_SECTPR}` },
    // A chapter opening: five blank title paragraphs sink it down the page.
    { text: '', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: '1 CHAPTER NAME', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: '', style: 'ChapTitle' },
    { text: PROSE, style: 'FirstPara' },
    { text: PROSE, style: 'ChapBody' },
    { text: 'Subhead', style: 'Subhead' },
    { text: PROSE, style: 'ChapBody' },
    { text: PROSE, style: 'ChapBody' },
  ],
};

/** A title page, a copyright page, then two chapters with a subtitle. */
const MANUSCRIPT: ParaSpec[] = [
  {
    runs: [{ text: 'The Cartographer of Small Hours', bold: true, fontSizePt: 22 }],
    alignment: 'center',
  },
  { text: '', leadingPageBreak: true },
  { text: 'Copyright © 2026 by A. N. Author', alignment: 'center' },
  { text: 'This book is a work of fiction. Any resemblance is coincidental.', alignment: 'center' },
  { text: 'First edition.', alignment: 'center' },
  { text: 'All rights reserved. No part of this book may be reproduced.', alignment: 'center' },
  { text: '', leadingPageBreak: true },
  { runs: [{ text: 'Chapter One', bold: true }], alignment: 'center' },
  { runs: [{ text: 'THE SHOP ON MARKET STREET', bold: true }], alignment: 'center' },
  { text: '\tThe morning came in grey and unhurried, the way mornings did in that town.' },
  { text: '\tHollis pushed the shop door open with his shoulder and waited for the bell.' },
  { text: '', leadingPageBreak: true },
  { runs: [{ text: 'Chapter Two', bold: true }], alignment: 'center' },
  { text: '\tThe corridor smelled of wet wool, and the notice board had not been changed.' },
];

describe('a real KDP interior template', () => {
  it('maps roles onto the template styles, not onto unused built-ins', async () => {
    const path = await writeDocx(join(dir, `t${counter++}.docx`), KDP_LIKE);
    const { profile } = await analyzeReference(await openDocx(path));

    // "Book Title" starts with "Book", which once read as a part heading and
    // let a style used once beat the chapter style used fourteen times.
    expect(profile.roleStyles.chapterTitle).toBe('ChapTitle');
    expect(profile.roleStyles.frontMatterTitle).toBe('BookTitle');
    expect(profile.roleStyles.copyright).toBe('CopyPage');
    expect(profile.roleStyles.body).toBe('ChapBody');
    expect(profile.roleStyles.bodyFirst).toBe('FirstPara');
    // The template's own subhead beats the unused built-in "heading 2".
    expect(profile.roleStyles.subheading).toBe('Subhead');
    // Unused "Subtitle" is a title-page style at 18pt; better to have none.
    expect(profile.roleStyles.chapterSubtitle).toBeNull();
  });

  it('never gives one style to two different roles', async () => {
    const path = await writeDocx(join(dir, `t${counter++}.docx`), KDP_LIKE);
    const { profile } = await analyzeReference(await openDocx(path));

    const assigned = Object.values(profile.roleStyles).filter(
      (id): id is string => id !== null,
    );
    // `frontMatter` and `copyright` may legitimately share a fallback, so
    // compare only the roles that must stay distinct.
    const distinct = [
      profile.roleStyles.chapterTitle,
      profile.roleStyles.frontMatterTitle,
      profile.roleStyles.copyright,
      profile.roleStyles.body,
      profile.roleStyles.bodyFirst,
    ];
    expect(new Set(distinct).size).toBe(distinct.length);
    expect(assigned.length).toBeGreaterThan(4);
  });

  it('measures how far the template sinks a chapter opening', async () => {
    const path = await writeDocx(join(dir, `t${counter++}.docx`), KDP_LIKE);
    const { profile } = await analyzeReference(await openDocx(path));

    expect(profile.chapterTitleBlanksBefore).toBe(5);
    expect(profile.chapterTitleBlanksAfter).toBe(2);
  });

  it('reads a copyright notice as page content, not as a chapter heading', async () => {
    const path = await writeDocx(join(dir, `m${counter++}.docx`), { paragraphs: MANUSCRIPT });
    const { analysis } = await analyzeManuscript(await openDocx(path));
    const byText = (needle: string) =>
      analysis.blocks.find((b) => b.preview.includes(needle));

    expect(byText('Copyright ©')?.role).toBe('copyright');
    expect(byText('work of fiction')?.role).toBe('copyright');
    expect(byText('First edition')?.role).toBe('copyright');
    expect(byText('All rights reserved')?.role).toBe('copyright');
    // The real chapters are still found, and nothing on the copyright page
    // was mistaken for one.
    expect(analysis.chapterCount).toBe(2);
    expect(byText('Chapter One')?.role).toBe('chapterTitle');
    expect(byText('Chapter Two')?.role).toBe('chapterTitle');
  });

  it('lets the reviewer override how far a chapter title is pushed down', async () => {
    const id = counter++;
    const reference = await writeDocx(join(dir, `r${id}.docx`), KDP_LIKE);
    const manuscript = await writeDocx(join(dir, `m${id}.docx`), { paragraphs: MANUSCRIPT });
    const output = join(dir, `o${id}.docx`);
    await formatManuscript({
      referencePath: reference,
      manuscriptPath: manuscript,
      outputPath: output,
      options: { chapterSpaceBefore: 2, chapterSpaceAfter: 0 },
    });

    const pkg = await loadDocx(output);
    const doc = await pkg.readXml(pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const rows = children(body, 'p').map((p) => ({
      style: attr(child(child(p, 'pPr'), 'pStyle'), 'val'),
      text: textOf(p).trim(),
    }));

    const twoAt = rows.findIndex((r) => r.text === 'Chapter Two');
    expect(rows.slice(twoAt - 2, twoAt).every((r) => r.style === 'ChapTitle' && r.text === '')).toBe(
      true,
    );
    // Three blanks would mean the override was ignored.
    expect(rows[twoAt - 3].text).not.toBe('');
    // Nothing blank below the title now that the space after is set to none.
    expect(rows[twoAt + 1].text).not.toBe('');
  });

  it('reproduces the chapter sink and the copyright page in the output', async () => {
    const id = counter++;
    const reference = await writeDocx(join(dir, `r${id}.docx`), KDP_LIKE);
    const manuscript = await writeDocx(join(dir, `m${id}.docx`), { paragraphs: MANUSCRIPT });
    const output = join(dir, `o${id}.docx`);
    await formatManuscript({ referencePath: reference, manuscriptPath: manuscript, outputPath: output });

    const pkg = await loadDocx(output);
    const doc = await pkg.readXml(pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const rows = children(body, 'p').map((p) => ({
      style: attr(child(child(p, 'pPr'), 'pStyle'), 'val'),
      text: textOf(p).trim(),
      breaks: child(child(p, 'pPr'), 'pageBreakBefore') !== null,
    }));

    const titleAt = rows.findIndex((r) => r.text === 'Chapter One');
    expect(titleAt).toBeGreaterThan(0);
    expect(rows[titleAt].style).toBe('ChapTitle');

    // Five blank title paragraphs sink the opening down the page.
    const sink = rows.slice(titleAt - 5, titleAt);
    expect(sink.every((r) => r.style === 'ChapTitle' && r.text === '')).toBe(true);
    expect(rows[titleAt].breaks).toBe(false);

    // Chapter two has no section break ahead of it, so exactly one page break
    // sits on the first blank of its sink — never on the title itself, which
    // would strand the blanks at the foot of the previous page.
    const twoAt = rows.findIndex((r) => r.text === 'Chapter Two');
    const sinkTwo = rows.slice(twoAt - 5, twoAt);
    expect(sinkTwo.every((r) => r.style === 'ChapTitle' && r.text === '')).toBe(true);
    expect(sinkTwo.filter((r) => r.breaks)).toHaveLength(1);
    expect(sinkTwo[0].breaks).toBe(true);
    expect(rows[twoAt].breaks).toBe(false);

    // Two blank paragraphs follow, then the subtitle joins the chapter head.
    expect(rows.slice(titleAt + 1, titleAt + 3).every((r) => r.style === 'ChapTitle' && r.text === '')).toBe(true);
    expect(rows[titleAt + 3]).toMatchObject({
      style: 'ChapTitle',
      text: 'THE SHOP ON MARKET STREET',
    });

    // The whole copyright page uses the template's copyright style.
    const copyright = rows.filter((r) => r.style === 'CopyPage');
    expect(copyright).toHaveLength(4);
    expect(copyright.map((r) => r.text).join(' ')).toContain('All rights reserved');

    // Body text lands on the template's body styles.
    expect(rows[titleAt + 4].style).toBe('FirstPara');
    expect(rows[titleAt + 5].style).toBe('ChapBody');
  });
});
