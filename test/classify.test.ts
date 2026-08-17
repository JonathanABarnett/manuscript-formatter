import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { analyzeManuscript } from '../src/core/analyze/manuscript.js';
import type { BlockRole } from '../src/core/types.js';
import { type DocSpec, openDocx, type ParaSpec, writeDocx } from './helpers/makeDocx.js';

const loadManuscript = async (path: string) => analyzeManuscript(await openDocx(path));

let dir = '';
let counter = 0;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'mf-cls-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

async function roles(spec: DocSpec): Promise<BlockRole[]> {
  const path = await writeDocx(join(dir, `m${counter++}.docx`), spec);
  const { analysis } = await loadManuscript(path);
  return analysis.blocks.map((b) => b.role);
}

const PROSE_1 =
  'The morning came in grey and unhurried, the way mornings did in that part of ' +
  'the country. Hollis pushed the shop door open with his shoulder, and the ' +
  'small bell above it made its usual complaint about being disturbed so early.';
const PROSE_2 =
  'The corridor smelled of wet wool and yesterday rain. At the far end a notice ' +
  'board carried three curling announcements, none of them newer than the ' +
  'spring, and a fourth that somebody had torn neatly in half and left there.';
const PROSE_3 =
  'Outside, past the glass, the square was emptying. Wind moved scraps of paper ' +
  'in slow circles, and the light had gone thin in the way it does when an ' +
  'afternoon has quietly decided to end rather earlier than anyone expected.';

/** A manuscript as authors actually deliver them: no styles, tabs, all caps. */
const RAW_MANUSCRIPT: ParaSpec[] = [
  { runs: [{ text: 'The Sample Book', bold: true, fontSizePt: 20 }], alignment: 'center' },
  { text: 'by A. N. Author', alignment: 'center' },
  { text: '', leadingPageBreak: true },
  { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center' },
  { text: '' },
  { text: `\t${PROSE_1}` },
  { text: `\t${PROSE_2}` },
  { text: '#', alignment: 'center' },
  { text: `\t${PROSE_3}` },
  { text: '', leadingPageBreak: true },
  { runs: [{ text: 'CHAPTER TWO', bold: true }], alignment: 'center' },
  { text: '' },
  { text: `\t${PROSE_1}` },
];

describe('manuscript classification', () => {
  it('reads an unstyled manuscript the way a typesetter would', async () => {
    expect(await roles({ paragraphs: RAW_MANUSCRIPT })).toEqual([
      'frontMatterTitle',
      'frontMatter',
      'pageBreak',
      'chapterTitle',
      'empty',
      'bodyFirst',
      'body',
      'sceneBreak',
      'bodyFirst',
      'pageBreak',
      'chapterTitle',
      'empty',
      'bodyFirst',
    ]);
  });

  it('trusts explicit Word heading styles', async () => {
    expect(
      await roles({
        styles: [
          { id: 'Normal', name: 'Normal', isDefault: true },
          { id: 'Heading1', name: 'Heading 1', outlineLevel: 0 },
          { id: 'Heading2', name: 'Heading 2', outlineLevel: 1 },
        ],
        paragraphs: [
          { text: 'The Arrival', style: 'Heading1' },
          { text: PROSE_1 },
          { text: 'A Later Section', style: 'Heading2' },
          { text: PROSE_2 },
        ],
      }),
    ).toEqual(['chapterTitle', 'bodyFirst', 'subheading', 'bodyFirst']);
  });

  it('reads a bare number on its own line as a chapter opening', async () => {
    expect(
      await roles({
        paragraphs: [
          { text: '7', alignment: 'center' },
          { text: PROSE_1 },
          { text: 'VIII', alignment: 'center' },
          { text: PROSE_2 },
        ],
      }),
    ).toEqual(['chapterTitle', 'bodyFirst', 'chapterTitle', 'bodyFirst']);
  });

  it('does not read a numeric book title as a chapter number', async () => {
    expect(
      await roles({
        paragraphs: [
          { runs: [{ text: 'Nineteen Eighty-Four', bold: true }], alignment: 'center' },
          { text: 'by George Orwell', alignment: 'center' },
          { text: '', leadingPageBreak: true },
          { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center' },
          { text: PROSE_1 },
        ],
      }),
    ).toEqual(['frontMatterTitle', 'frontMatter', 'pageBreak', 'chapterTitle', 'bodyFirst']);
  });

  it('still reads a spelled-out chapter number as a heading', async () => {
    expect(
      await roles({
        paragraphs: [
          { text: 'Twenty-One', alignment: 'center' },
          { text: PROSE_1 },
          { text: 'One Hundred', alignment: 'center' },
          { text: PROSE_2 },
        ],
      }),
    ).toEqual(['chapterTitle', 'bodyFirst', 'chapterTitle', 'bodyFirst']);
  });

  it('folds a title line beneath a chapter number into a subtitle', async () => {
    expect(
      await roles({
        paragraphs: [
          { text: 'Chapter Three', alignment: 'center' },
          { text: 'The Long Road Home', alignment: 'center' },
          { text: PROSE_1 },
        ],
      }),
    ).toEqual(['chapterTitle', 'chapterSubtitle', 'bodyFirst']);
  });

  it('does not mistake dialogue or fragments for headings', async () => {
    expect(
      await roles({
        paragraphs: [
          { text: PROSE_1 },
          { text: '"Get out," she said.' },
          { text: 'He did not move.' },
          { text: 'Then, slowly:' },
          { text: PROSE_2 },
        ],
      }),
    ).toEqual(['bodyFirst', 'body', 'body', 'body', 'body']);
  });

  it('recognises the usual scene-break ornaments', async () => {
    const result = await roles({
      paragraphs: [
        { text: PROSE_1 },
        { text: '* * *', alignment: 'center' },
        { text: PROSE_2 },
        { text: '***', alignment: 'center' },
        { text: PROSE_3 },
        { text: '~', alignment: 'center' },
        { text: PROSE_1 },
        { text: '---', alignment: 'center' },
        { text: PROSE_2 },
      ],
    });
    expect(result.filter((r) => r === 'sceneBreak')).toHaveLength(4);
  });

  it('keeps indented quotations and numbered lists distinct', async () => {
    expect(
      await roles({
        numbering: true,
        paragraphs: [
          { text: PROSE_1 },
          { text: 'The report was unambiguous about the risk that remained.', leftIndent: 720 },
          { text: 'First item', numId: 1 },
          { text: 'Second item', numId: 1 },
          { text: PROSE_2 },
        ],
      }),
    ).toEqual(['bodyFirst', 'blockQuote', 'listItem', 'listItem', 'body']);
  });

  it('counts the structure it found', async () => {
    const path = await writeDocx(join(dir, 'counts.docx'), { paragraphs: RAW_MANUSCRIPT });
    const { analysis } = await loadManuscript(path);

    expect(analysis.chapterCount).toBe(2);
    expect(analysis.partCount).toBe(0);
    expect(analysis.sceneBreakCount).toBe(1);
    expect(analysis.paragraphCount).toBe(9);
    expect(analysis.bodyStartIndex).toBe(3);
    expect(analysis.wordCount).toBeGreaterThan(100);
  });

  it('counts parts separately from chapters', async () => {
    const path = await writeDocx(join(dir, 'parts.docx'), {
      paragraphs: [
        { runs: [{ text: 'PART ONE', bold: true }], alignment: 'center' },
        { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center', leadingPageBreak: true },
        { text: PROSE_1 },
        { runs: [{ text: 'CHAPTER TWO', bold: true }], alignment: 'center', leadingPageBreak: true },
        { text: PROSE_2 },
      ],
    });
    const { analysis } = await loadManuscript(path);

    expect(analysis.partCount).toBe(1);
    expect(analysis.chapterCount).toBe(2);
  });

  it('treats front matter as front matter, not as chapter one', async () => {
    const path = await writeDocx(join(dir, 'front.docx'), { paragraphs: RAW_MANUSCRIPT });
    const { analysis } = await loadManuscript(path);

    expect(analysis.blocks[0].preview).toBe('The Sample Book');
    expect(analysis.blocks[0].role).toBe('frontMatterTitle');
    expect(analysis.blocks[3].structuralMarker).toBe(true);
    expect(analysis.blocks[0].structuralMarker).toBe(false);
  });
});
