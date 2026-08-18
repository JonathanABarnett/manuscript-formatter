import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
  checkChapterNumbers,
  formatChapterTitle,
  numberWords,
  parseChapterTitle,
  romanNumeral,
  wordsValue,
} from '../src/core/analyze/chapterNumbers.js';
import { chapterTitleTexts } from '../src/core/roles.js';
import { child, children, descendants, textOf } from '../src/core/ooxml/xml.js';
import { analyzeManuscript } from '../src/core/analyze/manuscript.js';
import { preflight } from '../src/core/preflight.js';
import { analyzeReference } from '../src/core/analyze/reference.js';
import { formatManuscript } from '../src/core/platform/node.js';
import { DEFAULT_FORMAT_OPTIONS, NO_EXTRA_SECTIONS } from '../src/core/types.js';
import { BOOK_TEMPLATE, loadDocx, openDocx, writeDocx, type ParaSpec } from './helpers/makeDocx.js';

describe('reading chapter numbers', () => {
  it('reads labels, digits, numerals and words', () => {
    expect(parseChapterTitle('Chapter 7')).toMatchObject({ label: 'Chapter', number: 7, form: 'digits', rest: '' });
    expect(parseChapterTitle('CHAPTER SEVEN')).toMatchObject({ number: 7, form: 'words', caps: true });
    expect(parseChapterTitle('Chapter Twenty-One')).toMatchObject({ number: 21, form: 'words', caps: false });
    expect(parseChapterTitle('Chapter twenty one')).toMatchObject({ number: 21 });
    expect(parseChapterTitle('Chapter XII')).toMatchObject({ number: 12, form: 'roman' });
    expect(parseChapterTitle('XII')).toMatchObject({ label: null, number: 12, form: 'roman', caps: false });
    expect(parseChapterTitle('12')).toMatchObject({ label: null, number: 12, form: 'digits' });
    expect(parseChapterTitle('Twelve')).toMatchObject({ label: null, number: 12, form: 'words' });
    expect(parseChapterTitle('Capítulo 3')).toMatchObject({ label: 'Capítulo', number: 3 });
  });

  it('keeps a title that follows the number', () => {
    expect(parseChapterTitle('Chapter 7: The Meeting')).toMatchObject({
      number: 7,
      separator: ': ',
      rest: 'The Meeting',
    });
    expect(parseChapterTitle('7. The Meeting')).toMatchObject({ number: 7, separator: '. ', rest: 'The Meeting' });
    expect(parseChapterTitle('Seven — The Meeting')).toMatchObject({ number: 7, rest: 'The Meeting' });
    expect(parseChapterTitle('Chapter 7 The Meeting')).toMatchObject({ number: 7, rest: 'The Meeting' });
  });

  it('does not mistake titles for numbers', () => {
    expect(parseChapterTitle('Prologue')).toBeNull();
    expect(parseChapterTitle('One Day More')).toBeNull();
    expect(parseChapterTitle('I Am Legend')).toBeNull();
    expect(parseChapterTitle('Nineteen Eighty-Four')).toBeNull();
    expect(parseChapterTitle('Chapters of Sand')).toBeNull();
    expect(parseChapterTitle('3rd Time Lucky')).toBeNull();
    expect(parseChapterTitle('MIX')).toBeNull();
    expect(parseChapterTitle('1984')).toBeNull();
    expect(parseChapterTitle('The Meeting')).toBeNull();
  });

  it('spells numbers out and back', () => {
    expect(numberWords(1)).toBe('One');
    expect(numberWords(21)).toBe('Twenty-One');
    expect(numberWords(40)).toBe('Forty');
    expect(numberWords(101)).toBe('One Hundred and One');
    expect(wordsValue('One Hundred and One')).toBe(101);
    expect(wordsValue('one twenty')).toBeNull();
    expect(romanNumeral(4)).toBe('IV');
    expect(romanNumeral(49)).toBe('XLIX');
  });

  it('writes a title in each style', () => {
    const parsed = parseChapterTitle('CHAPTER 7: The Meeting')!;
    expect(formatChapterTitle(parsed, 'keep', 7)).toBe('CHAPTER 7: The Meeting');
    expect(formatChapterTitle(parsed, 'keep', 8)).toBe('CHAPTER 8: The Meeting');
    expect(formatChapterTitle(parsed, 'chapterWords', 7)).toBe('CHAPTER SEVEN: The Meeting');
    expect(formatChapterTitle(parsed, 'chapterDigits', 7)).toBe('CHAPTER 7: The Meeting');
    expect(formatChapterTitle(parsed, 'words', 7)).toBe('SEVEN: The Meeting');
    expect(formatChapterTitle(parsed, 'digits', 7)).toBe('7: The Meeting');

    const bare = parseChapterTitle('12')!;
    expect(formatChapterTitle(bare, 'chapterWords', 12)).toBe('Chapter Twelve');
    expect(formatChapterTitle(bare, 'keep', 13)).toBe('13');
    const roman = parseChapterTitle('Chapter XII')!;
    expect(formatChapterTitle(roman, 'keep', 14)).toBe('Chapter XIV');
    // Spelled-out numbers are English, so a foreign label keeps its digits.
    const spanish = parseChapterTitle('Capítulo 3')!;
    expect(formatChapterTitle(spanish, 'chapterWords', 3)).toBe('Capítulo 3');
  });

  it('checks the sequence', () => {
    const titles = ['Prologue', 'Chapter 1', 'Chapter 2', 'Chapter 2', 'Chapter 5', 'Chapter 4'].map(
      (text, index) => ({ index, text }),
    );
    const report = checkChapterNumbers(titles);
    expect(report.numbered.map((n) => n.number)).toEqual([1, 2, 2, 5, 4]);
    expect(report.duplicates).toEqual([2]);
    expect(report.gaps).toEqual([3, 4]);
    expect(report.outOfOrder).toEqual([4]);
    expect(report.mixed).toBe(false);
    expect(checkChapterNumbers([{ index: 0, text: 'Chapter 1' }, { index: 1, text: 'Two' }]).mixed).toBe(true);
  });

  it('renumbers only the chapters that carry a number', () => {
    const blocks = ['Prologue', 'Chapter 3', 'Interlude', 'Chapter 9', 'Chapter Ten'].map((text, index) => ({
      index,
      text,
    }));
    const map = chapterTitleTexts(blocks, () => 'chapterTitle', {
      chapterNumberStyle: 'keep',
      renumberChapters: true,
    });
    expect([...map.entries()]).toEqual([
      [1, 'Chapter 1'],
      [3, 'Chapter 2'],
      [4, 'Chapter Three'],
    ]);
    const uniform = chapterTitleTexts(blocks, () => 'chapterTitle', {
      chapterNumberStyle: 'chapterDigits',
      renumberChapters: false,
    });
    expect([...uniform.entries()]).toEqual([[4, 'Chapter 10']]);
  });
});

describe('chapter numbers in a formatted book', () => {
  let dir = '';
  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'mf-chnum-'));
  });
  afterAll(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const PROSE =
    'The morning came in grey and unhurried, the way mornings did there. Hollis pushed the ' +
    'shop door open with his shoulder and waited for the bell.';
  const MANUSCRIPT: ParaSpec[] = [
    { runs: [{ text: 'Chapter 1: ' }, { text: 'Middlemarch', italic: true }, { text: ' Revisited' }], alignment: 'center' },
    { text: PROSE },
    { runs: [{ text: 'Chapter 3' }], alignment: 'center' },
    { text: PROSE },
    { runs: [{ text: 'Chapter 3' }], alignment: 'center' },
    { text: PROSE },
  ];

  it('warns about the sequence, then renumbers on request keeping the title emphasis', async () => {
    const referencePath = await writeDocx(join(dir, 'ref.docx'), BOOK_TEMPLATE);
    const manuscriptPath = await writeDocx(join(dir, 'ms.docx'), { paragraphs: MANUSCRIPT });
    const outputPath = join(dir, 'out.docx');

    const { profile } = await analyzeReference(await openDocx(referencePath));
    const { analysis } = await analyzeManuscript(await openDocx(manuscriptPath));
    const before = preflight({ profile, analysis, options: DEFAULT_FORMAT_OPTIONS });
    const sequence = before.checks.find((c) => c.id === 'chapter-numbers');
    expect(sequence?.level).toBe('check');
    expect(sequence?.detail).toContain('number 3 is used twice');
    expect(sequence?.detail).toContain('number 2 is missing');

    const options = { ...DEFAULT_FORMAT_OPTIONS, renumberChapters: true, chapterNumberStyle: 'chapterWords' as const };
    const after = preflight({ profile, analysis, options });
    expect(after.checks.find((c) => c.id === 'chapter-numbers')).toBeUndefined();
    expect(after.checks.find((c) => c.id === 'chapter-numbers-renumbered')?.level).toBe('ready');

    await formatManuscript({
      referencePath,
      manuscriptPath,
      outputPath,
      options: { ...options, extraSections: NO_EXTRA_SECTIONS },
    });
    const pkg = await loadDocx(outputPath);
    const doc = await pkg.readXml(pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const titles = children(body, 'p').filter((p) => /Chapter/.test(textOf(p)));
    expect(titles.map((p) => textOf(p))).toEqual([
      'Chapter One: Middlemarch Revisited',
      'Chapter Two',
      'Chapter Three',
    ]);
    // The italic run inside the first title survived the rewrite.
    const italic = descendants(titles[0], 'i');
    expect(italic).toHaveLength(1);
    expect(textOf(italic[0].parentNode?.parentNode ?? null)).toBe('Middlemarch');
  });
});
