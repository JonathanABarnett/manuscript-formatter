import { describe, expect, it } from 'vitest';
import { analyzeDocuments, analyzeReference } from '../src/core/format.js';
import { TYPICAL_NOVEL_WORDS, lineSpacingMultiple } from '../src/core/pageEstimate.js';
import { BOOK_LOOKS, TRIM_SIZES, estimatePagesForDesign } from '../src/core/templates/design.js';
import { estimatePageCount, preflight, type PreflightReport } from '../src/core/preflight.js';
import { buildSampleManuscript, buildTemplate } from '../src/core/templates/generate.js';
import { DEFAULT_FORMAT_OPTIONS, EMPTY_BOOK_DETAILS, NO_EXTRA_SECTIONS } from '../src/core/types.js';
import type { DocxInput, FormatOptions, ManuscriptAnalysis, ReferenceProfile } from '../src/core/types.js';

/**
 * The pre-download report. Its job is to be honest: never claim a book is
 * KDP-approved, and always leave the last word to Amazon's own previewer.
 */

const asInput = (data: Uint8Array, name: string): DocxInput => ({ data, name });

async function analyzed(): Promise<{ profile: ReferenceProfile; analysis: ManuscriptAnalysis }> {
  const [template, manuscript] = await Promise.all([
    buildTemplate('6x9', 'classic'),
    buildSampleManuscript(),
  ]);
  const { profile, analysis } = await analyzeDocuments(
    asInput(template, 'design.docx'),
    asInput(manuscript, 'sample.docx'),
  );
  return { profile, analysis };
}

const withOptions = (over: Partial<FormatOptions> = {}): FormatOptions => ({
  ...DEFAULT_FORMAT_OPTIONS,
  bookDetails: { ...EMPTY_BOOK_DETAILS },
  extraSections: { ...NO_EXTRA_SECTIONS },
  ...over,
});

const find = (report: PreflightReport, id: string) => report.checks.find((c) => c.id === id);

describe('the preflight report', () => {
  it('passes a built-in design with the sample book', async () => {
    const { profile, analysis } = await analyzed();
    const report = preflight({ profile, analysis, options: withOptions() });

    expect(find(report, 'chapters')?.level).toBe('ready');
    expect(find(report, 'page-numbers-ok')?.level).toBe('ready');
    expect(find(report, 'no-chapters')).toBeUndefined();
    expect(find(report, 'placeholder-text')).toBeUndefined();
    // The design's page-number footer must not be read as header wording.
    expect(find(report, 'header-placeholder')).toBeUndefined();
  });

  it('leads with the most serious finding', async () => {
    const { profile, analysis } = await analyzed();
    const noNumbers = { ...profile, hasPageNumbers: false };
    const report = preflight({ profile: noNumbers, analysis, options: withOptions() });

    expect(report.level).toBe('attention');
    expect(report.checks[0].level).toBe('attention');
    expect(find(report, 'page-numbers')?.title).toMatch(/no page numbers/i);
  });

  it('spots filler text left over from a template', async () => {
    const { profile, analysis } = await analyzed();
    const withFiller: ManuscriptAnalysis = {
      ...analysis,
      blocks: [
        ...analysis.blocks,
        {
          ...analysis.blocks[0],
          index: 9999,
          text: 'Insert your chapter one text here.',
          preview: 'Insert your chapter one text here.',
          isEmpty: false,
        },
      ],
    };
    const report = preflight({ profile, analysis: withFiller, options: withOptions() });

    const check = find(report, 'placeholder-text');
    expect(check?.level).toBe('attention');
    // The offending line is offered as an example so the reviewer can be shown
    // it, rather than being told to go and find it.
    expect(check?.examples?.[0]).toMatchObject({
      index: 9999,
      preview: 'Insert your chapter one text here.',
    });
  });

  it('flags template wording still sitting in the page headers', async () => {
    const { profile, analysis } = await analyzed();
    const report = preflight({
      profile: { ...profile, headerFooterText: ['Book Title Goes Here'] },
      analysis,
      options: withOptions(),
    });

    expect(find(report, 'header-placeholder')?.level).toBe('attention');
    expect(find(report, 'header-text')).toBeUndefined();
  });

  it('asks the author to read real header wording without alarming them', async () => {
    const { profile, analysis } = await analyzed();
    const report = preflight({
      profile: { ...profile, headerFooterText: ['A. N. Author'] },
      analysis,
      options: withOptions(),
    });

    expect(find(report, 'header-text')?.level).toBe('check');
    expect(find(report, 'header-placeholder')).toBeUndefined();
  });

  it('warns when a picture is wider than the text block', async () => {
    const { profile, analysis } = await analyzed();
    const textWidth =
      profile.pageSetup.widthTwips -
      profile.pageSetup.margins.left -
      profile.pageSetup.margins.right -
      profile.pageSetup.margins.gutter;
    const wide: ManuscriptAnalysis = {
      ...analysis,
      blocks: [
        ...analysis.blocks,
        { ...analysis.blocks[0], index: 9998, imageWidthTwips: textWidth + 1000 },
      ],
    };
    const report = preflight({ profile, analysis: wide, options: withOptions() });

    expect(find(report, 'wide-images')?.level).toBe('attention');
  });

  it('warns about two title pages before they happen', async () => {
    const { profile, analysis } = await analyzed();
    const report = preflight({
      profile,
      analysis,
      options: withOptions({
        bookDetails: { ...EMPTY_BOOK_DETAILS, title: 'A Book' },
        extraSections: { ...NO_EXTRA_SECTIONS, titlePage: true },
        replaceFrontMatter: false,
      }),
    });

    expect(find(report, 'duplicate-front')?.level).toBe('attention');

    // Ticking the replace option clears it.
    const fixed = preflight({
      profile,
      analysis,
      options: withOptions({
        bookDetails: { ...EMPTY_BOOK_DETAILS, title: 'A Book' },
        extraSections: { ...NO_EXTRA_SECTIONS, titlePage: true },
        replaceFrontMatter: true,
      }),
    });
    expect(find(fixed, 'duplicate-front')).toBeUndefined();
  });

  it('explains the blank pages that recto chapter openings create', async () => {
    const { profile, analysis } = await analyzed();
    const report = preflight({ profile, analysis, options: withOptions({ chapterStart: 'oddPage' }) });
    expect(find(report, 'blank-pages')?.level).toBe('check');
  });

  it('never promises the book is approved, and defers to KDP', async () => {
    const { profile, analysis } = await analyzed();
    const report = preflight({ profile, analysis, options: withOptions() });
    const all = report.checks.map((c) => `${c.title} ${c.detail}`).join(' ');

    expect(all).not.toMatch(/approved|guaranteed|100%/i);
    expect(all).toMatch(/Print Previewer/);
  });
});

describe('estimating how long the book will be', () => {
  it('lands in a sane range for a real design', async () => {
    const { profile, analysis } = await analyzed();
    const pages = estimatePageCount(profile, analysis);

    expect(pages).not.toBeNull();
    // The sample is short; the point is that the figure is plausible, not exact.
    expect(pages!).toBeGreaterThan(0);
    expect(pages!).toBeLessThan(10);
  });

  it('scales with the length of the manuscript', async () => {
    const { profile, analysis } = await analyzed();
    const longer = { ...analysis, wordCount: analysis.wordCount * 100 };
    const short = estimatePageCount(profile, analysis)!;
    const long = estimatePageCount(profile, longer)!;

    expect(long).toBeGreaterThan(short * 50);
  });

  it('quotes the same figure when choosing a size as when checking the book', async () => {
    // The size picker works from a design's constants and the report works
    // from a document already read. They must not disagree.
    for (const trim of TRIM_SIZES) {
      for (const look of BOOK_LOOKS) {
        const { profile } = await analyzeReference({
          data: await buildTemplate(trim.id, look.id),
          name: `${trim.id}-${look.id}.docx`,
        });
        const fromDocument = estimatePageCount(profile, {
          wordCount: TYPICAL_NOVEL_WORDS,
        } as ManuscriptAnalysis);
        const fromDesign = estimatePagesForDesign(trim, look, TYPICAL_NOVEL_WORDS);

        expect(fromDesign).toBe(fromDocument);
      }
    }
  });

  it('reads exact line spacing as a height, not as a multiple', () => {
    // Amazon's template sets 9pt type on an exact 220 twip (11pt) line.
    // Treating 220 as 220/240 of a line made lines a third too tight and
    // under-counted a book's length by about a quarter.
    expect(lineSpacingMultiple(220, 'exact', 9)).toBeCloseTo(11 / 9, 5);
    expect(lineSpacingMultiple(240, 'atLeast', 12)).toBeCloseTo(1, 5);
    // `auto` really does count 240ths of a line.
    expect(lineSpacingMultiple(276, 'auto', 11)).toBeCloseTo(1.15, 5);
    expect(lineSpacingMultiple(240, null, 11)).toBeCloseTo(1, 5);
    // Nonsense falls back rather than producing an absurd page count.
    expect(lineSpacingMultiple(null, 'auto', 11)).toBeCloseTo(1.15, 5);
    expect(lineSpacingMultiple(0, 'exact', 11)).toBeCloseTo(1.15, 5);
  });

  it('gives up rather than guessing when the design has no body size', async () => {
    const { profile, analysis } = await analyzed();
    expect(estimatePageCount({ ...profile, bodyFontSizePt: null }, analysis)).toBeNull();
  });
});

describe('KDP-specific checks', () => {
  it('measures how sharply a picture will print and warns when it is placed too large', async () => {
    const { buildDocx, pngHeader } = await import('./helpers/makeDocx.js');
    const { analyzeManuscript } = await import('../src/core/analyze/manuscript.js');
    const paragraphs = [
      { text: 'Chapter One' },
      { text: 'A picture follows.', image: true },
    ];
    // 300 pixels shown at one inch: exactly sharp enough.
    const sharp = await analyzeManuscript({
      data: await buildDocx({ paragraphs, image: true, imageBytes: pngHeader(300, 300) }),
      name: 'sharp.docx',
    });
    expect(sharp.analysis.blocks[1].imageMinDpi).toBe(300);
    // The same 300 pixels stretched over three inches: 100 dots per inch.
    const soft = await analyzeManuscript({
      data: await buildDocx({ paragraphs, image: true, imageBytes: pngHeader(300, 300), imageExtentEmu: 914400 * 3 }),
      name: 'soft.docx',
    });
    expect(soft.analysis.blocks[1].imageMinDpi).toBe(100);

    const { profile } = await analyzed();
    const okReport = preflight({ profile, analysis: sharp.analysis, options: withOptions() });
    expect(find(okReport, 'image-resolution')).toBeUndefined();
    expect(find(okReport, 'image-resolution-ok')?.level).toBe('ready');
    const softReport = preflight({ profile, analysis: soft.analysis, options: withOptions() });
    expect(find(softReport, 'image-resolution')?.level).toBe('attention');
    expect(find(softReport, 'image-resolution')?.detail).toContain('100 dots per inch');
    expect(find(softReport, 'image-resolution')?.examples?.[0].index).toBe(1);
  });

  it('warns when the estimated length is outside what KDP prints', async () => {
    const { profile, analysis } = await analyzed();
    const tiny = { ...analysis, wordCount: 500 };
    expect(find(preflight({ profile, analysis: tiny, options: withOptions() }), 'too-short')?.level).toBe('attention');
    const huge = { ...analysis, wordCount: 600_000 };
    expect(find(preflight({ profile, analysis: huge, options: withOptions() }), 'too-long')?.level).toBe('attention');
    const novel = { ...analysis, wordCount: 80_000 };
    expect(find(preflight({ profile, analysis: novel, options: withOptions() }), 'too-short')).toBeUndefined();
    expect(find(preflight({ profile, analysis: novel, options: withOptions() }), 'too-long')).toBeUndefined();
  });

  it('checks the ISBN typed for the copyright page', async () => {
    const { profile, analysis } = await analyzed();
    const withIsbn = (isbn: string) =>
      preflight({
        profile,
        analysis,
        options: withOptions({
          bookDetails: { ...EMPTY_BOOK_DETAILS, author: 'A. N. Author', copyrightYear: '2026', isbn },
          extraSections: { ...NO_EXTRA_SECTIONS, copyrightPage: true },
        }),
      });
    expect(find(withIsbn('978-0-306-40615-7'), 'isbn-ok')?.level).toBe('ready');
    expect(find(withIsbn('0-306-40615-2'), 'isbn-ok')?.level).toBe('ready');
    expect(find(withIsbn('978-0-306-40615-8'), 'isbn')?.level).toBe('attention');
    expect(find(withIsbn('978-0-306-40615'), 'isbn')?.detail).toContain('wrong number of digits');
    // Nothing to say when no ISBN was given.
    expect(find(withIsbn(''), 'isbn')).toBeUndefined();
    expect(find(withIsbn(''), 'isbn-ok')).toBeUndefined();
  });

  it('spots empty chapters and scene breaks with nothing on one side', async () => {
    const { profile, analysis } = await analyzed();
    const stub = (index: number, role: ManuscriptAnalysis['blocks'][number]['role'], text: string) => ({
      ...analysis.blocks[0],
      index,
      role,
      autoRole: role,
      text,
      preview: text,
      isEmpty: false,
    });
    const odd: ManuscriptAnalysis = {
      ...analysis,
      blocks: [
        stub(0, 'chapterTitle', 'Chapter One'),
        stub(1, 'sceneBreak', '* * *'),
        stub(2, 'body', 'Some text.'),
        stub(3, 'sceneBreak', '* * *'),
        stub(4, 'sceneBreak', '* * *'),
        stub(5, 'body', 'More text.'),
        stub(6, 'chapterTitle', 'Chapter Two'),
        stub(7, 'chapterTitle', 'Chapter Three'),
        stub(8, 'body', 'Text at last.'),
        stub(9, 'sceneBreak', '* * *'),
      ],
    };
    const report = preflight({ profile, analysis: odd, options: withOptions() });
    expect(find(report, 'empty-chapters')?.examples?.map((e) => e.index)).toEqual([6]);
    expect(find(report, 'stranded-scene-breaks')?.examples?.map((e) => e.index)).toEqual([1, 4, 9]);
  });

  it('says what the tidy-up options would touch, and confirms when they are on', async () => {
    const { profile, analysis } = await analyzed();
    const typed: ManuscriptAnalysis = {
      ...analysis,
      straightQuoteCount: 42,
      doubleSpaceCount: 7,
      underlinedRunCount: 3,
    };
    const off = preflight({
      profile,
      analysis: typed,
      options: withOptions({ smartTypography: false, collapseMultipleSpaces: false, underlineToItalic: false }),
    });
    expect(find(off, 'straight-quotes')?.level).toBe('check');
    expect(find(off, 'straight-quotes')?.title).toContain('42 straight quotes');
    expect(find(off, 'double-spaces')?.title).toContain('7 places');
    expect(find(off, 'underlining')?.title).toContain('3 underlined passages');

    const on = preflight({
      profile,
      analysis: typed,
      options: withOptions({ smartTypography: true, collapseMultipleSpaces: true, underlineToItalic: true }),
    });
    expect(find(on, 'straight-quotes')).toBeUndefined();
    expect(find(on, 'straight-quotes-on')?.level).toBe('ready');
    expect(find(on, 'double-spaces-on')?.level).toBe('ready');
    expect(find(on, 'underlining-on')?.level).toBe('ready');

    // A manuscript without those habits gets no note either way.
    const clean = preflight({ profile, analysis, options: withOptions() });
    expect(find(clean, 'straight-quotes')).toBeUndefined();
    expect(find(clean, 'straight-quotes-on')).toBeUndefined();
  });
});
