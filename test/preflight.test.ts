import { describe, expect, it } from 'vitest';
import { analyzeDocuments, analyzeReference } from '../src/core/format.js';
import { TYPICAL_NOVEL_WORDS } from '../src/core/pageEstimate.js';
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

  it('gives up rather than guessing when the design has no body size', async () => {
    const { profile, analysis } = await analyzed();
    expect(estimatePageCount({ ...profile, bodyFontSizePt: null }, analysis)).toBeNull();
  });
});
