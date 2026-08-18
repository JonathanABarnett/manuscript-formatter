import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/ooxml/package.js';
import { attr, child, children, descendants, textOf } from '../src/core/ooxml/xml.js';
import { analyzeManuscript } from '../src/core/analyze/manuscript.js';
import { formatManuscript, suggestOutputPath } from '../src/core/platform/node.js';
import { BOOK_TEMPLATE, loadDocx, openDocx, type ParaSpec, writeDocx } from './helpers/makeDocx.js';
import { NO_EXTRA_SECTIONS } from '../src/core/types.js';

const loadManuscript = async (path: string) => analyzeManuscript(await openDocx(path));

let dir = '';
let counter = 0;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'mf-fmt-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const PROSE_1 =
  'The morning came in grey and unhurried, the way mornings did there. Hollis ' +
  'pushed the shop door open with his shoulder and waited for the bell.';
const PROSE_2 =
  'The corridor smelled of wet wool and yesterday rain. A notice board at the ' +
  'far end carried three curling announcements, none of them newer than spring.';
const PROSE_3 =
  'Outside, past the glass, the square was emptying. Wind moved scraps of paper ' +
  'in slow circles, and the light had gone thin the way it does in November.';

/** Opts out of the contents page that is offered by default, so these tests
 *  assert on the manuscript's own paragraphs alone. */
const PLAIN = { extraSections: NO_EXTRA_SECTIONS };

const RAW_MANUSCRIPT: ParaSpec[] = [
  { runs: [{ text: 'The Sample Book', bold: true, fontSizePt: 20 }], alignment: 'center' },
  { text: 'by A. N. Author', alignment: 'center' },
  { text: '', leadingPageBreak: true },
  { runs: [{ text: 'CHAPTER ONE', bold: true, fontSizePt: 16, color: 'FF0000' }], alignment: 'center' },
  { text: '' },
  {
    runs: [
      { text: '\t' + PROSE_1 + ' ', fontName: 'Courier New', fontSizePt: 12 },
      { text: 'the small bell', italic: true, fontName: 'Courier New', fontSizePt: 12 },
      { text: ' had been there for forty years.', fontName: 'Courier New', fontSizePt: 12 },
    ],
  },
  { text: `\t${PROSE_2}` },
  { text: '#', alignment: 'center' },
  { text: `\t${PROSE_3}` },
  { text: '', leadingPageBreak: true },
  { runs: [{ text: 'CHAPTER TWO', bold: true, fontSizePt: 16 }], alignment: 'center' },
  { text: `\t${PROSE_1}` },
];

interface OutParagraph {
  styleId: string | null;
  text: string;
  pageBreakBefore: boolean;
  sectionType: string | null;
  alignment: string | null;
  firstLineIndent: string | null;
}

interface OutDoc {
  paragraphs: OutParagraph[];
  bodySectPr: Element | null;
  pkg: DocxPackage;
}

async function readOutput(path: string): Promise<OutDoc> {
  const pkg = await loadDocx(path);
  const doc = await pkg.readXml(pkg.documentPath);
  const body = child(doc?.documentElement ?? null, 'body');
  if (!body) throw new Error('output has no body');
  const paragraphs = children(body, 'p').map((p) => {
    const pPr = child(p, 'pPr');
    const sectPr = child(pPr, 'sectPr');
    const ind = child(pPr, 'ind');
    return {
      styleId: attr(child(pPr, 'pStyle'), 'val'),
      text: textOf(p).trim(),
      pageBreakBefore: child(pPr, 'pageBreakBefore') !== null,
      sectionType: sectPr ? attr(child(sectPr, 'type'), 'val') : null,
      alignment: attr(child(pPr, 'jc'), 'val'),
      firstLineIndent: attr(ind, 'firstLine'),
    };
  });
  return { paragraphs, bodySectPr: child(body, 'sectPr'), pkg };
}

async function setup(manuscript: ParaSpec[] = RAW_MANUSCRIPT): Promise<{
  reference: string;
  manuscript: string;
  output: string;
}> {
  const id = counter++;
  return {
    reference: await writeDocx(join(dir, `ref${id}.docx`), BOOK_TEMPLATE),
    manuscript: await writeDocx(join(dir, `ms${id}.docx`), { paragraphs: manuscript }),
    output: join(dir, `out${id}.docx`),
  };
}

describe('formatting a manuscript against a reference', () => {
  it('leaves both source documents byte-for-byte unchanged', async () => {
    const paths = await setup();
    const before = await Promise.all([readFile(paths.reference), readFile(paths.manuscript)]);

    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
    });

    const after = await Promise.all([readFile(paths.reference), readFile(paths.manuscript)]);
    expect(after[0].equals(before[0])).toBe(true);
    expect(after[1].equals(before[1])).toBe(true);
  });

  it('carries the reference page geometry into the output', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
    });

    const out = await readOutput(paths.output);
    const pgSz = child(out.bodySectPr, 'pgSz');
    const pgMar = child(out.bodySectPr, 'pgMar');
    expect(attr(pgSz, 'w')).toBe('7920');
    expect(attr(pgSz, 'h')).toBe('12240');
    expect(attr(pgMar, 'gutter')).toBe('360');
    expect(child(out.bodySectPr, 'titlePg')).not.toBeNull();
  });

  it('maps every paragraph onto a style from the reference', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: PLAIN,
    });

    const out = await readOutput(paths.output);
    const styles = out.paragraphs.map((p) => p.styleId);
    expect(styles).toEqual([
      'BookTitle',
      'FirstParagraph',
      'ChapterTitle',
      'FirstParagraph',
      'BodyText',
      'SceneBreak',
      'FirstParagraph',
      'ChapterTitle',
      'FirstParagraph',
    ]);
  });

  it('preserves the manuscript text exactly, including its emphasis', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: PLAIN,
    });

    const { analysis } = await loadManuscript(paths.manuscript);
    const expected = analysis.blocks.filter((b) => !b.isEmpty).map((b) => b.text.trim());
    const out = await readOutput(paths.output);
    expect(out.paragraphs.map((p) => p.text)).toEqual(expected);

    const doc = await out.pkg.readXml(out.pkg.documentPath);
    const italicRuns = descendants(doc?.documentElement ?? null, 'i');
    expect(italicRuns).toHaveLength(1);
    expect(textOf(italicRuns[0].parentNode?.parentNode ?? null)).toContain('the small bell');
  });

  it('sets underlined passages in italics, unless told to keep the underlining', async () => {
    const underlined: ParaSpec[] = [
      { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center' },
      {
        runs: [
          { text: `${PROSE_1} She had read ` },
          { text: 'Middlemarch', underline: true },
          { text: ' twice that year.' },
        ],
      },
    ];
    const paths = await setup(underlined);
    const { analysis } = await loadManuscript(paths.manuscript);
    expect(analysis.underlinedRunCount).toBe(1);

    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: PLAIN,
    });
    let out = await readOutput(paths.output);
    let root = (await out.pkg.readXml(out.pkg.documentPath))?.documentElement ?? null;
    expect(descendants(root, 'u')).toHaveLength(0);
    const italic = descendants(root, 'i');
    expect(italic).toHaveLength(1);
    expect(textOf(italic[0].parentNode?.parentNode ?? null)).toBe('Middlemarch');

    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { ...PLAIN, underlineToItalic: false },
    });
    out = await readOutput(paths.output);
    root = (await out.pkg.readXml(out.pkg.documentPath))?.documentElement ?? null;
    expect(descendants(root, 'u')).toHaveLength(1);
    expect(descendants(root, 'i')).toHaveLength(0);
  });

  it("drops the manuscript's own fonts, sizes and colours", async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
    });

    const out = await readOutput(paths.output);
    const doc = await out.pkg.readXml(out.pkg.documentPath);
    const root = doc?.documentElement ?? null;
    expect(descendants(root, 'rFonts')).toHaveLength(0);
    expect(descendants(root, 'sz')).toHaveLength(0);
    expect(descendants(root, 'color')).toHaveLength(0);
  });

  it('strips the tab characters used as manual indents', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
    });

    const out = await readOutput(paths.output);
    const doc = await out.pkg.readXml(out.pkg.documentPath);
    for (const t of descendants(doc?.documentElement ?? null, 't')) {
      expect(t.textContent ?? '').not.toMatch(/^\t/);
    }
  });

  it('starts each chapter on a new page but not the first thing in the file', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { ...PLAIN, chapterStart: 'newPage' },
    });

    const out = await readOutput(paths.output);
    const chapters = out.paragraphs.filter((p) => p.styleId === 'ChapterTitle');
    expect(chapters).toHaveLength(2);
    expect(chapters.every((c) => c.pageBreakBefore)).toBe(true);
    expect(out.paragraphs[0].pageBreakBefore).toBe(false);
  });

  it('opens chapters on a recto page when asked', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { chapterStart: 'oddPage' },
    });

    const out = await readOutput(paths.output);
    const breaks = out.paragraphs.filter((p) => p.sectionType === 'oddPage');
    expect(breaks.length).toBeGreaterThanOrEqual(1);

    // Exactly one section may restart page numbering, and it must be the
    // first: otherwise the last chapter would begin again at page 1.
    const doc = await out.pkg.readXml(out.pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const starts = descendants(body, 'pgNumType').map((el) => attr(el, 'start'));
    expect(starts.filter((s) => s !== null)).toEqual(['1']);
    expect(starts[0]).toBe('1');
  });

  it('keeps the page-numbering restart when there is only one section', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { chapterStart: 'newPage' },
    });

    const out = await readOutput(paths.output);
    expect(attr(child(out.bodySectPr, 'pgNumType'), 'start')).toBe('1');
  });

  it('can leave the front matter out', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { includeFrontMatter: false },
    });

    const out = await readOutput(paths.output);
    expect(out.paragraphs[0].styleId).toBe('ChapterTitle');
    expect(out.paragraphs.some((p) => p.text === 'The Sample Book')).toBe(false);
  });

  it('honours a per-paragraph role override', async () => {
    const paths = await setup();
    const { analysis } = await loadManuscript(paths.manuscript);
    const chapterTwo = analysis.blocks.find((b) => b.preview === 'CHAPTER TWO');
    expect(chapterTwo).toBeDefined();

    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { roleOverrides: { [chapterTwo!.index]: 'subheading' } },
    });

    const out = await readOutput(paths.output);
    const row = out.paragraphs.find((p) => p.text === 'CHAPTER TWO');
    expect(row?.styleId).toBe('Subhead');
    expect(row?.pageBreakBefore).toBe(false);
  });

  it('honours a role-to-style override', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { roleStyles: { sceneBreak: 'BlockQuote' } },
    });

    const out = await readOutput(paths.output);
    expect(out.paragraphs.find((p) => p.text === '#')?.styleId).toBe('BlockQuote');
  });

  it('can substitute the ornament used for scene breaks', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { sceneBreakText: '* * *' },
    });

    const out = await readOutput(paths.output);
    expect(out.paragraphs.some((p) => p.styleId === 'SceneBreak' && p.text === '* * *')).toBe(true);
    expect(out.paragraphs.some((p) => p.text === '#')).toBe(false);
  });

  it('reports what it did', async () => {
    const paths = await setup();
    const result = await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: PLAIN,
    });

    expect(result.stats.chapters).toBe(2);
    expect(result.stats.sceneBreaks).toBe(1);
    expect(result.stats.paragraphsWritten).toBe(9);
    expect(result.stats.blanksRemoved).toBeGreaterThan(0);
    expect(result.outputPath).toBe(paths.output);
  });

  it('refuses to write over either source document', async () => {
    const paths = await setup();
    await expect(
      formatManuscript({
        referencePath: paths.reference,
        manuscriptPath: paths.manuscript,
        outputPath: paths.manuscript,
      }),
    ).rejects.toThrow(/must be different/);
    await expect(
      formatManuscript({
        referencePath: paths.reference,
        manuscriptPath: paths.manuscript,
        outputPath: paths.reference,
      }),
    ).rejects.toThrow(/must be different/);
  });

  it('suggests an output name that does not collide', async () => {
    const paths = await setup();
    const suggested = await suggestOutputPath(paths.manuscript);
    expect(suggested).toMatch(/\(formatted\)\.docx$/);
    expect(suggested).not.toBe(paths.manuscript);
  });

  it('carries numbered lists across with their numbering intact', async () => {
    const id = counter++;
    const reference = await writeDocx(join(dir, `nref${id}.docx`), BOOK_TEMPLATE);
    const manuscript = await writeDocx(join(dir, `nms${id}.docx`), {
      numbering: true,
      paragraphs: [
        { text: PROSE_1 },
        { text: 'First point', numId: 1 },
        { text: 'Second point', numId: 1 },
        { text: PROSE_2 },
      ],
    });
    const output = join(dir, `nout${id}.docx`);
    await formatManuscript({ referencePath: reference, manuscriptPath: manuscript, outputPath: output });

    const pkg = await loadDocx(output);
    const doc = await pkg.readXml(pkg.documentPath);
    const numIds = descendants(child(doc?.documentElement ?? null, 'body'), 'numId').map((el) =>
      attr(el, 'val'),
    );
    expect(numIds).toHaveLength(2);
    expect(numIds[0]).toBe(numIds[1]);

    const numbering = await pkg.readXml('word/numbering.xml');
    expect(numbering).not.toBeNull();
    const defined = children(numbering?.documentElement ?? null, 'num').map((el) => attr(el, 'numId'));
    expect(defined).toContain(numIds[0]);
    expect(children(numbering?.documentElement ?? null, 'abstractNum').length).toBeGreaterThan(0);
  });

  it('produces a package Word can open', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
    });

    const pkg = await loadDocx(paths.output);
    expect(pkg.has('[Content_Types].xml')).toBe(true);
    expect(pkg.has('word/styles.xml')).toBe(true);
    // Every part named in the document relationships must actually be present.
    const rels = await pkg.relsFor(pkg.documentPath);
    for (const rel of rels.all()) {
      if (rel.targetMode === 'External') continue;
      expect(pkg.has(pkg.resolveTarget(pkg.documentPath, rel.target))).toBe(true);
    }
    await expect(pkg.readXml(pkg.documentPath)).resolves.not.toBeNull();
  });
});

describe('chapter-opening pages without a running head', () => {
  const HEADED_DESIGN = { ...BOOK_TEMPLATE, headerText: 'BOOK TITLE', footerText: 'page' };
  const THREE_CHAPTERS: ParaSpec[] = [
    { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center' },
    { text: PROSE_1 },
    { runs: [{ text: 'CHAPTER TWO', bold: true }], alignment: 'center' },
    { text: PROSE_2 },
    { runs: [{ text: 'CHAPTER THREE', bold: true }], alignment: 'center' },
    { text: PROSE_3 },
  ];

  it('is suggested when the design has headers, and gives each chapter a section with a distinct first page', async () => {
    const id = counter++;
    const reference = await writeDocx(join(dir, `ref${id}.docx`), HEADED_DESIGN);
    const manuscript = await writeDocx(join(dir, `ms${id}.docx`), { paragraphs: THREE_CHAPTERS });
    const output = join(dir, `out${id}.docx`);
    const { analyzeReference, suggestOptions } = await import('../src/core/format.js');
    const { profile } = await analyzeReference(await openDocx(reference));
    expect(suggestOptions(profile).chapterOpenerNoHeader).toBe(true);

    await formatManuscript({
      referencePath: reference,
      manuscriptPath: manuscript,
      outputPath: output,
      options: { ...PLAIN, chapterStart: 'newPage', chapterOpenerNoHeader: true },
    });
    const out = await readOutput(output);
    const chapters = out.paragraphs.filter((p) => p.styleId === 'ChapterTitle');
    expect(chapters).toHaveLength(3);
    // Sections do the page-starting, not page breaks.
    expect(chapters.some((c) => c.pageBreakBefore)).toBe(false);
    expect(out.paragraphs.filter((p) => p.sectionType === 'nextPage')).toHaveLength(2);

    const doc = await out.pkg.readXml(out.pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const sectPrs = descendants(body, 'sectPr');
    expect(sectPrs).toHaveLength(3);
    for (const sectPr of sectPrs) {
      expect(child(sectPr, 'titlePg')).not.toBeNull();
      const footers = children(sectPr, 'footerReference').map((f) => attr(f, 'type'));
      // The page number keeps printing on the opener; the header does not.
      expect(footers).toContain('first');
      expect(children(sectPr, 'headerReference').map((h) => attr(h, 'type'))).not.toContain('first');
      // titlePg keeps its place in the schema sequence.
      const names = [...sectPr.childNodes].filter((n) => n.nodeType === 1).map((n) => (n as Element).localName);
      expect(names.indexOf('titlePg')).toBeGreaterThan(names.indexOf('pgMar'));
    }
    // Only the first section restarts the numbering.
    const starts = descendants(body, 'pgNumType').map((el) => attr(el, 'start'));
    expect(starts.filter((s) => s !== null)).toEqual(['1']);
  });

  it('leaves the old page-break behaviour alone when turned off', async () => {
    const id = counter++;
    const reference = await writeDocx(join(dir, `ref${id}.docx`), HEADED_DESIGN);
    const manuscript = await writeDocx(join(dir, `ms${id}.docx`), { paragraphs: THREE_CHAPTERS });
    const output = join(dir, `out${id}.docx`);
    await formatManuscript({
      referencePath: reference,
      manuscriptPath: manuscript,
      outputPath: output,
      options: { ...PLAIN, chapterStart: 'newPage', chapterOpenerNoHeader: false },
    });
    const out = await readOutput(output);
    const chapters = out.paragraphs.filter((p) => p.styleId === 'ChapterTitle');
    expect(chapters.slice(1).every((c) => c.pageBreakBefore)).toBe(true);
    expect(out.paragraphs.filter((p) => p.sectionType !== null)).toHaveLength(0);
  });
});

describe('opening a chapter with small capitals', () => {
  it('sets the first words of the first paragraph in small caps, splitting a run where it must', async () => {
    const paths = await setup([
      { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center' },
      { runs: [{ text: '“Well,” said ' }, { text: 'Hollis', italic: true }, { text: ', pushing the door.' }] },
      { text: PROSE_2 },
      { text: '#', alignment: 'center' },
      { text: PROSE_3 },
      { runs: [{ text: 'CHAPTER TWO', bold: true }], alignment: 'center' },
      { text: `${PROSE_1}` },
    ]);
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: { ...PLAIN, leadInSmallCaps: true },
    });
    const out = await readOutput(paths.output);
    const doc = await out.pkg.readXml(out.pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const paragraphs = children(body, 'p');
    const smallCapText = (p: Element) =>
      descendants(p, 'r')
        .filter((r) => child(child(r, 'rPr'), 'smallCaps') !== null)
        .map((r) => textOf(r))
        .join('');
    // Cut at the comma, carrying the closing quote; the rest of the run is untouched.
    expect(smallCapText(paragraphs[1])).toBe('“Well,”');
    expect(textOf(paragraphs[1])).toBe('“Well,” said Hollis, pushing the door.');
    // Only chapter openers, not the paragraph after a scene break.
    expect(smallCapText(paragraphs[4])).toBe('');
    // Up to four words when no punctuation intervenes.
    expect(smallCapText(paragraphs[6])).toBe('The morning came in');
    expect(textOf(paragraphs[6])).toBe(PROSE_1);
  });

  it('does nothing unless asked', async () => {
    const paths = await setup();
    await formatManuscript({
      referencePath: paths.reference,
      manuscriptPath: paths.manuscript,
      outputPath: paths.output,
      options: PLAIN,
    });
    const out = await readOutput(paths.output);
    const doc = await out.pkg.readXml(out.pkg.documentPath);
    expect(descendants(doc?.documentElement ?? null, 'smallCaps')).toHaveLength(0);
  });
});
