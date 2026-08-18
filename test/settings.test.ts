import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { attr, child, children } from '../src/core/ooxml/xml.js';
import { analyzeReference } from '../src/core/analyze/reference.js';
import { analyzeManuscript } from '../src/core/analyze/manuscript.js';
import { suggestOptions } from '../src/core/format.js';
import { formatManuscript } from '../src/core/platform/node.js';
import { NO_EXTRA_SECTIONS } from '../src/core/types.js';
import {
  BOOK_TEMPLATE,
  loadDocx,
  openDocx,
  writeDocx,
  type DocSpec,
  type ParaSpec,
} from './helpers/makeDocx.js';

let dir = '';
let counter = 0;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'mf-settings-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const MANUSCRIPT: ParaSpec[] = [
  { runs: [{ text: 'CHAPTER ONE', bold: true }], alignment: 'center' },
  {
    text:
      'The morning came in grey and unhurried, the way mornings did there. Hollis pushed the ' +
      'shop door open with his shoulder and waited for the bell.',
  },
];

async function run(
  reference: DocSpec,
  manuscript: DocSpec,
  options: Record<string, unknown> = {},
): Promise<{ settings: Element | null; styles: Element | null }> {
  const id = counter++;
  const referencePath = await writeDocx(join(dir, `ref${id}.docx`), reference);
  const manuscriptPath = await writeDocx(join(dir, `ms${id}.docx`), manuscript);
  const outputPath = join(dir, `out${id}.docx`);
  await formatManuscript({
    referencePath,
    manuscriptPath,
    outputPath,
    options: { extraSections: NO_EXTRA_SECTIONS, ...options },
  });
  const pkg = await loadDocx(outputPath);
  const settings = await pkg.readXml('word/settings.xml');
  const styles = await pkg.readXml('word/styles.xml');
  return {
    settings: settings?.documentElement ?? null,
    styles: styles?.documentElement ?? null,
  };
}

const names = (root: Element | null): string[] =>
  root ? [...root.childNodes].filter((n) => n.nodeType === 1).map((n) => (n as Element).localName ?? '') : [];

describe('print settings in the output', () => {
  it('asks Word to embed fonts on save, creating a settings part when the design has none', async () => {
    const { settings } = await run(BOOK_TEMPLATE, { paragraphs: MANUSCRIPT });
    expect(settings).not.toBeNull();
    expect(child(settings, 'embedTrueTypeFonts')).not.toBeNull();
    expect(child(settings, 'saveSubsetFonts')).not.toBeNull();
  });

  it('turns hyphenation on for a justified design that forgot it', async () => {
    const id = counter++;
    const referencePath = await writeDocx(join(dir, `ref${id}.docx`), BOOK_TEMPLATE);
    const { profile } = await analyzeReference(await openDocx(referencePath));
    expect(profile.bodyJustified).toBe(true);
    expect(profile.hyphenates).toBe(false);
    expect(suggestOptions(profile).hyphenate).toBe(true);

    const { settings } = await run(BOOK_TEMPLATE, { paragraphs: MANUSCRIPT });
    expect(child(settings, 'autoHyphenation')).not.toBeNull();
    expect(child(settings, 'doNotHyphenateCaps')).not.toBeNull();
    expect(attr(child(settings, 'consecutiveHyphenLimit'), 'val')).toBe('2');
  });

  it('respects the choice not to hyphenate, even when the design did', async () => {
    const hyphenatingDesign: DocSpec = {
      ...BOOK_TEMPLATE,
      settings: '<w:mirrorMargins/><w:autoHyphenation/><w:characterSpacingControl w:val="compressPunctuation"/>',
    };
    const id = counter++;
    const referencePath = await writeDocx(join(dir, `ref${id}.docx`), hyphenatingDesign);
    const { profile } = await analyzeReference(await openDocx(referencePath));
    expect(profile.hyphenates).toBe(true);

    const { settings } = await run(hyphenatingDesign, { paragraphs: MANUSCRIPT }, { hyphenate: false });
    expect(child(settings, 'autoHyphenation')).toBeNull();
    // What the design already had is left alone.
    expect(child(settings, 'mirrorMargins')).not.toBeNull();
  });

  it('keeps the settings elements in schema order', async () => {
    const design: DocSpec = {
      ...BOOK_TEMPLATE,
      settings: '<w:mirrorMargins/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="compressPunctuation"/>',
    };
    const { settings } = await run(design, { paragraphs: MANUSCRIPT }, { hyphenate: true, extraSections: { ...NO_EXTRA_SECTIONS, contents: true } });
    const order = names(settings);
    const rank = (name: string): number => order.indexOf(name);
    expect(rank('embedTrueTypeFonts')).toBeGreaterThanOrEqual(0);
    expect(rank('embedTrueTypeFonts')).toBeLessThan(rank('saveSubsetFonts'));
    expect(rank('saveSubsetFonts')).toBeLessThan(rank('mirrorMargins'));
    expect(rank('defaultTabStop')).toBeLessThan(rank('autoHyphenation'));
    expect(rank('autoHyphenation')).toBeLessThan(rank('consecutiveHyphenLimit'));
    expect(rank('consecutiveHyphenLimit')).toBeLessThan(rank('doNotHyphenateCaps'));
    expect(rank('doNotHyphenateCaps')).toBeLessThan(rank('characterSpacingControl'));
    expect(rank('characterSpacingControl')).toBeLessThan(rank('updateFields'));
  });
});

describe('language', () => {
  it("reads the manuscript's language and writes it into the output defaults", async () => {
    const id = counter++;
    const manuscriptPath = await writeDocx(join(dir, `ms${id}.docx`), {
      paragraphs: MANUSCRIPT,
      language: 'en-GB',
    });
    const { analysis } = await analyzeManuscript(await openDocx(manuscriptPath));
    expect(analysis.language).toBe('en-GB');

    const { styles } = await run(
      { ...BOOK_TEMPLATE, language: 'en-US' },
      { paragraphs: MANUSCRIPT, language: 'en-GB' },
    );
    const defaults = child(child(child(styles, 'docDefaults'), 'rPrDefault'), 'rPr');
    expect(attr(child(defaults, 'lang'), 'val')).toBe('en-GB');
  });

  it("leaves the design's language alone when the manuscript does not say", async () => {
    const { styles } = await run({ ...BOOK_TEMPLATE, language: 'en-US' }, { paragraphs: MANUSCRIPT });
    const defaults = child(child(child(styles, 'docDefaults'), 'rPrDefault'), 'rPr');
    expect(attr(child(defaults, 'lang'), 'val')).toBe('en-US');
    // And no stray language is invented for the paragraph styles.
    for (const style of children(styles, 'style')) {
      expect(child(child(style, 'rPr'), 'lang')).toBeNull();
    }
  });
});
