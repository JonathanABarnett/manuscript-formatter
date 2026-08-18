import { describe, expect, it } from 'vitest';
import { analyzeReference, analyzeManuscript, formatToBuffer } from '../src/core/format.js';
import { DocxPackage } from '../src/core/ooxml/package.js';
import { attr, child, children, descendants, textOf } from '../src/core/ooxml/xml.js';
import { BOOK_LOOKS, GUTTER_IN, TRIM_SIZES } from '../src/core/templates/design.js';
import { buildSampleManuscript, buildTemplate } from '../src/core/templates/generate.js';
import type { DocxInput } from '../src/core/types.js';

/**
 * The built-in designs travel the same code path as a template an author
 * supplies, so the analyzer has to be able to read what the generator writes.
 * If these drift apart, Quick Start silently produces the wrong book.
 */

const asInput = (data: Uint8Array, name: string): DocxInput => ({ data, name });

const TWIPS_PER_INCH = 1440;

describe('the built-in book designs', () => {
  it('offers five sizes and four looks, with a recommended default', () => {
    expect(TRIM_SIZES).toHaveLength(5);
    expect(BOOK_LOOKS).toHaveLength(4);
    expect(TRIM_SIZES.filter((t) => t.recommended)).toHaveLength(1);
    expect(TRIM_SIZES.find((t) => t.recommended)?.id).toBe('6x9');
  });

  for (const trim of TRIM_SIZES) {
    for (const look of BOOK_LOOKS) {
      it(`builds a ${trim.id} ${look.id} design the analyzer can read`, async () => {
        const data = await buildTemplate(trim.id, look.id);
        const { profile } = await analyzeReference(asInput(data, `${trim.id}-${look.id}.docx`));

        // Page geometry survives the round trip exactly.
        expect(profile.pageSetup.widthTwips).toBe(Math.round(trim.widthIn * TWIPS_PER_INCH));
        expect(profile.pageSetup.heightTwips).toBe(Math.round(trim.heightIn * TWIPS_PER_INCH));
        expect(profile.pageSetup.margins.gutter).toBe(GUTTER_IN * TWIPS_PER_INCH);
        expect(profile.pageSetup.mirrorMargins).toBe(true);

        // Every role the composer needs resolves to a real style.
        expect(profile.roleStyles.body).toBe('BodyText');
        expect(profile.roleStyles.bodyFirst).toBe('FirstParagraph');
        expect(profile.roleStyles.chapterTitle).toBe('ChapterTitle');
        expect(profile.roleStyles.sceneBreak).toBe('SceneBreak');
        expect(profile.roleStyles.blockQuote).toBe('BlockQuote');
        expect(profile.roleStyles.copyright).toBe('CopyrightPage');
        expect(profile.roleStyles.subheading).toBe('Subhead');
        expect(profile.roleStyles.frontMatterTitle).toBe('BookTitle');

        // The chapter sink is read back exactly as the design set it.
        expect(profile.chapterTitleBlanksBefore).toBe(look.chapterBlanksBefore);
        expect(profile.chapterTitleBlanksAfter).toBe(look.chapterBlanksAfter);

        // Page numbers are present, which the preflight check looks for.
        expect(profile.hasFooters).toBe(true);
        expect(profile.hasPageNumbers).toBe(true);

        expect(profile.bodyFontName).toBe(look.bodyFont);
        expect(profile.bodyFontSizePt).toBe(look.bodySizePt);
        expect(profile.warnings.filter((w) => /no chapter-title style/i.test(w))).toHaveLength(0);
      });
    }
  }

  it('ships no placeholder prose into a formatted book', async () => {
    const template = await buildTemplate('6x9', 'classic');
    const manuscript = await buildSampleManuscript();
    const out = await formatToBuffer({
      reference: asInput(template, 'design.docx'),
      manuscript: asInput(manuscript, 'sample.docx'),
      options: {},
    });

    const pkg = await DocxPackage.fromBuffer(Buffer.from(out.data), 'out.docx');
    const doc = await pkg.readXml(pkg.documentPath);
    const text = textOf(child(doc?.documentElement ?? null, 'body'));

    // None of the design's own specimen sentences may reach the finished book.
    expect(text).not.toMatch(/specimen/i);
    expect(text).not.toMatch(/replaces it/i);
    expect(text).not.toMatch(/even colour from top to bottom/i);
    expect(text).toContain('Hollis');
  });
});

describe('the sample manuscript', () => {
  it('looks like a manuscript an author would actually send', async () => {
    const data = await buildSampleManuscript();
    const { analysis } = await analyzeManuscript(asInput(data, 'sample.docx'));

    expect(analysis.chapterCount).toBe(2);
    expect(analysis.sceneBreakCount).toBe(2);
    expect(analysis.wordCount).toBeGreaterThan(200);

    const roles = analysis.blocks.map((b) => b.role);
    expect(roles).toContain('frontMatterTitle');
    expect(roles).toContain('copyright');
    expect(roles).toContain('chapterTitle');
    expect(roles).toContain('bodyFirst');
  });

  it('formats cleanly against every built-in design', async () => {
    const manuscript = await buildSampleManuscript();

    for (const trim of TRIM_SIZES) {
      const template = await buildTemplate(trim.id, 'classic');
      const out = await formatToBuffer({
        reference: asInput(template, 'design.docx'),
        manuscript: asInput(manuscript, 'sample.docx'),
        options: {},
      });

      expect(out.stats.chapters).toBe(2);
      expect(out.stats.sceneBreaks).toBe(2);
      expect(out.warnings.filter((w) => /could not|failed|dropped/i.test(w))).toHaveLength(0);

      const pkg = await DocxPackage.fromBuffer(Buffer.from(out.data), 'out.docx');
      const doc = await pkg.readXml(pkg.documentPath);
      const body = child(doc?.documentElement ?? null, 'body');

      // The output keeps the chosen trim size and its page-number footer.
      const pgSz = child(child(body, 'sectPr'), 'pgSz');
      expect(attr(pgSz, 'w')).toBe(String(Math.round(trim.widthIn * TWIPS_PER_INCH)));
      const rels = await pkg.relsFor(pkg.documentPath);
      expect(rels.all().some((r) => r.type.endsWith('/footer'))).toBe(true);

      // Chapter titles land on the design's own chapter style.
      const titles = children(body, 'p').filter((p) =>
        ['CHAPTER ONE', 'CHAPTER TWO'].includes(textOf(p).trim()),
      );
      expect(titles).toHaveLength(2);
      for (const t of titles) {
        expect(attr(child(child(t, 'pPr'), 'pStyle'), 'val')).toBe('ChapterTitle');
      }

      // No style id may dangle: Word would silently fall back to Normal.
      const styles = await pkg.readXml('word/styles.xml');
      const known = new Set(
        children(styles?.documentElement ?? null, 'style').map((el) => attr(el, 'styleId')),
      );
      for (const el of descendants(body, 'pStyle')) {
        expect(known.has(attr(el, 'val'))).toBe(true);
      }
    }
  });
});
